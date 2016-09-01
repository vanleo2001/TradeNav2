VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7L.ocx"
Begin VB.Form frmDropTree 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrDropTree 
      Interval        =   50
      Left            =   3960
      Top             =   240
   End
   Begin VSFlex7LCtl.VSFlexGrid fgDropTree 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      _cx             =   6165
      _cy             =   4683
      _ConvInfo       =   1
      Appearance      =   0
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
      ScrollTrack     =   -1  'True
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
Attribute VB_Name = "frmDropTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmDropTree.frm
'' Description: Form that allows for a drop-down tree instead of a drop-down
''              list box
''
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date      Author      Description
'' 03/13/02  D Jarmuth   Created
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Const AUTOSEARCH_DELAY = 2 'seconds

#If 0 Then
Private Type mPrivate
    bMousePressed As Boolean
    bScrolling As Boolean
    bActivating As Boolean
    bLoaded As Boolean
    ctlParent As Control
    lpParent As POINTAPI
    strNameSearch As String
    lSelectedRow As Long
    SelectedItem As cgdDropTreeItem
    dWhenLastHidden As Double
    bHasPicture As Boolean
End Type
Private m As mPrivate

Public Property Get SelectedRow() As Long
    SelectedRow = m.lSelectedRow
End Property
Public Property Let SelectedRow(ByVal lValue As Long)
    m.lSelectedRow = lValue
End Property
Public Property Get SelectedItem() As cgdDropTreeItem
    Set SelectedItem = m.SelectedItem
End Property
Public Property Get Loaded() As Boolean
    Loaded = m.bLoaded
End Property
Public Property Get WhenLastHidden() As Double
    WhenLastHidden = m.dWhenLastHidden
End Property
Public Property Get Activating() As Boolean
    Activating = m.bActivating
End Property
Public Property Get HasPicture() As Boolean
    HasPicture = m.bHasPicture
End Property
Friend Property Let HasPicture(ByVal bValue As Boolean)
    m.bHasPicture = bValue
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HideMe
'' Description: Hide the form if it is currently visible
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub HideMe()
On Error Resume Next

    tmrDropTree.Enabled = False
    If m.bLoaded Then
        If Me.Visible Then Me.Hide
    End If
    Set m.ctlParent = Nothing
    m.bMousePressed = False
    m.dWhenLastHidden = GetTickCount

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgDropTree_AfterCollapse
'' Description: When the user expands or collapses a node, start the name
''              search over
'' Inputs:      Row expanded or collapsed, Whether it was expanded or collapsed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgDropTree_AfterCollapse(ByVal Row As Long, ByVal State As Integer)
On Error Resume Next

    m.strNameSearch = ""

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgDropTree_BeforeScroll
'' Description: Before the user starts to scroll, turn the scrolling variable
''              on
'' Inputs:      Old Top Row, Old Left Col, New Top Row, New Left Col, Cancel
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgDropTree_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)
    
    ' If the Mouse Pressed is true, that means that the timer went off before
    ' the BeforeScroll event took place, so cancel out of the BeforeScroll
    If m.bMousePressed Then
        Cancel = True
        
    ' Otherwise turn the Scrolling variable on to let the timer turn it off
    Else
        m.bScrolling = True
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgDropTree_BeforeScrollTip
'' Description: Before the user starts dragging the scroll bar, turn the
''              scrolling variable on
'' Inputs:      Row to get the scroll tip from
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgDropTree_BeforeScrollTip(ByVal Row As Long)
    
    ' Need to also turn the Scrolling variable on here so that we catch the
    ' case where the user is dragging the scroll bar because this fires instead
    ' of the BeforeScroll event in that case
    m.bScrolling = True

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgDropTree_Click
'' Description: When the user clicks on a node that is not on the outside
''              level of the tree, select it and hide the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgDropTree_Click()

    ' If the mouse row is no longer valid, exit
    If fgDropTree.MouseRow < 0 Then Exit Sub

    With fgDropTree
        ' Set the current row to the current mouse row
        .Row = .MouseRow
        .RowSel = .Row
        
        ' Only select if not on the outside level of the tree
        'If .RowOutlineLevel(.Row) > 0 Then
        If .TextMatrix(.Row, 3) = True Then
            m.lSelectedRow = .Row
            ItemSelected
            HideMe
        End If
    End With

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgDropTree_DblClick
'' Description: If the user double clicks on the tree, toggle the expansion
''              of the node that they clicked on
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgDropTree_DblClick()
On Error Resume Next

    Dim lRow As Long                    ' Row that the user clicked on
    
    With fgDropTree
        ' Save the current mouse row
        lRow = .MouseRow
        
        ' Toggle the expansion of the current mouse row
        If .IsCollapsed(lRow) = flexOutlineCollapsed Then
            .IsCollapsed(lRow) = flexOutlineExpanded
        Else
            .IsCollapsed(lRow) = flexOutlineCollapsed
        End If
    End With

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgDropTree_KeyDown
'' Description: Handle some extra keystrokes from the user
'' Inputs:      Code of the Key pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgDropTree_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    Dim lRow As Long                    ' Currently selected row
    Dim lFirstChild As Long             ' First child of the current row
    
    With fgDropTree
        ' Save the current row
        lRow = .Row
        
        Select Case KeyCode
            ' Go to top-most node of the tree
            Case vbKeyHome
                KeyCode = 0
                ShowRow 0
                m.strNameSearch = ""
            
            ' Go to bottom (last non-hidden row)
            Case vbKeyEnd
                KeyCode = 0
                For lRow = .Rows - 1 To 1 Step -1
                    If Not .RowHidden(lRow) Then
                        ShowRow lRow
                        Exit For
                    End If
                Next
                m.strNameSearch = ""
            
            ' Expand, or move to next row (only if child or sibling)
            Case vbKeyRight
                KeyCode = 0
                lFirstChild = .GetNodeRow(lRow, flexNTFirstChild)
                If lFirstChild > -1 Then
                    .IsCollapsed(lRow) = flexOutlineExpanded
                    ShowRow lFirstChild
                ElseIf lRow + 1 < .Rows - 1 Then
                    ShowRow lRow + 1
                End If
                m.strNameSearch = ""
                
            ' Collapse, or move to parent
            Case vbKeyLeft
                KeyCode = 0
                lFirstChild = .GetNodeRow(lRow, flexNTFirstChild)
                If lFirstChild > -1 And .IsCollapsed(lRow) = flexOutlineExpanded Then
                    .IsCollapsed(lRow) = flexOutlineCollapsed
                Else
                    lRow = .GetNodeRow(lRow, flexNTParent)
                    If lRow > -1 Then
                        ShowRow lRow
                    End If
                End If
                m.strNameSearch = ""
                
            ' Clear name search on these keys
            Case vbKeyUp, vbKeyDown, vbKeyPageUp, vbKeyPageDown
                m.strNameSearch = ""
        
        End Select
        
        If KeyCode <> vbKeyShift And KeyCode <> vbKeyControl Then .Col = 0
    End With

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgDropTree_KeyPress
'' Description: Handle user key presses
'' Inputs:      Key pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgDropTree_KeyPress(KeyAscii As Integer)
On Error Resume Next

    Static lLastKeyTime As Long         ' Time that the last key was pressed
    Dim lRow As Long                    ' Current row
    Dim strChk As String                ' String to check against name search
    
    With fgDropTree
        ' Save the current row
        lRow = .Row
        
        Select Case KeyAscii
            Case 13 ' Enter Key
                'If .RowOutlineLevel(lRow) > 0 Then
                If .TextMatrix(lRow, 3) = True Then
                    m.lSelectedRow = lRow
                    ItemSelected
                    HideMe
                End If
                
            Case 27: ' Esc Key
                m.lSelectedRow = -1&
                HideMe
            
            Case 32 To 127 ' Keyboard characters
                ' if long enough since last key, start over
                If ElapsedSeconds(lLastKeyTime) > AUTOSEARCH_DELAY _
                    Or Len(m.strNameSearch) = 0 Then
                        m.strNameSearch = UCase(Chr(KeyAscii))
                        lRow = lRow + 1
                Else
                    m.strNameSearch = m.strNameSearch & UCase(Chr(KeyAscii))
                End If
                
                Do While lRow < .Rows
                    If Not .RowHidden(lRow) Then
                        strChk = UCase(Trim(.TextMatrix(lRow, 0)))
                        If m.strNameSearch = Left(strChk, Len(m.strNameSearch)) Then
                            ShowRow lRow
                            Exit Do
                        End If
                    End If
                    lRow = lRow + 1
                Loop
                lLastKeyTime = GetTickCount()
                
            Case Else
                m.strNameSearch = ""
        End Select
        
        .Col = 0
    End With

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgDropTree_LostFocus
'' Description: If the tree loses focus, try to get the focus back
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgDropTree_LostFocus()
On Error Resume Next

    fgDropTree.SetFocus

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgDropTree_MouseDown
'' Description: When the user presses the mouse on the grid, turn the Mouse
''              Pressed variable off
'' Inputs:      Button Pressed, Shift/Ctrl/Alt status, Location of click
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgDropTree_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' Turn the mouse pressed variable off to let the timer know we are still
    ' on this form
    m.bMousePressed = False

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgDropTree_MouseUp
'' Description: When the user presses the mouse on the grid, turn the Mouse
''              Pressed variable off
'' Inputs:      Button Pressed, Shift/Ctrl/Alt status, Location of click
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgDropTree_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' Turn the mouse pressed variable off to let the timer know we are still
    ' on this form
    m.bMousePressed = False

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Deactivate
'' Description: When the form gets deactivated, hide it
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Deactivate()

    HideMe

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: When the form is loaded, set the loaded flag
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()

    m.bLoaded = True
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_LostFocus
'' Description: If the form loses focus, hide it
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_LostFocus()

    HideMe

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: When the form is unloaded, disable the timer and clean up
'' Inputs:      Whether or not to cancel the unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)

    tmrDropTree.Enabled = False
    Set m.ctlParent = Nothing
    m.bLoaded = False

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tmrDropTree_Timer
'' Description: When the timer goes off, check to see if the parent has moved
''              and if so, hide the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrDropTree_Timer()
On Error Resume Next
    
    Dim hActive&, dStart#
        
    ' See if this form is no longer active
    hActive = GetActiveWindow
    If hActive <> Me.hWnd And hActive <> 0 Then
        m.lSelectedRow = -1&
        HideMe
    
    ' Or see if mouse got pressed from somewhere off this form
    ' (in case of it being modal)
    ElseIf MouseIsPressed(True) Then
        ' If we are scrolling, turn the scrolling variable off
        If m.bScrolling Then
            m.bScrolling = False
            DoEvents
        
        ' Otherwise, set flag and wait to see if flag gets cleared
        Else
            m.bMousePressed = True
            dStart = GetTickCount
            Do While m.bMousePressed
                DoEvents
                If GetTickCount - dStart > 200 Then
                    m.lSelectedRow = -1&
                    HideMe
                    Exit Sub
                End If
            Loop
        End If
        
    ' If no mouse button is being pressed, yet the scrolling is on, turn it off
    ElseIf m.bScrolling Then
        m.bScrolling = False
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PlaceForm
'' Description: Place the form according to the location of the parent and make
''              sure that it shows on the screen
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PlaceForm()

    Dim lIndex As Long                  ' Index into a for loop
    Dim lpPoint As POINTAPI             ' Pixel Location
    Dim lRC As Long                     ' Return code
    Dim dScreenWidth As Double          ' Width of the screen
    Dim dScreenHeight As Double         ' Height of the screen

    ' If there is no parent, don't bother to do anything
    If m.ctlParent Is Nothing Then Exit Sub

    ' See if parent has moved
    lpPoint.X = 0
    lpPoint.Y = 0
    lRC = ClientToScreen(m.ctlParent.hWnd, lpPoint)
    If lpPoint.X = m.lpParent.X And lpPoint.Y = m.lpParent.Y Then
        Exit Sub
    End If
    
    ' Save parent location
    m.lpParent.X = lpPoint.X
    m.lpParent.Y = lpPoint.Y

    ' Get screen width and height
    dScreenWidth = Screen.Width
    dScreenHeight = Screen.Height
    
    ' Adjust if multiple-monitor wide (center in the first monitor)
    For lIndex = 2 To 99
        If dScreenWidth / lIndex < dScreenHeight * 1.24 Then
            dScreenWidth = dScreenWidth / (lIndex - 1) ' lIndex-1 = # monitors wide
            Exit For
        End If
    Next

    ' Align with parent's bottom left corner
    lpPoint.X = 0 'm.ctlParent.Width \ Screen.TwipsPerPixelX
    lpPoint.Y = m.ctlParent.Height \ Screen.TwipsPerPixelY
    lRC = ClientToScreen(m.ctlParent.hWnd, lpPoint)
    
    ' Subtract 2 pixels (don't know why!)
    Me.Top = (lpPoint.Y - 2) * Screen.TwipsPerPixelY
    Me.Left = (lpPoint.X - 2) * Screen.TwipsPerPixelX '- Me.Width

    ' Adjust if off screen
    If Me.Left + Me.Width > dScreenWidth Then
        ' Right-align instead
        Me.Left = Me.Left - (Me.Width - m.ctlParent.Width)
    End If
    If Me.Top + Me.Height > dScreenHeight Then
        ' Put on top of parent
        Me.Top = Me.Top - Me.Height - m.ctlParent.Height
    End If
    
    ' Adjust if still off screen
    If Me.Left + Me.Width > dScreenWidth Then
        Me.Left = dScreenWidth - Me.Width
    End If
    If Me.Top + Me.Height > dScreenHeight Then
        Me.Top = dScreenHeight - Me.Height
    End If
    If Me.Left < 0 Then Me.Left = 0
    If Me.Top < 0 Then Me.Top = 0

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ParentMoved
'' Description: See if the parent has moved since we last saved its location
'' Inputs:      None
'' Returns:     TRUE if the parent moved, FALSE otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ParentMoved() As Boolean

    Dim lpPoint As POINTAPI             ' Current location of parent
    Dim lRC As Long                     ' Return code
        
    ' Only check if there is a parent
    If Not m.ctlParent Is Nothing Then
        lpPoint.X = 0
        lpPoint.Y = 0
        lRC = ClientToScreen(m.ctlParent.hWnd, lpPoint)
        If lpPoint.X <> m.lpParent.X Or lpPoint.Y <> m.lpParent.Y Then
            ParentMoved = True
        End If
    End If

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize the grid and show the form
'' Inputs:      Parent control, Tree to fill grid with, Default Selected Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMe(Parent As Control, DropTree As cgdDropTree)

    Dim iIndex As Integer               ' Index into a for loop
    Dim bVisible As Boolean             ' Is the form visible?
    
    ' Is the form visible?
    bVisible = Me.Visible
    
    ' Set the local Parent control to the one passed in
    Set m.ctlParent = Parent
    
    ' Init grid
    If True Then 'Not m.bLoaded Then 'Or fgDropTree.Rows <> nRows Or fgDropTree.ColWidth(0) <> lCellSize Then

        m.lpParent.X = -9999 ' to force placing form

        With fgDropTree
            .Rows = 0
            .FixedRows = 0
            .Cols = 4
            .FixedCols = 0
            .ColHidden(1) = True
            .ColHidden(2) = True
            .ColHidden(3) = True
            .ColDataType(3) = flexDTBoolean
            .SelectionMode = flexSelectionListBox
            .ExtendLastCol = True
            .OutlineBar = flexOutlineBarSimpleLeaf
            .GridLines = flexGridNone
            .AllowSelection = False
            .AllowUserResizing = flexResizeNone
            .FocusRect = flexFocusNone
            .ExplorerBar = flexExNone
            .ScrollTrack = True
            .ScrollTips = True
            .Editable = flexEDNone
            
            FillTree DropTree
            If DropTree.Selected <> 0 Then
                ShowRow DropTree.Selected - 1
            End If
            
            .RowHeight(-1) = 17 * Screen.TwipsPerPixelY

            If Not m.ctlParent Is Nothing Then
                .BackColor = m.ctlParent.BackColor
            End If

            .BackColorBkg = .BackColor
            .SheetBorder = .BackColor
        End With
        
        ' Size form
        Me.Move Me.Left, Me.Top, fgDropTree.Width, fgDropTree.Height
    End If
    
    m.lSelectedRow = 0&
    m.bMousePressed = False
    PlaceForm
    ShowTree
    
    On Error Resume Next
    fgDropTree.SetFocus
    tmrDropTree.Enabled = True
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillTree
'' Description: Fill the tree control with the given GD Tree
'' Inputs:      cGdTree with information to fill the tree with
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FillTree(DropTree As cgdDropTree)
On Error Resume Next

    Dim lIndex As Long

    With fgDropTree
        .Redraw = flexRDNone
        For lIndex = 1 To DropTree.Count
            .AddItem DropTree.Item(lIndex).Text
            .IsSubtotal(.Rows - 1) = True
            .RowOutlineLevel(.Rows - 1) = DropTree.NodeLevel(lIndex)
            If m.bHasPicture Then .Cell(flexcpPicture, .Rows - 1, 0) = DropTree.Item(lIndex).Picture
            .TextMatrix(.Rows - 1, 1) = DropTree.Item(lIndex).Key
            .TextMatrix(.Rows - 1, 2) = DropTree.Item(lIndex).ToolTipText
            .TextMatrix(.Rows - 1, 3) = DropTree.Item(lIndex).Selectable
        Next lIndex
        .Redraw = flexRDBuffered
    End With
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowRow
'' Description: Show a certain row in the tree.  Also make sure to show all
''              children that can be shown (if the row has children)
'' Inputs:      Row to be shown
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowRow(ByVal lRow As Long)
On Error Resume Next
    
    Dim lLastChild As Long              ' Last child of the given row
    
    With fgDropTree
        ' Make the row passed in the current row
        .Row = lRow
        .RowSel = lRow
        
        ' Make last child visible if possible
        lLastChild = .GetNodeRow(lRow, flexNTLastChild)
        If lLastChild > 0 Then .ShowCell lLastChild, 0
        
        ' Make sure selected row is visible
        .ShowCell lRow, 0
    End With

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowTree
'' Description: Show the form either modally or non-modally depending on the
''              situation
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowTree()
    
    If m.bActivating Then Exit Sub
    
    m.bActivating = True
    DoEvents
    On Error GoTo ShowModal
    tmrDropTree.Enabled = True
    Me.Show
    GoTo ShowExit
    
ShowModal:
    On Error GoTo 0
    Me.Show 1

ShowExit:
    DoEvents
    m.bActivating = False
    Exit Sub

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ItemSelected
'' Description: Notify the parent that an item has been selected
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ItemSelected()
        
    If Not m.ctlParent Is Nothing Then
        m.ctlParent.Text = " " ' trigger
    End If
    HideMe

End Sub
#End If
