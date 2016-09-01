VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmLibraryAddItem 
   Caption         =   "Form1"
   ClientHeight    =   5070
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   6330
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3615
      Left            =   4860
      TabIndex        =   2
      Top             =   120
      Width           =   1335
      Begin VB.CheckBox chkShowLocal 
         Caption         =   "&Show Local Rules"
         Height          =   435
         Left            =   0
         TabIndex        =   8
         Top             =   2460
         Width           =   1335
      End
      Begin VB.CommandButton cmdUsedIn 
         Caption         =   "Used &In"
         Height          =   495
         Left            =   0
         TabIndex        =   6
         Top             =   1740
         Width           =   1335
      End
      Begin VB.CommandButton cmdUses 
         Caption         =   "&Uses"
         Height          =   495
         Left            =   0
         TabIndex        =   5
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   495
         Left            =   0
         TabIndex        =   4
         Top             =   540
         Width           =   1335
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   495
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1335
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgItems 
      Height          =   3255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4395
      _cx             =   7752
      _cy             =   5741
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
   Begin RichTextLib.RichTextBox txtPreview 
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   3960
      Width           =   6120
      _ExtentX        =   10795
      _ExtentY        =   1720
      _Version        =   393217
      BackColor       =   -2147483648
      ReadOnly        =   -1  'True
      TextRTF         =   $"frmLibraryAddItem.frx":0000
   End
   Begin VB.Label lblHeader 
      Caption         =   "Highlight the items which you wish to move to your library"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4395
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Begin VB.Menu mnuUses 
         Caption         =   "&Uses"
      End
      Begin VB.Menu mnuUsedIn 
         Caption         =   "Used &In"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangeFont 
         Caption         =   "&Change Font"
      End
   End
End
Attribute VB_Name = "frmLibraryAddItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmLibraryAddItem.frm
'' Description: Form to allow user to add items to a library
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 04/03/2013   DAJ         Move Strategy Baskets into database
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    bOK As Boolean
    astrGridItems As cGdArray
    astrLibItems As cGdArray
    lRowHeight As Long
    Security As cSecurity
End Type
Private m As mPrivate

Public Enum eGDLibAddItemsMode
    eGDLibAddItemsMode_All = 0
    eGDLibAddItemsMode_Uses = 1
    eGDLibAddItemsMode_UsedIn = 2
End Enum

Private Enum eGDCols
    eGDCol_Select = 0
    eGDCol_Name = 1
    eGDCol_ItemType = 2
    eGDCol_LibraryName = 3
    eGDCol_ItemTypeCat = 4
    eGDCol_LastModified = 5
    eGDCol_Preview = 6
    eGDCol_ID = 7
    eGDCol_SecurityLevel = 8
    eGDCol_Password = 9
    eGDCol_CannotDelete = 10
    eGDCol_SystemNumber = 11
End Enum
Private Const kGridCols = 12

Private Function GDCol(ByVal lColumn As eGDCols) As Long
    GDCol = lColumn
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and show the form
'' Inputs:      None
'' Returns:     True if OK pressed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(Mode As eGDLibAddItemsMode, ByVal lItemID As Long, ByVal strItemType As String, _
                        astrLibItems As cGdArray) As cGdArray
On Error GoTo ErrSection:

    Dim astrItems As New cGdArray       ' Items to return
    Dim lCol As Long                    ' Index into a for loop
    Dim lRow As Long                    ' Index into a for loop
    Dim strToAdd As String              ' String to add to the array

    If Parse(strItemType, " ", 2) = "Rule" Then strItemType = "Rule"

    Screen.MousePointer = vbHourglass
    With fgItems
        .Redraw = flexRDNone
        InitGrid
        Set m.astrLibItems = astrLibItems
        Select Case Mode
            Case eGDLibAddItemsMode_All
                LoadGrid
                'chkShowLocal.Visible = True
            Case eGDLibAddItemsMode_UsedIn
                LoadUsedIn lItemID, strItemType
                'chkShowLocal.Visible = False
            Case eGDLibAddItemsMode_Uses
                LoadUses lItemID, strItemType
                'chkShowLocal.Visible = False
        End Select
        
        ' If we have any rows, select the first row
        If .Rows > .FixedRows Then
            .Row = .FixedRows
            .RowSel = .FixedRows
        End If
        .Redraw = flexRDBuffered
    End With
    
    ' Only show the Uses and Used In buttons when in "All" mode
    If Mode <> eGDLibAddItemsMode_All Then
        cmdUsedIn.Visible = False
        cmdUses.Visible = False
    End If
    Screen.MousePointer = vbDefault
    
    ShowForm Me, True
    If m.bOK Then
        With fgItems
            astrItems.Create eGDARRAY_Strings
            'For lRow = .FixedRows To .Rows - 1
            '    If CheckedCell(fgItems, lRow, GDCol(eGDCol_Select)) = True Then
            For lRow = 0 To .SelectedRows - 1
                    strToAdd = .TextMatrix(.SelectedRow(lRow), 0)
                    For lCol = 1 To .Cols - 1
                        strToAdd = strToAdd & vbTab & .TextMatrix(.SelectedRow(lRow), lCol)
                    Next lCol
                    astrItems.Add strToAdd
            Next lRow
            '    End If
            'Next lRow
        End With
        Set ShowMe = astrItems
    End If
    
ErrExit:
    Set m.astrGridItems = Nothing
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmLibraryAddItem.ShowMe", eGDRaiseError_Raise, g.strAppPath
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkShowLocal_Click
'' Description: Show or Hide the local rules as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkShowLocal_Click()
On Error GoTo ErrSection:

    ShowLocalRules

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryAddItem.chkShowLocal.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Close the form without saving
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
    RaiseError "frmLibraryAddItem.cmdCancel.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: Close the form and save
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOk_Click()
On Error GoTo ErrSection:

    m.bOK = True
    Me.Hide
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryAddItem.cmdOK.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdUsedIn_Click
'' Description: Show a list of things that the given item is used in
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdUsedIn_Click()
On Error GoTo ErrSection:

    UsedIn

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryAddItem.cmdUsedIn.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdUses_Click
'' Description: Show a list of things that the given item uses
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdUses_Click()
On Error GoTo ErrSection:

    Uses

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryAddItem.cmdUses.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgItems_AfterSort
'' Description: When the user sorts the grid, set the back colors again
'' Inputs:      Column sorted, Order sorted in
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgItems_AfterSort(ByVal Col As Long, Order As Integer)
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw

    With fgItems
        lRedraw = .Redraw
        .Redraw = flexRDNone
        SetBackColors
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryAddItem.fgItems.AfterSort", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgItems_BeforeEdit
'' Description: Only allow the user to edit the 'Select' column
'' Inputs:      Row and Column to be edited, Whether to Cancel the edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgItems_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    If Col <> GDCol(eGDCol_Select) Then
        Cancel = True
    ElseIf Not cmdUsedIn.Visible Then
        With fgItems
            If UCase(.TextMatrix(Row, GDCol(eGDCol_LibraryName))) <> "USER LIBRARY" Then
                Cancel = True
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryAddItem.fgItems.BeforeEdit", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgItems_AfterRowColChange
'' Description: Enable/Disable controls based on the Item Type
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgItems_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    If fgItems.Redraw <> flexRDNone Then
        If fgItems.RowSel <> fgItems.Row Then
            fgItems.RowSel = fgItems.Row
        End If
        
        Enable cmdUsedIn, UCase(Left(fgItems.TextMatrix(fgItems.Row, GDCol(eGDCol_ItemType)), 8)) <> "BASKET"
        ItemPreview
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryAddItem.fgItems.AfterRowColChange", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgItems_MouseDown
'' Description: Handle the user clicking in the grid
'' Inputs:      Mouse Button pressed, Shift/Ctrl/Alt status, Mouse Location
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgItems_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim lMouseRow As Long
    Dim lMouseCol As Long
    
    With fgItems
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If Button = vbRightButton Then
            If lMouseRow >= .FixedRows And lMouseRow < .Rows Then
                .RowSel = lMouseRow
                If Not .IsSelected(lMouseRow) Then
                    .Row = lMouseRow
                End If
            End If
            
            mnuUses.Enabled = (lMouseRow >= .FixedRows And lMouseRow < .Rows)
            mnuUsedIn.Enabled = (lMouseRow >= .FixedRows And lMouseRow < .Rows)
            
            PopupMenu mnuPopUp
            If mnuPopUp.Tag = "Uses" Then
                Uses
            ElseIf mnuPopUp.Tag = "UsedIn" Then
                UsedIn
            End If
            mnuPopUp.Tag = ""
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryAddItem.fgItems.MouseDown", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgItems_MouseMove
'' Description: Handle the user moving their mouse over the grid
'' Inputs:      Mouse Button pressed, Shift/Ctrl/Alt status, Mouse Location
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgItems_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    Dim lMouseRow As Long               ' Row of the grid the mouse is in
    Dim lMouseCol As Long               ' Column of the grid the mouse is in
    Dim strTooltip As String            ' Tooltip text
    
    With fgItems
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If (lMouseRow < .FixedRows) And (lMouseRow >= 0) Then
            strTooltip = "Sort By: " & Trim(.TextMatrix(lMouseRow, lMouseCol))
        Else
            strTooltip = ""
        End If
        
        If strTooltip <> .ToolTipText Then
            .ToolTipText = strTooltip
        End If
    End With

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_KeyDown
'' Description: Handle the user pressing a key on the form
'' Inputs:      Code of key pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyF1 Then
        KeyCode = 0
        g.Help.ShowF1Help Me
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryAddItem.Form.KeyDown", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize the controls and the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strPlacement As String          ' Form Placement from the registry
    Dim strFont As String

    Width = 11000
    CenterTheForm Me
    Caption = "Move Items to a Library..."
    Icon = Picture16("kSelect")
    chkShowLocal = GetIniFileProperty("LibAdd_ShowLocal", vbChecked, "Library", g.strIniFile)
    strPlacement = GetIniFileProperty("LibAdd", "", "Placement", g.strIniFile)
    If strPlacement <> "" Then SetFormPlacement Me, strPlacement, "LHT"
    
    Set m.Security = New cSecurity
    
    mnuPopUp.Visible = False
    strFont = GetIniFileProperty("LibraryAddItems", "", "Fonts", g.strIniFile)
    If strFont <> "" Then FontFromString fgItems.Font, strFont
    
    ' Hide the Show Local check box and turn it off (DAJ: 01/05/2004)...
    chkShowLocal.Value = vbUnchecked
    chkShowLocal.Visible = False
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryAddItem.Form.Load", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Close the form if the user hits the 'X'
'' Inputs:      Whether to Cancel the Unload, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode = 0 Then
        Cancel = True
        m.bOK = False
        Me.Hide
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryAddItem.Form.QueryUnload", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Move and resize the controls as the form resizes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    Dim lScaleWidth As Long             ' Minimum scale width
    Dim lScaleHeight As Long            ' Minimum scale height
    
    ' Figure out the minum scale height and width
    lScaleWidth = lblHeader.Width + fraButtons.Width + (lblHeader.Left * 3)
    lScaleHeight = fraButtons.Height + txtPreview.Height + (fraButtons.Top * 3)
    If LimitFormSize(Me, lScaleWidth, lScaleHeight) Then Exit Sub
    
    With fraButtons
        .Move ScaleWidth - fraButtons.Width - lblHeader.Left
    End With
    
    With txtPreview
        .Move lblHeader.Left, ScaleHeight - txtPreview.Height - lblHeader.Top, _
                ScaleWidth - (lblHeader.Left * 2)
    End With
    
    With fgItems
        .Move lblHeader.Left, lblHeader.Height + (lblHeader.Top * 2), _
                ScaleWidth - fraButtons.Width - (lblHeader.Left * 3), _
                ScaleHeight - lblHeader.Height - txtPreview.Height - (lblHeader.Top * 4)
    End With

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: When the form unloads, save some information about it
'' Inputs:      Whether to Cancel the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    SetIniFileProperty "LibAdd_ShowLocal", chkShowLocal, "Library", g.strIniFile
    SetIniFileProperty "LibAdd", GetFormPlacement(Me), "Placement", g.strIniFile
    SetIniFileProperty "LibraryAddItems", FontToString(fgItems.Font), "Fonts", g.strIniFile

ErrExit:
    Set m.Security = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryAddItem.Form.Unload", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuChangeFont_Click
'' Description: Allow the user to change fonts on the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuChangeFont_Click()
On Error GoTo ErrSection:

    ChangeGridFont fgItems

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryAddItem.mnuChangeFont.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuUsedIn_Click
'' Description: Allow the user to see what items the current item is used in
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuUsedIn_Click()
On Error GoTo ErrSection:

    mnuPopUp.Tag = "UsedIn"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryAddItem.mnuUsedIn.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuUses_Click
'' Description: Allow the user to see what the current item uses
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuUses_Click()
On Error GoTo ErrSection:

    mnuPopUp.Tag = "Uses"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryAddItem.mnuUses.Click", eGDRaiseError_Show, g.strAppPath
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

    With fgItems
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Clear
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExSortShowAndMove
        .ExtendLastCol = True
        .SelectionMode = flexSelectionListBox
        .AllowUserResizing = flexResizeColumns
        .AllowSelection = True
        .AutoSearch = flexSearchFromTop
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .ScrollBars = flexScrollBarBoth
        .ScrollTrack = True
        .WordWrap = False
        .FixedCols = 0
        .FixedRows = 1
        .Rows = 1
        .Cols = kGridCols
        m.lRowHeight = .RowHeight(0)
        .RowHeightMax = .Height - (.RowHeight(0) * 2)
        
        .TextMatrix(0, GDCol(eGDCol_Select)) = "Select"
        .TextMatrix(0, GDCol(eGDCol_Name)) = "Name"
        .TextMatrix(0, GDCol(eGDCol_ItemType)) = "Item"
        .TextMatrix(0, GDCol(eGDCol_LibraryName)) = "Library"
        .TextMatrix(0, GDCol(eGDCol_ItemTypeCat)) = "Item Type"
        .TextMatrix(0, GDCol(eGDCol_LastModified)) = "Last Modified"
        
        .ColHidden(GDCol(eGDCol_Select)) = True
        .ColHidden(GDCol(eGDCol_Preview)) = True
        .ColHidden(GDCol(eGDCol_ID)) = True
        .ColHidden(GDCol(eGDCol_SystemNumber)) = True
        .ColHidden(GDCol(eGDCol_SecurityLevel)) = True
        .ColHidden(GDCol(eGDCol_Password)) = True
        .ColHidden(GDCol(eGDCol_CannotDelete)) = True
        
        .ColDataType(GDCol(eGDCol_Select)) = flexDTBoolean
        .ColDataType(GDCol(eGDCol_CannotDelete)) = flexDTBoolean
        
        .ColFormat(GDCol(eGDCol_LastModified)) = DateAndTime("Format")
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        .AutoSize 0, .Cols - 1, False, 75
        
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryAddItem.InitGrid", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGrid
'' Description: Load up the grid with stuff from the User Library
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim rs As Recordset                 ' Recordset from the database
    Dim rs2 As Recordset                ' Recordset from the database
    Dim strPreview As String            ' Preview string
    Dim strItemTypeCat As String        ' Item Type category
    Dim strItemType As String           ' Item type

    Set m.astrGridItems = New cGdArray
    m.astrGridItems.Create eGDARRAY_Strings

    With fgItems
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .ColHidden(GDCol(eGDCol_LibraryName)) = True
        
        'Load Systems
        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblSystems] " & _
            "WHERE [LibraryID]=" & Str(kUserLibrary) & ";", dbOpenSnapshot)
        If Not (rs.BOF And rs.EOF) Then rs.MoveFirst
        Do Until rs.EOF
            If rs!CheckSum = BuildCheckSum(rs, "tblSystems") Then
                AddRow rs!SystemName, "Strategy" & vbTab & "0", rs!LastModified, "N/A", _
                    rs!Notes, rs!SystemNumber, rs!SecurityLevel, DecryptField(rs!Password), _
                    rs!CannotDelete, ""
                        
                Set rs2 = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] " & _
                    "WHERE [SystemNumber]=" & Str(rs!SystemNumber) & " " & _
                    "ORDER BY [Name];", dbOpenSnapshot)
                If Not (rs2.BOF And rs2.EOF) Then rs2.MoveFirst
                Do Until rs2.EOF
                    If rs2!CheckSum = BuildCheckSum(rs2, "tblRules") Then
                        If rs2!BuySell = True Then
                            If rs2!RuleType = 0 Then
                                strItemTypeCat = "Long Entry"
                            Else
                                strItemTypeCat = "Short Exit"
                            End If
                        Else
                            If rs2!RuleType = 0 Then
                                strItemTypeCat = "Short Entry"
                            Else
                                strItemTypeCat = "Long Exit"
                            End If
                        End If
                        AddRow rs2!Name, "Local Rule" & vbTab & rs2!SystemNumber, rs2!LastModified, strItemTypeCat, _
                            DecryptField(rs2!PreviewRTF), rs2!RuleID, _
                            rs2!SecurityLevel, DecryptField(rs2!Password), rs2!CannotDelete, ""
                    End If
                    
                    rs2.MoveNext
                Loop
            End If
            
            rs.MoveNext
        Loop
        
        ' Load all Functions
        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctions] " & _
                "WHERE [LibraryID]=" & Str(kUserLibrary) & ";", dbOpenSnapshot)
        If Not (rs.BOF And rs.EOF) Then rs.MoveFirst
        Do Until rs.EOF
            If rs!CheckSum = BuildCheckSum(rs, "tblFunctions") Then
                strPreview = "Usage: " & rs!TradeSenseUsage & Chr(13) & Chr(10) & _
                         "Description: " & rs!Description
    
                AddRow rs!FunctionName, "Function" & vbTab & "0", _
                    rs!LastModified, ImplementationTypeDesc(rs!ImplementationTypeID), _
                    strPreview, rs!FunctionID, _
                    rs!SecurityLevel, DecryptField(rs!Password), rs!CannotDelete, ""
            End If
            
            rs.MoveNext
        Loop

        ' Load all Shared Rules
        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] " & _
            "WHERE [LibraryID]=" & Str(kUserLibrary) & " AND [SystemNumber]=0;", dbOpenSnapshot)
        If Not (rs.BOF And rs.EOF) Then rs.MoveFirst
        Do Until rs.EOF
            If rs!CheckSum = BuildCheckSum(rs, "tblRules") Then
                If rs!BuySell = True Then
                    If rs!RuleType = 0 Then
                        strItemTypeCat = "Long Entry"
                    Else
                        strItemTypeCat = "Short Exit"
                    End If
                Else
                    If rs!RuleType = 0 Then
                        strItemTypeCat = "Short Entry"
                    Else
                        strItemTypeCat = "Long Exit"
                    End If
                End If
                If rs!SystemNumber = 0 Then strItemType = "Shared Rule" Else strItemType = "Local Rule"
                AddRow rs!Name, strItemType & vbTab & rs!SystemNumber, rs!LastModified, strItemTypeCat, _
                    DecryptField(rs!PreviewRTF), rs!RuleID, _
                    rs!SecurityLevel, DecryptField(rs!Password), rs!CannotDelete, ""
            End If
            
            rs.MoveNext
        Loop
        
        ' Strategy Baskets
        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblStrategyBaskets] WHERE [LibraryID]=" & Str(kUserLibrary) & ";", dbOpenDynaset)
        Do While Not rs.EOF
            If rs!CheckSum = BuildCheckSum(rs, "tblStrategyBaskets") Then
                AddRow rs!Name, "Basket" & vbTab & "0", rs!LastModified, "N/A", _
                    rs!Description, rs!StrategyBasketID, rs!SecurityLevel, DecryptField(rs!Password), _
                    rs!CannotDelete, ""
            End If
            
            rs.MoveNext
        Loop
        
        ShowLocalRules
        If .Rows > 1 Then
            .Row = 1
            .RowSel = 1
            .Col = GDCol(eGDCol_Name)
            .Sort = flexSortGenericAscending
            Enable cmdUsedIn, UCase(Left(.TextMatrix(.Row, GDCol(eGDCol_ItemType)), 8)) <> "BASKET"
            ItemPreview
        End If
        
        SetBackColors
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryAddItem.LoadGrid", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddRow
'' Description: Add a row to the grid
'' Inputs:      Recordset to use to fill the row in the grid
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddRow(pName As Variant, pItemType As Variant, pLastMod As Variant, _
    pstrItemTypeCat As Variant, pPreview As Variant, pID As Variant, _
    pSecurityLevel As Variant, pPassword As Variant, pCannotDelete As Variant, _
    pLibraryName As Variant)
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim lSystemNumber As Long           ' System number of a rule
    Dim strItemType As String           ' Item type of the current item
    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' Item found in the grid
    Dim strSearch As String             ' String to search for
    Dim lRow As Long
    Dim strPreview As String
    
    strItemType = Parse(Str(pItemType), vbTab, 1)
    lSystemNumber = CLng(Parse(Str(pItemType), vbTab, 2))
    
    strPreview = Replace(Replace(pPreview, vbCrLf, "||"), Chr(9), Chr(1))
    
    With fgItems
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        strSearch = strItemType & vbTab & Str(pID)
        If Not m.astrLibItems.BinarySearch(strSearch) Then
            If Not m.astrGridItems.BinarySearch(strSearch, lIndex) Then
                m.astrGridItems.Add strSearch, lIndex
                
                If lSystemNumber = 0 Then
                    .Rows = .Rows + 1
                    lRow = .Rows - 1
                    .TextMatrix(lRow, GDCol(eGDCol_Name)) = pName
                    .TextMatrix(lRow, GDCol(eGDCol_ItemType)) = strItemType
                    If IsNull(pstrItemTypeCat) Then .TextMatrix(lRow, GDCol(eGDCol_ItemTypeCat)) = "" Else .TextMatrix(lRow, GDCol(eGDCol_ItemTypeCat)) = pstrItemTypeCat
                    .TextMatrix(lRow, GDCol(eGDCol_LastModified)) = Str(pLastMod) ' DateFormat(pLastMod) & " " & Format(pLastMod, "hh:mm:ss AM/PM")
                    .TextMatrix(lRow, GDCol(eGDCol_Preview)) = strPreview
                    .TextMatrix(lRow, GDCol(eGDCol_ID)) = Str(pID)
                    .TextMatrix(lRow, GDCol(eGDCol_SecurityLevel)) = Str(pSecurityLevel)
                    .TextMatrix(lRow, GDCol(eGDCol_Password)) = NullChk(pPassword)
                    .TextMatrix(lRow, GDCol(eGDCol_CannotDelete)) = pCannotDelete
                    CheckedCell(fgItems, lRow, GDCol(eGDCol_CannotDelete)) = CBool(pCannotDelete)
                    .TextMatrix(lRow, GDCol(eGDCol_SecurityLevel)) = Str(pSecurityLevel)
                    .TextMatrix(lRow, GDCol(eGDCol_SystemNumber)) = Str(lSystemNumber)
                    .TextMatrix(lRow, GDCol(eGDCol_LibraryName)) = pLibraryName
                Else
                    lRow = .Rows - 1
                    AddLocalRule lRow, GDCol(eGDCol_Name), "    " & pName
                    AddLocalRule lRow, GDCol(eGDCol_ItemType), strItemType
                    AddLocalRule lRow, GDCol(eGDCol_ItemTypeCat), NullChk(pstrItemTypeCat)
                    'AddLocalRule lRow, GDCol(eGDCol_LastModified), DateFormat(pLastMod) & " " & Format(pLastMod, "hh:mm:ss AM/PM")
                    AddLocalRule lRow, GDCol(eGDCol_ID), Str(pID)
                    AddLocalRule lRow, GDCol(eGDCol_SystemNumber), Str(lSystemNumber)
                End If
            End If
        End If
        
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmLibraryAddItem.AddRow", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetBackColors
'' Description: Set the background color of the rows appropriately
'' Inputs:      Grid that is currently active
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetBackColors()
On Error GoTo ErrSection:

    Dim bAlt As Boolean                 ' Is this an alternate row?
    Dim lRow As Long                    ' Index into a for loop
    Dim lRedraw As Long                 ' Current state of the redraw
    
    With fgItems
        lRedraw = .Redraw
        .Redraw = flexRDNone
        For lRow = .FixedRows To .Rows - 1
            If .RowHidden(lRow) = False Then
                If Not bAlt Then
                    .Cell(flexcpBackColor, lRow, 0, lRow, .Cols - 1) = .BackColor
                Else
                    .Cell(flexcpBackColor, lRow, 0, lRow, .Cols - 1) = .BackColorAlternate
                End If
                bAlt = Not bAlt
            End If
        Next lRow
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmLibraryAddItem.SetBackColors", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ItemPreview
'' Description: Show the preview for the currently selected item
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ItemPreview()
On Error GoTo ErrSection:
    
    Dim lLength As Long
    Dim Rule As Object
    Dim lRow As Long
    Dim strPreview As String
    
    Set Rule = CreateObject(g.strCommonDLL & "cRule")
    txtPreview.Text = ""
    lRow = fgItems.RowSel
    
    strPreview = fgItems.TextMatrix(lRow, GDCol(eGDCol_Preview))
    strPreview = Replace(Replace(strPreview, "||", vbCrLf), Chr(1), Chr(9))
        
    If InStr(fgItems.TextMatrix(lRow, GDCol(eGDCol_ItemType)), "Rule") Then
        If m.Security.CanPreview(CByte(ValOfText(fgItems.Cell(flexcpValue, lRow, GDCol(eGDCol_SecurityLevel))))) Then
            txtPreview.TextRTF = Rule.GetRTF(strPreview)
        Else
            txtPreview.SelColor = vbBlack
            txtPreview.Text = "Not authorized to view"
        End If
    Else
        With txtPreview
            If Len(strPreview) = 0 Then
                .Text = "No Description"
                .SelStart = 0
                .SelLength = Len(.Text)
                .SelColor = vbBlack
                .SelBold = False
                .SelItalic = False
                .SelLength = 0
            Else
                .Text = strPreview
                lLength = InStr(.Text, Chr(13)) - InStr(.Text, ": ") - 2
                If lLength > 0 Then
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SelColor = vbBlack
                    .SelItalic = False
                    .SelStart = InStr(.Text, ": ") + 1
                    .SelLength = InStr(.Text, Chr(13)) - InStr(.Text, ": ") - 2
                    If .SelLength > 0 Then
                        .SelBold = True
                    Else
                        .SelBold = False
                    End If
                    .SelLength = 0
                Else
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SelColor = vbBlack
                    .SelBold = False
                    .SelItalic = False
                    .SelLength = 0
                End If
            End If
        End With
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmLibraryAddItem.ItemPreview", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadUsedIn
'' Description: Load the grid with items the item was used in
'' Inputs:      Item ID, Item Type
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadUsedIn(ByVal lItemID As Long, ByVal strItemType As String)
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim rs As Recordset                 ' Recordset from the database
    Dim strPreview As String            ' Preview string for the item
    Dim strItemTypeCat As String        ' Item Type category
    
    Set m.astrGridItems = New cGdArray
    m.astrGridItems.Create eGDARRAY_Strings
    
    With fgItems
        lRedraw = .Redraw
        .Redraw = flexRDNone

        Select Case UCase(strItemType)
            Case "STRATEGY"
                .ColHidden(GDCol(eGDCol_ItemTypeCat)) = True
                Set rs = g.dbNav.OpenRecordset("SELECT tblStrategyBaskets.*, tblLibrarys.LibraryName " & _
                                "FROM tblLibrarys INNER JOIN (tblStrategyBaskets INNER JOIN tblStrategyBasketItems ON tblStrategyBaskets.StrategyBasketID = tblStrategyBasketItems.StrategyBasketID) ON tblLibrarys.LibraryID = tblStrategyBaskets.LibraryID " & _
                                "WHERE (((tblStrategyBasketItems.SystemNumber)=" & Str(lItemID) & "));", dbOpenSnapshot)
                If Not (rs.BOF And rs.EOF) Then rs.MoveFirst
                Do While Not rs.EOF
                    If rs!CheckSum = BuildCheckSum(rs, "tblStrategyBaskets") Then
                        AddRow rs!Name, "Basket" & vbTab & "0", rs!LastModified, _
                                "N/A", rs!Description, rs!StrategyBasketID, rs!SecurityLevel, _
                                DecryptField(rs!Password), rs!CannotDelete, rs!LibraryName
                        If rs!LibraryID = kUserLibrary Then
                            CheckedCell(fgItems, .Rows - 1, GDCol(eGDCol_Select)) = True
                        End If
                    End If
                    
                    rs.MoveNext
                Loop
            
            Case "RULE"
                .ColHidden(GDCol(eGDCol_ItemType)) = True
                .ColHidden(GDCol(eGDCol_ItemTypeCat)) = True
                Set rs = g.dbNav.OpenRecordset("SELECT tblSystems.*, tblLibrarys.LibraryName " & _
                                "FROM tblLibrarys INNER JOIN (tblSystems INNER JOIN tblSystemRules ON tblSystems.SystemNumber = tblSystemRules.SystemNumber) ON tblLibrarys.LibraryID = tblSystems.LibraryID " & _
                                "WHERE (((tblSystemRules.RuleID)=" & Str(lItemID) & "));", dbOpenSnapshot)
                If Not (rs.BOF And rs.EOF) Then rs.MoveFirst
                Do While Not rs.EOF
                    If rs!CheckSum = BuildCheckSum(rs, "tblSystems") Then
                        AddRow rs!SystemName, "Strategy" & vbTab & "0", rs!LastModified, _
                                "N/A", rs!Notes, rs!SystemNumber, rs!SecurityLevel, _
                                DecryptField(rs!Password), rs!CannotDelete, rs!LibraryName
                        If rs!LibraryID = kUserLibrary Then
                            CheckedCell(fgItems, .Rows - 1, GDCol(eGDCol_Select)) = True
                        End If
                    End If
                    
                    rs.MoveNext
                Loop

            Case "FUNCTION"
                .ColHidden(GDCol(eGDCol_ItemTypeCat)) = True
                Set rs = g.dbNav.OpenRecordset("SELECT tblFunctions.*, tblLibrarys.LibraryName " & _
                                "FROM (tblLibrarys INNER JOIN tblFunctions ON tblLibrarys.LibraryID = tblFunctions.LibraryID) INNER JOIN tblFunctionRefs ON tblFunctions.FunctionID = tblFunctionRefs.FunctionID " & _
                                "WHERE (((tblFunctionRefs.FunctionIDRef)=" & Str(lItemID) & "));", dbOpenSnapshot)
                If Not (rs.BOF And rs.EOF) Then rs.MoveFirst
                Do While Not rs.EOF
                    If rs!CheckSum = BuildCheckSum(rs, "tblFunctions") Then
                        strPreview = "Usage: " & rs!TradeSenseUsage & Chr(13) & Chr(10) & _
                                 "Description: " & rs!Description
            
                        AddRow rs!FunctionName, "Function" & vbTab & "0", _
                            rs!LastModified, ImplementationTypeDesc(rs!ImplementationTypeID), _
                            strPreview, rs!FunctionID, _
                            rs!SecurityLevel, DecryptField(rs!Password), rs!CannotDelete, rs!LibraryName
                        If rs!LibraryID = kUserLibrary Then
                            CheckedCell(fgItems, .Rows - 1, GDCol(eGDCol_Select)) = True
                        End If
                    End If
                    
                    rs.MoveNext
                Loop
                
                Set rs = g.dbNav.OpenRecordset("SELECT tblRules.*, tblLibrarys.LibraryName " & _
                                "FROM (tblLibrarys INNER JOIN tblRules ON tblLibrarys.LibraryID = tblRules.LibraryID) INNER JOIN tblFunctionRules ON tblRules.RuleID = tblFunctionRules.RuleID " & _
                                "WHERE (((tblFunctionRules.FunctionIDRef)=" & Str(lItemID) & "));", dbOpenSnapshot)
                If Not (rs.BOF And rs.EOF) Then rs.MoveFirst
                Do Until rs.EOF
                    If rs!CheckSum = BuildCheckSum(rs, "tblRules") Then
                        If rs!BuySell = True Then
                            If rs!RuleType = 0 Then
                                strItemTypeCat = "Long Entry"
                            Else
                                strItemTypeCat = "Short Exit"
                            End If
                        Else
                            If rs!RuleType = 0 Then
                                strItemTypeCat = "Short Entry"
                            Else
                                strItemTypeCat = "Long Exit"
                            End If
                        End If
                        AddRow rs!Name, "Rule" & vbTab & rs!SystemNumber, rs!LastModified, strItemTypeCat, _
                            DecryptField(rs!PreviewRTF), rs!RuleID, _
                            rs!SecurityLevel, DecryptField(rs!Password), rs!CannotDelete, rs!LibraryName
                        If rs!LibraryID = kUserLibrary Then
                            CheckedCell(fgItems, .Rows - 1, GDCol(eGDCol_Select)) = True
                        End If
                    End If
                    
                    rs.MoveNext
                Loop
                
        End Select

        If .Rows > 1 Then
            .Row = 1
            .RowSel = 1
            .Col = GDCol(eGDCol_Name)
            .Sort = flexSortGenericAscending
            Enable cmdUsedIn, UCase(Left(.TextMatrix(.Row, GDCol(eGDCol_ItemType)), 8)) <> "BASKET"
            ItemPreview
        End If
    
        SetBackColors
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = lRedraw
    End With

ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryAddItem.LoadUsedIn", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadUses
'' Description: Load the grid with items the item uses
'' Inputs:      Item ID, Item Type
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadUses(ByVal lItemID As Long, ByVal strItemType As String)
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim rs As Recordset                 ' Recordset from the database
    Dim rs2 As Recordset
    Dim rs3 As Recordset
    Dim strItemTypeCat As String
    Dim strPreview As String
    
    Set m.astrGridItems = New cGdArray
    m.astrGridItems.Create eGDARRAY_Strings
    
    With fgItems
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        Select Case UCase(strItemType)
            Case "BASKET"
                LoadUsesStrategyBasket lItemID
            
            Case "STRATEGY"
                LoadUsesRuleStrategy lItemID
            
            Case "RULE"
                LoadUsesFunctionRule lItemID
            
            Case "FUNCTION"
                LoadUsesFunctionFunction lItemID
        
        End Select

        If .Rows > 1 Then
            .Row = 1
            .RowSel = 1
            .Col = GDCol(eGDCol_Name)
            .Sort = flexSortGenericAscending
            Enable cmdUsedIn, UCase(Left(.TextMatrix(.Row, GDCol(eGDCol_ItemType)), 8)) <> "BASKET"
            ItemPreview
        End If
    
        SetBackColors
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = lRedraw
    End With

ErrExit:
    Set rs = Nothing
    Set rs2 = Nothing
    Set rs3 = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryAddItem.LoadUses", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadUsesStrategyBasket
'' Description: Load the grid with strategies that the given basket uses
'' Inputs:      Basket ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadUsesStrategyBasket(ByVal lBasketID As Long)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database

    Set rs = g.dbNav.OpenRecordset("SELECT tblSystems.*, tblLibrarys.LibraryName " & _
                "FROM tblLibrarys INNER JOIN (tblSystems INNER JOIN tblStrategyBasketItems ON tblSystems.SystemNumber = tblStrategyBasketItems.SystemNumber) ON tblLibrarys.LibraryID = tblSystems.LibraryID " & _
                "WHERE (((tblStrategyBasketItems.StrategyBasketID)=" & Str(lBasketID) & "));", dbOpenSnapshot)
    If Not (rs.BOF And rs.EOF) Then rs.MoveFirst
    Do While Not rs.EOF
        If rs!CheckSum = BuildCheckSum(rs, "tblSystems") Then
            AddRow rs!SystemName, "Strategy" & vbTab & "0", rs!LastModified, "N/A", _
                rs!Notes, rs!SystemNumber, rs!SecurityLevel, DecryptField(rs!Password), _
                rs!CannotDelete, rs!LibraryName
            If rs!LibraryID = kUserLibrary Then
                CheckedCell(fgItems, fgItems.Rows - 1, GDCol(eGDCol_Select)) = True
            End If
            
            LoadUsesRuleStrategy rs!SystemNumber
        End If
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryAddItem.LoadUsesStrategyBasket", , g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadUsesRuleStrategy
'' Description: Load the grid with rules that the given strategy uses
'' Inputs:      Strategy ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadUsesRuleStrategy(ByVal lStrategyID As Long)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim strPreview As String            ' Preview text
    Dim strItemTypeCat As String        ' Item type category

    Set rs = g.dbNav.OpenRecordset("SELECT tblRules.*, tblLibrarys.LibraryName " & _
            "FROM tblLibrarys INNER JOIN (tblRules INNER JOIN tblSystemRules ON tblRules.RuleID = tblSystemRules.RuleID) ON tblLibrarys.LibraryID = tblRules.LibraryID " & _
            "WHERE (((tblSystemRules.SystemNumber)=" & Str(lStrategyID) & "));", dbOpenSnapshot)
    If Not (rs.BOF And rs.EOF) Then rs.MoveFirst
    Do While Not rs.EOF
        If rs!CheckSum = BuildCheckSum(rs, "tblRules") Then
            If rs!SystemNumber = 0 Then
                If rs!BuySell = True Then
                    If rs!RuleType = 0 Then
                        strItemTypeCat = "Long Entry"
                    Else
                        strItemTypeCat = "Short Exit"
                    End If
                Else
                    If rs!RuleType = 0 Then
                        strItemTypeCat = "Short Entry"
                    Else
                        strItemTypeCat = "Long Exit"
                    End If
                End If
                AddRow rs!Name, "Shared Rule" & vbTab & rs!SystemNumber, _
                    rs!LastModified, strItemTypeCat, DecryptField(rs!PreviewRTF), rs!RuleID, _
                    rs!SecurityLevel, DecryptField(rs!Password), rs!CannotDelete, rs!LibraryName
                If rs!LibraryID = kUserLibrary Then
                    CheckedCell(fgItems, fgItems.Rows - 1, GDCol(eGDCol_Select)) = True
                End If
            End If
            
            LoadUsesFunctionRule rs!RuleID
        End If
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryAddItem.LoadUsesRuleStrategy", , g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadUsesFunctionRule
'' Description: Load the grid with functions that the given rule uses
'' Inputs:      Rule ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadUsesFunctionRule(ByVal lRuleID As Long)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim strPreview As String            ' Preview text for the function

    Set rs = g.dbNav.OpenRecordset("SELECT tblFunctions.*, tblLibrarys.LibraryName " & _
                "FROM tblLibrarys INNER JOIN (tblFunctions INNER JOIN tblFunctionRules ON tblFunctions.FunctionID = tblFunctionRules.FunctionIDRef) ON tblLibrarys.LibraryID = tblFunctions.LibraryID " & _
                "WHERE (((tblFunctionRules.RuleID)=" & Str(lRuleID) & "));", dbOpenSnapshot)
    If Not (rs.BOF And rs.EOF) Then rs.MoveFirst
    Do While Not rs.EOF
        If rs!CheckSum = BuildCheckSum(rs, "tblFunctions") Then
            strPreview = "Usage: " & rs!TradeSenseUsage & Chr(13) & Chr(10) & _
                     "Description: " & rs!Description
            AddRow rs!FunctionName, "Function" & vbTab & "0", _
                rs!LastModified, ImplementationTypeDesc(rs!ImplementationTypeID), _
                strPreview, rs!FunctionID, _
                rs!SecurityLevel, DecryptField(rs!Password), rs!CannotDelete, rs!LibraryName
            If rs!LibraryID = kUserLibrary Then
                CheckedCell(fgItems, fgItems.Rows - 1, GDCol(eGDCol_Select)) = True
            End If
            
            LoadUsesFunctionFunction rs!FunctionID
        End If
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryAddItem.LoadUsesFunctionRule", , g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadUsesFunctionFunction
'' Description: Load the grid with functions that the given function uses
'' Inputs:      Function ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadUsesFunctionFunction(ByVal lFunctionID As Long)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim strPreview As String            ' Preview text for the function

    Set rs = g.dbNav.OpenRecordset("SELECT tblFunctions.*, tblLibrarys.LibraryName " & _
                "FROM (tblLibrarys INNER JOIN tblFunctions ON tblLibrarys.LibraryID = tblFunctions.LibraryID) INNER JOIN tblFunctionRefs ON tblFunctions.FunctionID = tblFunctionRefs.FunctionIDRef " & _
                "WHERE (((tblFunctionRefs.FunctionID)=" & Str(lFunctionID) & "));", dbOpenSnapshot)
    If Not (rs.BOF And rs.EOF) Then rs.MoveFirst
    Do While Not rs.EOF
        If rs!CheckSum = BuildCheckSum(rs, "tblFunctions") Then
            strPreview = "Usage: " & rs!TradeSenseUsage & Chr(13) & Chr(10) & _
                     "Description: " & rs!Description
            AddRow rs!FunctionName, "Function" & vbTab & "0", _
                rs!LastModified, ImplementationTypeDesc(rs!ImplementationTypeID), _
                strPreview, rs!FunctionID, _
                rs!SecurityLevel, DecryptField(rs!Password), rs!CannotDelete, rs!LibraryName
            If rs!LibraryID = kUserLibrary Then
                CheckedCell(fgItems, fgItems.Rows - 1, GDCol(eGDCol_Select)) = True
            End If
        End If
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryAddItem.LoadUsesFunctionFunction", , g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddLocalRule
'' Description: Add a local rule to the same row as the appropriate system
'' Inputs:      Row and Column to add to, Item to add
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddLocalRule(ByVal lRow As Long, ByVal lCol As Long, ByVal strItem As String)
On Error GoTo ErrSection:

    With fgItems
        .TextMatrix(lRow, lCol) = .TextMatrix(lRow, lCol) & vbCrLf & strItem
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryAddItem.AddLocalRule", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowLocalRules
'' Description: Show/Hide the local rules as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowLocalRules()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lRedraw As Long                 ' Current state of the grid's redraw

    With fgItems
        lRedraw = .Redraw
        .Redraw = flexRDNone
        For lIndex = .FixedRows To .Rows - 1
            If chkShowLocal = vbChecked Then
                .RowHeight(lIndex) = RowHeight(Me, .Cell(flexcpFont, lIndex, 1), .TextMatrix(lIndex, 1)) + 50
            Else
                .RowHeight(lIndex) = .RowHeight(0)
            End If
        Next lIndex
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLibraryAddItem.ShowLocalRules", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Uses
'' Description: Allow the user to select items that an item uses
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Uses()
On Error GoTo ErrSection:

    Dim frm As New frmLibraryAddItem    ' Form to show
    Dim astrItems As New cGdArray       ' List of items chosen from the form
    Dim lIndex As Long                  ' Index into a for loop
    Dim lRow As Long                    ' Index into a for loop
    Dim strID As String
    Dim strItemType As String

    With fgItems
        strID = Parse(.TextMatrix(.RowSel, GDCol(eGDCol_ID)), vbCrLf, 1)
        strItemType = Parse(.TextMatrix(.RowSel, GDCol(eGDCol_ItemType)), vbCrLf, 1)
        Set astrItems = frm.ShowMe(eGDLibAddItemsMode_Uses, CLng(strID), strItemType, m.astrLibItems)
        If astrItems.Size > 0 Then
            For lIndex = 0 To astrItems.Size - 1
                For lRow = .FixedRows To .Rows - 1
                    If .TextMatrix(lRow, GDCol(eGDCol_Name)) = Parse(astrItems(lIndex), vbTab, GDCol(eGDCol_Name) + 1) Then
                        If .TextMatrix(lRow, GDCol(eGDCol_ID)) = Parse(astrItems(lIndex), vbTab, GDCol(eGDCol_ID) + 1) Then
                            CheckedCell(fgItems, lRow, GDCol(eGDCol_Select)) = True
                            Exit For
                        End If
                    End If
                Next lRow
            Next lIndex
            CheckedCell(fgItems, .RowSel, GDCol(eGDCol_Select)) = True
        End If
    End With

ErrExit:
    Set frm = Nothing
    Set astrItems = Nothing
    Exit Sub
    
ErrSection:
    Set frm = Nothing
    Set astrItems = Nothing
    RaiseError "frmLibraryAddItem.Uses", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UsedIn
'' Description: Allow the user to select items that an item is used in
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UsedIn()
On Error GoTo ErrSection:

    Dim frm As New frmLibraryAddItem    ' Form to show
    Dim astrItems As New cGdArray       ' List of items chosen from the form
    Dim lIndex As Long                  ' Index into a for loop
    Dim lRow As Long                    ' Index into a for loop
    Dim strID As String                 ' ID for the item
    Dim strItemType As String           ' Type for the item

    With fgItems
        If UCase(Left(.TextMatrix(.RowSel, GDCol(eGDCol_ItemType)), 8)) = "BASKET" Then
            Err.Raise vbObjectError + 1000, , "Cannot view UsedIn for a Strategy Basket"
        Else
            strID = Parse(.TextMatrix(.RowSel, GDCol(eGDCol_ID)), vbCrLf, 1)
            strItemType = Parse(.TextMatrix(.RowSel, GDCol(eGDCol_ItemType)), vbCrLf, 1)
            
            Set astrItems = frm.ShowMe(eGDLibAddItemsMode_UsedIn, CLng(strID), strItemType, m.astrLibItems)
            If astrItems.Size > 0 Then
                For lIndex = 0 To astrItems.Size - 1
                    For lRow = .FixedRows To .Rows - 1
                        If .TextMatrix(lRow, GDCol(eGDCol_Name)) = Parse(astrItems(lIndex), vbTab, GDCol(eGDCol_Name) + 1) Then
                            If .TextMatrix(lRow, GDCol(eGDCol_ID)) = Parse(astrItems(lIndex), vbTab, GDCol(eGDCol_ID) + 1) Then
                                CheckedCell(fgItems, lRow, GDCol(eGDCol_Select)) = True
                                Exit For
                            End If
                        End If
                    Next lRow
                Next lIndex
                CheckedCell(fgItems, .RowSel, GDCol(eGDCol_Select)) = True
            End If
        End If
    End With

ErrExit:
    Set frm = Nothing
    Set astrItems = Nothing
    Exit Sub
    
ErrSection:
    Set frm = Nothing
    Set astrItems = Nothing
    RaiseError "frmLibraryAddItem.UsedIn", eGDRaiseError_Raise, g.strAppPath
    
End Sub
