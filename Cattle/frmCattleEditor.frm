VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form frmCattleEditor 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrMenu 
      Enabled         =   0   'False
      Left            =   660
      Top             =   2520
   End
   Begin ActiveToolBars.SSActiveToolBars tbToolbar 
      Left            =   120
      Top             =   2460
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131083
      ToolBarsCount   =   1
      ToolsCount      =   6
      Tools           =   "frmCattleEditor.frx":0000
      ToolBars        =   "frmCattleEditor.frx":1A53
   End
   Begin VSFlex7LCtl.VSFlexGrid fgObjects 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2595
      _cx             =   4577
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
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Begin VB.Menu mnuAdd 
         Caption         =   "Add"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove"
      End
      Begin VB.Menu mnuManage 
         Caption         =   "Manage"
      End
   End
End
Attribute VB_Name = "frmCattleEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmCattleEditor.frm
'' Description: Form for allowing user to edit certain cattle stuff
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 03/07/2014   DAJ         Moved into NavCattle.DLL
'' 03/20/2014   DAJ         Added a "Click Here" line
'' 04/15/2014   DAJ         Give user ability to manage ingredients from ration editor
'' 04/24/2014   DAJ         Fix for being able to delete "click here" rows
'' 05/22/2014   DAJ         Renamed frmTurnkeyManage to frmCattleManage
'' 05/22/2014   DAJ         Renamed frmTurnkeyEditor to frmCattleEditor; Renamed g.Turnkey to g.Cattle
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Enum eGDManageModes
    eGDManageMode_Ration
End Enum

Private Enum eGDRationCols
    eGDRationCols_IngredientID = 0
    eGDRationCols_Ingredient = 1
    eGDRationCols_PoundsFed = 2
    eGDRationCols_PercentMarkup = 3
    
    eGDRationCols_NumCols
End Enum

Private Type mPrivate
    nMode As eGDManageModes             ' Mode of the form
    bClosing As Boolean                 ' Is the form closing?
    
    strDryFeedPct As String             ' Default dry feed percent
    strFeedyardID As String             ' Feed Yard ID
    Ration As cBrokerMessage            ' Ration object
    astrIngredients As cGdArray         ' Array of ingredients
    strName As String                   ' Name of the object
    iButton As Integer                  ' Mouse button pressed
    bAlreadyDone As Boolean             ' Has the form activate code already been done?

    lExtendCol As Long                  ' Extend column
    lPrevColWidth As Long               ' Previous column width
End Type
Private m As mPrivate

Private Property Get Dirty() As Boolean
    Dirty = tbToolbar.Tools("ID_Save").Enabled
End Property
Private Property Let Dirty(ByVal bDirty As Boolean)
    tbToolbar.Tools("ID_Save").Enabled = bDirty
End Property

Private Property Get RationCol(ByVal nCol As eGDRationCols) As Long
    RationCol = nCol
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMeRation
'' Description: Setup and show form for managing rations
'' Inputs:      Ration
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMeRation(Ration As cBrokerMessage)
On Error GoTo ErrSection:

    m.nMode = eGDManageMode_Ration
    SetEditorCaption Me, "Ration", Ration("RationName")
    m.bClosing = False
    Set m.Ration = Ration
    m.strName = Ration("RationName")
    m.bAlreadyDone = False
    
    InitGridRation
    LoadGridRation
    LoadIngredients
    Dirty = False
    
    With tbToolbar
        .Tools("ID_Save").TooltipText = "Save the ration"
        .Tools("ID_SaveAs").TooltipText = "Save a copy of the ration"
        .Tools("ID_Rename").TooltipText = "Rename the ration"
        .Tools("ID_Exit").TooltipText = "Exit the dialog"
        .Tools("ID_Add").TooltipText = "Add an ingredient to the ration"
        .Tools("ID_Remove").TooltipText = "Remove the selected ingredient from the ration"
    End With

    ShowForm Me, eForm_Modal, g.frmMain

ErrExit:
    Unload Me
    Exit Sub
    
ErrSection:
    Unload Me
    RaiseError "frmCattleEditor.ShowMeRation"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Cattle_Ingredient
'' Description: Ingredient record returned from the cattle source
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Cattle_Ingredient(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim strFirstField As String         ' First field in the message

    If m.bClosing = False Then
        If (m.nMode = eGDManageMode_Ration) And (Len(strMessage) > 0) Then
            strFirstField = Parse(strMessage, vbTab, 1)
            
            If UCase(strFirstField) = "BEGIN" Then
            ElseIf UCase(strFirstField) = "END" Then
                LoadIngredients
            Else
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleEditor.Cattle_Ingredient", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Cattle_Ration
'' Description: Ration record returned from the cattle source
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Cattle_Ration(ByVal strMessage As String)
On Error GoTo ErrSection:

    Dim Ration As cBrokerMessage        ' Ration object
    Dim lRow As Long                    ' Row in the grid
    Dim strFirstField As String         ' First field in the message
    Dim bMatch As Boolean               ' Is this a match?

    If (m.bClosing = False) And (Len(strMessage) > 0) Then
        Select Case m.nMode
            Case eGDManageMode_Ration
                strFirstField = Parse(strMessage, vbTab, 1)
                
                If UCase(strFirstField) = "BEGIN" Then
                ElseIf UCase(strFirstField) = "END" Then
                Else
                    Set Ration = New cBrokerMessage
                    Ration.FromString strMessage
                    
                    bMatch = False
                    If Len(m.Ration("ID")) = 0 Then
                        bMatch = (Ration("RequestID") = m.Ration("RequestID"))
                    Else
                        bMatch = (Ration("ID") = m.Ration("ID"))
                    End If
                    
                    If bMatch Then
                        Set m.Ration = Ration
                        LoadGridRation
                        
                        Dirty = False
                    End If
                End If
        
        End Select
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleEditor.Cattle_Ration", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgObjects_AfterEdit
'' Description: Handle the user changing the value of a cell
'' Inputs:      Row, Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgObjects_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    If m.nMode = eGDManageMode_Ration Then
        With fgObjects
            If Col = RationCol(eGDRationCols_Ingredient) Then
                .TextMatrix(Row, RationCol(eGDRationCols_IngredientID)) = g.Cattle.IngredientIDForName(.TextMatrix(Row, RationCol(eGDRationCols_Ingredient)))
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleEditor.fgObjects_AfterEdit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgObjects_AfterRowColChange
'' Description: Handle the user moving cells
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgObjects_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    Dim bValidDataRow As Boolean        ' Is the new row a valid data row?

    bValidDataRow = ValidDataRow(NewRow)

    If m.nMode = eGDManageMode_Ration Then
        If bValidDataRow = True Then
            EditCell fgObjects
        End If
        
        tbToolbar.Tools("ID_Remove").Enabled = bValidDataRow
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleEditor.fgObjects_AfterRowColChange"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgObjects_BeforeEdit
'' Description: Decide if we want the cell to be edited
'' Inputs:      Row, Column, Cancel Edit?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgObjects_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    Select Case m.nMode
        Case eGDManageMode_Ration
            Cancel = RowIsClickHereLine(Row) Or RowIsClickHereManageLine(Row)
        
        Case Else
            Cancel = True
            
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleEditor.fgObjects_BeforeEdit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgObjects_BeforeMouseDown
'' Description: Bring up the context menu on a right-click in the grid
'' Inputs:      Mouse Button pressed, Shift/Ctrl/Alt status, Mouse location,
''              Cancel the Mouse Down?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgObjects_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Mouse row in the grid

    m.iButton = Button
    If Button = vbRightButton Then
        With fgObjects
            lMouseRow = .MouseRow
            
            .Row = lMouseRow
            
            If Validate(False) Then
                Select Case m.nMode
                    Case eGDManageMode_Ration
                        mnuAdd.Caption = "Add Ingredient"
                        mnuEdit.Visible = False
                        mnuRemove.Caption = "Remove Ingredient"
                        mnuManage.Caption = "Manage Ingredients"
                
                End Select
                
                Enable mnuEdit, ValidDataRow(lMouseRow)
                Enable mnuRemove, mnuEdit.Enabled
                
                PopupMenu mnuPopup
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleEditor.fgObjects_BeforeMouseDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgObjects_BeforeRowColChange
'' Description: Handle the user wanting to change the current cell
'' Inputs:      Old Row and Column, New Row and Column, Cancel?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgObjects_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    If (NewRow <> OldRow) And (OldRow <> -1&) And (NewRow <> -1&) Then
        Select Case m.nMode
            Case eGDManageMode_Ration
                If Len(fgObjects.TextMatrix(OldRow, RationCol(eGDRationCols_Ingredient))) = 0 Then
                    InfBox "You must select an ingredient", "!", , "Error"
                    Cancel = True
                    
                    EditCell fgObjects, , RationCol(eGDRationCols_Ingredient)
                End If
                
        End Select
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleEditor.fgObjects_BeforeRowColChange"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgObjects_Click
'' Description: Handle the user clicking in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgObjects_Click()
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Mouse row in the grid
    
    With fgObjects
        lMouseRow = .MouseRow
        
        If m.iButton = vbLeftButton Then
            If ValidGridRow(fgObjects, lMouseRow) Then
                If RowIsClickHereLine(lMouseRow) Then
                    AddObject
                ElseIf RowIsClickHereManageLine(lMouseRow) Then
                    ManageObjects
                End If
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleEditor.fgObjects_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgObjects_DblClick
'' Description: If the user double clicks on an object, edit it
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgObjects_DblClick()
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Mouse row in the grid
    
    With fgObjects
        lMouseRow = .MouseRow
        
        If m.iButton = vbLeftButton Then
            If ValidGridRow(fgObjects, lMouseRow) Then
                .Row = lMouseRow
                EditObject
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleEditor.fgObjects_DblClick"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgObjects_KeyDown
'' Description: Handle the user pressing the Insert or Delete keys in the grid
'' Inputs:      Key Pressed, Shift/Ctrl/Alt Status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgObjects_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyInsert Then
        AddObject
    ElseIf KeyCode = vbKeyDelete Then
        RemoveObject
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleEditor.fgObjects_KeyDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgObjects_KeyPress
'' Description: Handle the user pressing the Enter key in the grid
'' Inputs:      Key Pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgObjects_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    If KeyAscii = vbKeyReturn Then
        EditObject
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleEditor.fgObjects_KeyPress"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgObjects_ValidateEdit
'' Description: Validate the information entered by the user
'' Inputs:      Row, Column, Cancel Edit?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgObjects_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim dValue As Double                ' Value of the edit text

    If m.bClosing = False Then
        With fgObjects
            If m.nMode = eGDManageMode_Ration Then
                If ValidDataRow(Row) Then
                    Cancel = False
                    
                    Select Case Col
                        Case RationCol(eGDRationCols_Ingredient):
                            'If Len(.EditText) = 0 Then
                            '    InfBox "You must select an ingredient", "!", , "Error"
                            '    Cancel = True
                            'End If
                        
                        Case RationCol(eGDRationCols_PoundsFed):
                            If IsAlpha(.EditText) Then
                                InfBox "Pounds Fed must be a number", "!", , "Error"
                                Cancel = True
                            End If
                        Case RationCol(eGDRationCols_PercentMarkup):
                            If IsAlpha(.EditText) Then
                                InfBox "Percent Markup must be a number", "!", , "Error"
                                Cancel = True
                            End If
                    
                    End Select
                    
                    If (Cancel = False) And (.EditText <> .TextMatrix(Row, Col)) Then
                        Dirty = True
                    End If
                Else
                    Cancel = True
                End If
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleEditor.fgObjects_ValidateEdit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: Handle the form getting activated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:

    If m.bAlreadyDone = False Then
        m.bAlreadyDone = True
        
        If fgObjects.Rows = fgObjects.FixedRows + 2 Then
            If RowIsClickHereLine(fgObjects.FixedRows) Then
                AddObject
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleEditor.Form_Activate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Setup form when it is loaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Icon = g.AppBridge.Picture16(g.Cattle.IconName)
    
    g.Styler.StyleForm Me
    
    PlaceForm Me
    
    With tbToolbar
        .Tools("ID_Save").Picture = g.AppBridge.Picture16(g.AppBridge.ToolbarIcon("kSave"))
        .Tools("ID_SaveAs").Picture = g.AppBridge.Picture16(g.AppBridge.ToolbarIcon("kSaveAs"))
        .Tools("ID_Rename").Picture = g.AppBridge.Picture16(g.AppBridge.ToolbarIcon("kRename"))
        .Tools("ID_Exit").Picture = g.AppBridge.Picture16(g.AppBridge.ToolbarIcon("kCancel"))
    End With
    
    mnuPopup.Visible = False
    
    tmrMenu.Enabled = False
    tmrMenu.Interval = 10
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleEditor.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Determine whether or not to let the form close
'' Inputs:      Cancel Unload?, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        Cancel = True
        
        If AskToSave Then
            m.bClosing = True
            Hide
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmCattleEditor.Form_QueryUnload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Size and move controls as the form is resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    Dim lSpace As Long                  ' Space between the controls
    Dim lMinScaleWidth As Long          ' Minimum scale width for the form
    Dim lMinScaleHeight As Long         ' Minimum scale height for the form
    
    lSpace = 120
    lMinScaleWidth = (1215 * 3) + (lSpace * 3)
    lMinScaleHeight = 2295 + (lSpace * 2)
    
    If LimitFormSize(Me, lMinScaleWidth, lMinScaleHeight) = False Then
        If m.nMode = eGDManageMode_Ration Then
            With fgObjects
                .Move lSpace, lSpace, ScaleWidth - (lSpace * 2), ScaleHeight - (lSpace * 2)
            End With
        End If
        
        ExtendCustomColumn
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Clean up when the form is unloaded
'' Inputs:      Cancel the Unload?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    SaveFormPlacement Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleEditor.Form_Unload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuAdd_Click
'' Description: Handle the user wanting to add an object from the context menu
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuAdd_Click()
On Error GoTo ErrSection:

    tmrMenu.Tag = "Add"
    tmrMenu.Enabled = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleEditor.mnuAdd_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuEdit_Click
'' Description: Handle the user wanting to edit an object from the context menu
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuEdit_Click()
On Error GoTo ErrSection:

    tmrMenu.Tag = "Edit"
    tmrMenu.Enabled = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleEditor.mnuEdit_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuManage_Click
'' Description: Handle the user wanting to manage objects from the context menu
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuManage_Click()
On Error GoTo ErrSection:

    tmrMenu.Tag = "Manage"
    tmrMenu.Enabled = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleEditor.mnuManage_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuRemove_Click
'' Description: Handle the user wanting to remove an object from the context menu
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuRemove_Click()
On Error GoTo ErrSection:

    tmrMenu.Tag = "Remove"
    tmrMenu.Enabled = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleEditor.mnuRemove_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tbToolbar_ToolClick
'' Description: Handle the user clicking on an item in the toolbar
'' Inputs:      Tool clicked
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tbToolbar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
On Error GoTo ErrSection:

    Select Case UCase(Tool.ID)
        Case "ID_SAVE", "ID_SAVEAS", "ID_RENAME"
            Save UCase(Tool.ID)
        
        Case "ID_EXIT"
            If AskToSave Then
                Hide
            End If
            
        Case "ID_ADD"
            AddObject
        
        Case "ID_REMOVE"
            RemoveObject
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleEditor.tbToolbar_ToolClick"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tmrMenu_Timer
'' Description: Handle the menu timer
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrMenu_Timer()
On Error GoTo ErrSection:

    Dim strTag As String                ' Tag from the control
    
    tmrMenu.Enabled = False
    strTag = tmrMenu.Tag
    tmrMenu.Tag = ""

    Select Case UCase(strTag)
        Case "ADD"
            AddObject
            
        Case "EDIT"
            EditObject
            
        Case "MANAGE"
            ManageObjects
            
        Case "REMOVE"
            RemoveObject
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleEditor.tmrMenu_Timer"
    
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

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    
    With fgObjects
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = True
        .BackColorAlternate = .BackColor
        .BackColorBkg = vbApplicationWorkspace
        .Editable = flexEDNone
        .ExplorerBar = flexExSortShow
        .ExtendLastCol = True
        .GridLines = flexGridFlat
        .GridLinesFixed = flexGridInset
        .MergeCells = flexMergeFree
        .OutlineBar = flexOutlineBarNone
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        .WordWrap = True
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleEditor.InitGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitGridRation
'' Description: Initialize the grid for a ration
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitGridRation()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    
    With fgObjects
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        InitGrid
        
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDKbdMouse
        .SelectionMode = flexSelectionFree
        .TabBehavior = flexTabCells
        
        .Rows = 1
        .FixedRows = 1
        .Cols = RationCol(eGDRationCols_NumCols)
        .FixedCols = 0
        
        .TextMatrix(0, RationCol(eGDRationCols_IngredientID)) = "IngredientID"
        .TextMatrix(0, RationCol(eGDRationCols_Ingredient)) = "Ingredient"
        .TextMatrix(0, RationCol(eGDRationCols_PoundsFed)) = "Pounds Fed"
        .TextMatrix(0, RationCol(eGDRationCols_PercentMarkup)) = "% Markup"
        
        .ColHidden(RationCol(eGDRationCols_IngredientID)) = True
        
        m.lExtendCol = 1
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleEditor.InitGridRation"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGridRation
'' Description: Load the grid for a ration
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGridRation()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim astrIngredient As cGdArray      ' Ingredient information split out into an array
    Dim astrPoundsFed As cGdArray       ' Pounds Fed information split out into an array
    Dim astrPctMarkup As cGdArray       ' Percent Markup informaion split out into an array
    Dim lIndex As Long                  ' Index into a for loop
    Dim lRow As Long                    ' Row in the grid
    
    With fgObjects
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Rows = .FixedRows
        
        Set astrIngredient = New cGdArray
        Set astrPoundsFed = New cGdArray
        Set astrPctMarkup = New cGdArray
        
        astrIngredient.SplitFields m.Ration("IngredientID"), "|"
        astrPoundsFed.SplitFields m.Ration("PoundsFed"), "|"
        astrPctMarkup.SplitFields m.Ration("PercentMarkup"), "|"
        
        For lIndex = 0 To astrIngredient.Size - 1
            If (Len(astrIngredient(lIndex)) > 0) And (astrIngredient(lIndex) <> "-1") Then
                lRow = AddObject(False)
                
                .TextMatrix(lRow, RationCol(eGDRationCols_IngredientID)) = astrIngredient(lIndex)
                .TextMatrix(lRow, RationCol(eGDRationCols_Ingredient)) = g.Cattle.IngredientNameForID(astrIngredient(lIndex))
                .TextMatrix(lRow, RationCol(eGDRationCols_PoundsFed)) = astrPoundsFed(lIndex)
                .TextMatrix(lRow, RationCol(eGDRationCols_PercentMarkup)) = astrPctMarkup(lIndex)
            End If
        Next lIndex
        
        AddClickHereLine
        AddClickHereManageLine
        
        .AutoSize 0, .Cols - 1, False, 75
        ExtendCustomColumn
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleEditor.LoadGridRation"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExtendCustomColumn
'' Description: Adjust all column widths to accomodate custom extended column
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ExtendCustomColumn(Optional ByVal nResizeCol As Long = -1)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim lTotal As Long                  ' Total width
    Dim lDiff As Long                   ' Difference in column width

    With fgObjects
        ' if column being resized is after the extended column,
        ' then change the width of the next visible column instead
        If nResizeCol >= m.lExtendCol Then
            .Redraw = flexRDNone
            lDiff = .ColWidth(nResizeCol) - m.lPrevColWidth
            For lIndex = nResizeCol + 1 To .Cols - 1
                If Not .ColHidden(lIndex) Then
                    .ColWidth(lIndex) = .ColWidth(lIndex) - lDiff
                    Exit For
                End If
            Next
            m.lPrevColWidth = 0
        End If
        
        ' size the custom extended column in order to fill the client width
        .ColHidden(m.lExtendCol) = True
        .Redraw = flexRDBuffered '(must do this so .ClientWidth will be correct)
        .Redraw = flexRDNone
        lTotal = 0
        For lIndex = 0 To .Cols - 1
            If Not .ColHidden(lIndex) Then
                lTotal = lTotal + .ColWidth(lIndex)
            End If
        Next
        lTotal = .ClientWidth - lTotal
        If lTotal > 0 Then .ColWidth(m.lExtendCol) = lTotal
        .ColHidden(m.lExtendCol) = False
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleEditor.ExtendCustomColumn"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SelectedObject
'' Description: Grab the object on the selected row in the grid
'' Inputs:      None
'' Returns:     Selected Object ( Nothing if none )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SelectedObject() As cBrokerMessage
On Error GoTo ErrSection:

    Dim ReturnObject As cBrokerMessage  ' Object to return
    
    Set ReturnObject = Nothing
    With fgObjects
        If ValidGridRow(fgObjects) Then
            Set ReturnObject = .RowData(.Row)
        End If
    End With
    
    Set SelectedObject = ReturnObject

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCattleEditor.SelectedObject"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddObject
'' Description: Add a new object
'' Inputs:      Edit cell
'' Returns:     Row in the grid
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AddObject(Optional ByVal bEditCell As Boolean = True) As Long
On Error GoTo ErrSection:
    
    Dim lClickHereLine As Long          ' Click here line in the grid
    Dim lRow As Long                    ' Row in the grid

    lClickHereLine = ClickHereLine

    Select Case m.nMode
        Case eGDManageMode_Ration
            With fgObjects
                .Rows = .Rows + 1
                
                If lClickHereLine = -1& Then
                    lRow = .Rows - 1
                Else
                    .RowPosition(.Rows - 1) = lClickHereLine
                    lRow = lClickHereLine
                End If
                    
                .MergeRow(lRow) = False
                
                If bEditCell Then
                    EditCell fgObjects, lRow, RationCol(eGDRationCols_Ingredient)
                End If
            End With
            
            Dirty = True
            
    End Select

    AddObject = lRow

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCattleEditor.AddObject"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EditObject
'' Description: Edit an existing object
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EditObject()
On Error GoTo ErrSection:

    Dim ToEdit As cBrokerMessage        ' Object to edit

    With fgObjects
        If ValidGridRow(fgObjects) Then
            If m.nMode = eGDManageMode_Ration Then
            Else
                Set ToEdit = SelectedObject
                If Not SelectedObject Is Nothing Then
                    Select Case m.nMode
                    End Select
                End If
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleEditor.EditObject"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveObject
'' Description: Remove an existing object
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RemoveObject()
On Error GoTo ErrSection:

    Dim ToRemove As cBrokerMessage      ' Object to remove
    Dim strResponse As String           ' Response back from the InfBox

    With fgObjects
        If ValidDataRow Then
            If m.nMode = eGDManageMode_Ration Then
                If Len(.TextMatrix(.Row, RationCol(eGDRationCols_Ingredient))) = 0 Then
                    strResponse = "Y"
                Else
                    strResponse = InfBox("Are you sure you want to remove '" & .TextMatrix(.Row, RationCol(eGDRationCols_Ingredient)) & "'?", "?", "+Yes|-No", "Confirmation")
                End If
                
                If strResponse = "Y" Then
                    .RemoveItem .Row
                    
                    If .Rows > .FixedRows + 2 Then
                        .Row = .Rows - 3
                        .Col = RationCol(eGDRationCols_Ingredient)
                    Else
                        .Row = -1&
                        .Col = -1&
                    End If
                    
                    Dirty = True
                End If
            Else
                Set ToRemove = SelectedObject
                If Not SelectedObject Is Nothing Then
                    Select Case m.nMode
                    End Select
                End If
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleEditor.RemoveObject"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ManageObjects
'' Description: Allow the user to manage objects
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ManageObjects()
On Error GoTo ErrSection:

    Dim frm As frmCattleManage         ' Management form

    Select Case m.nMode
        Case eGDManageMode_Ration
            Set frm = New frmCattleManage
            frm.ShowMeIngredients
            
    End Select

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmCattleEditor.ManageObjects"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FindRow
'' Description: Find the row for the given ID
'' Inputs:      ID
'' Returns:     Row where that ID exists ( -1 if not found )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FindRow(ByVal strID As String) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim RowObject As cBrokerMessage     ' Row object
    
    lReturn = -1&
    With fgObjects
        For lIndex = .FixedRows To .Rows - 1
            If TypeOf .RowData(lIndex) Is cBrokerMessage Then
                Set RowObject = .RowData(lIndex)
                If RowObject("ID") = strID Then
                    lReturn = lIndex
                    Exit For
                End If
            End If
        Next lIndex
    End With
    
    FindRow = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCattleEditor.FindRow"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AskToSave
'' Description: See if we need to prompt the user to save changes
'' Inputs:      None
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AskToSave() As Boolean
On Error GoTo ErrSection:
    
    Dim bReturn As Boolean              ' Return value for the function
    Dim strResponse As String           ' Response from the InfBox
    
    bReturn = True
    If m.nMode = eGDManageMode_Ration Then
        If Dirty Then
            strResponse = InfBox("Do you want to save your changes?", "?", "+Yes|No|-Cancel", Caption)
            Select Case strResponse
                Case "C"
                    bReturn = False
                
                Case "Y"
                    bReturn = Save
                
                Case "N"
                
            End Select
        End If
    End If
    
    AskToSave = bReturn
        
ErrExit:
    Exit Function

ErrSection:
    AskToSave = True
    RaiseError "frmCattleEditor.AskToSave"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Save the changes to the object
'' Inputs:      None
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Save(Optional ByVal strTool As String = "ID_SAVE") As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim strText As String               ' Text for the InfBox
    Dim strTitle As String              ' Title for the InfBox
    Dim strNewName As String            ' New name selected from the InfBox
    Dim strRationID As String           ' Ration ID
    Dim lIndex As Long                  ' Index into a for loop
    Dim astrIngredientID As cGdArray    ' Array of ingredient IDs
    Dim astrPoundsFed As cGdArray       ' Array of pounds fed
    Dim astrPercentMarkup As cGdArray   ' Array of percent markup
    Dim FeedYard As cBrokerMessage      ' Selected feedyard

    bReturn = True
    If Validate Then
        If m.nMode = eGDManageMode_Ration Then
            bReturn = False
            
            If Len(m.strName) = 0 Then
                strText = "Save the current Ration as..."
                strTitle = "Save"
                strNewName = InfBox(strText, , , strTitle, , , , , , "string", m.strName)
            ElseIf UCase(strTool) = "ID_SAVEAS" Then
                strText = "Save the current Ration as..."
                strTitle = "Save As"
                strNewName = InfBox(strText, , , strTitle, , , , , , "string", m.strName)
            ElseIf UCase(strTool) = "ID_RENAME" Then
                strText = "Rename the current Ration as..."
                strTitle = "Rename"
                strNewName = InfBox(strText, , , strTitle, , , , , , "string", m.strName)
            Else
                strNewName = m.strName
            End If
            
            strRationID = m.Ration("ID")
            Do While (Len(strNewName) > 0) And (strNewName <> m.strName)
                If g.Cattle.RationNameExists(strNewName, strRationID) Then
                    InfBox "'" & strNewName & "' already exists.  Please select a new name", "!", , strTitle & " Error"
                Else
                    Exit Do
                End If
                
                strNewName = InfBox(strText, , , strTitle, , , , , , "string", m.strName)
            Loop
            
            If Len(strNewName) > 0 Then
                Set FeedYard = g.Cattle.SelectedFeedYard
                If FeedYard Is Nothing Then
                    m.Ration.Add "FeedYardID", ""
                Else
                    m.Ration.Add "FeedYardID", FeedYard("ID")
                End If
                
                If UCase(strTool) = "ID_SAVEAS" Then
                    m.Ration.Add "ID", ""
                End If
                m.Ration.Add "RationName", strNewName
                
                With fgObjects
                    Set astrIngredientID = New cGdArray
                    astrIngredientID.Create eGDARRAY_Strings, .Rows - 1
                    Set astrPoundsFed = New cGdArray
                    astrPoundsFed.Create eGDARRAY_Strings, .Rows - 1
                    Set astrPercentMarkup = New cGdArray
                    astrPercentMarkup.Create eGDARRAY_Strings, .Rows - 1
                
                    For lIndex = .FixedRows To .Rows - 1
                        If (RowIsClickHereLine(lIndex) = False) And (RowIsClickHereManageLine(lIndex) = False) Then
                            astrIngredientID(lIndex - .FixedRows) = .TextMatrix(lIndex, RationCol(eGDRationCols_IngredientID))
                            astrPoundsFed(lIndex - .FixedRows) = .TextMatrix(lIndex, RationCol(eGDRationCols_PoundsFed))
                            astrPercentMarkup(lIndex - .FixedRows) = .TextMatrix(lIndex, RationCol(eGDRationCols_PercentMarkup))
                        End If
                    Next lIndex
                End With
                
                m.Ration.Add "IngredientID", astrIngredientID.JoinFields("|")
                m.Ration.Add "PoundsFed", astrPoundsFed.JoinFields("|")
                m.Ration.Add "PercentMarkup", astrPercentMarkup.JoinFields("|")
                
                g.Cattle.UpdateRation m.Ration
                
                m.strName = strNewName
                SetEditorCaption Me, "Ration", strNewName
                bReturn = True
                Dirty = False
            End If
        End If
    End If
    
    Save = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCattleEditor.Save"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadIngredients
'' Description: Load up the ingredient names
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadIngredients()
On Error GoTo ErrSection:

    Dim Ingredients As cGdTree          ' Collection of ingredients
    Dim lIndex As Long                  ' Index into a for loop
    Dim Ingredient As cBrokerMessage    ' Ingredient object

    Set m.astrIngredients = New cGdArray
    m.astrIngredients.Create eGDARRAY_Strings
    
    Set Ingredients = g.Cattle.Ingredients
    For lIndex = 1 To Ingredients.Count
        Set Ingredient = Ingredients(lIndex)
        m.astrIngredients.Add Ingredient("Ingredient")
    Next lIndex
    
    fgObjects.ColComboList(RationCol(eGDRationCols_Ingredient)) = m.astrIngredients.JoinFields("|")

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleEditor.LoadIngredients"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Validate
'' Description: Validate the information in the grid
'' Inputs:      Show Message?
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Validate(Optional ByVal bShowMessage As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    bReturn = True
    With fgObjects
        For lIndex = .FixedRows To .Rows - 1
            If Len(.TextMatrix(lIndex, RationCol(eGDRationCols_Ingredient))) = 0 Then
                If bShowMessage Then
                    InfBox "You must select an ingredient", "!", , "Error"
                    EditCell fgObjects, lIndex, RationCol(eGDRationCols_Ingredient)
                End If
                
                bReturn = False
                Exit For
            End If
        Next lIndex
    End With
    
    Validate = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCattleEditor.Validate"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClickHereLine
'' Description: Determine the row in the grid that is the click here line
'' Inputs:      None
'' Returns:     Row of the click here line ( -1 if not found )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ClickHereLine() As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop

    lReturn = -1&
    With fgObjects
        For lIndex = .FixedRows To .Rows - 1
            If RowIsClickHereLine(lIndex) Then
                lReturn = lIndex
                Exit For
            End If
        Next lIndex
    End With

    ClickHereLine = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCattleEditor.ClickHereLine"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClickHereManageLine
'' Description: Determine the row in the grid that is the click here manage line
'' Inputs:      None
'' Returns:     Row of the click here line ( -1 if not found )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ClickHereManageLine() As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop

    lReturn = -1&
    With fgObjects
        For lIndex = .FixedRows To .Rows - 1
            If RowIsClickHereManageLine(lIndex) Then
                lReturn = lIndex
                Exit For
            End If
        Next lIndex
    End With

    ClickHereManageLine = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCattleEditor.ClickHereManageLine"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddClickHereLine
'' Description: Add the click here line if it doesn't exist
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddClickHereLine()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim strText As String               ' Text to put in the grid
    Dim lStartCol As Long               ' Starting column
    Dim lEndCol As Long                 ' Ending column

    If ClickHereLine = -1& Then
        With fgObjects
            nRedraw = .Redraw
            .Redraw = flexRDNone
            
            .Rows = .Rows + 1
            .MergeRow(.Rows - 1) = True
            
            Select Case m.nMode
                Case eGDManageMode_Ration
                    .TextMatrix(.Rows - 1, RationCol(eGDRationCols_IngredientID)) = "-1"
                    
                    strText = "Click here to add an ingredient to the ration"
                    lStartCol = RationCol(eGDRationCols_Ingredient)
                    lEndCol = RationCol(eGDRationCols_PercentMarkup)
                    
            End Select
            
            .Cell(flexcpText, .Rows - 1, lStartCol, .Rows - 1, lEndCol) = strText
            .Cell(flexcpForeColor, .Rows - 1, lStartCol, .Rows - 1, lEndCol) = vbBlue
            .Cell(flexcpFontUnderline, .Rows - 1, lStartCol, .Rows - 1, lEndCol) = True
            .Cell(flexcpAlignment, .Rows - 1, lStartCol, .Rows - 1, lEndCol) = flexAlignLeftCenter
            
            .Redraw = nRedraw
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleEditor.AddClickHereLine"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddClickHereManageLine
'' Description: Add the click here line if it doesn't exist
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddClickHereManageLine()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim strText As String               ' Text to put in the grid
    Dim lStartCol As Long               ' Starting column
    Dim lEndCol As Long                 ' Ending column

    If ClickHereManageLine = -1& Then
        With fgObjects
            nRedraw = .Redraw
            .Redraw = flexRDNone
            
            .Rows = .Rows + 1
            .MergeRow(.Rows - 1) = True
            
            Select Case m.nMode
                Case eGDManageMode_Ration
                    .TextMatrix(.Rows - 1, RationCol(eGDRationCols_IngredientID)) = "-2"
                    
                    strText = "Click here to manage ingredients"
                    lStartCol = RationCol(eGDRationCols_Ingredient)
                    lEndCol = RationCol(eGDRationCols_PercentMarkup)
                    
            End Select
            
            .Cell(flexcpText, .Rows - 1, lStartCol, .Rows - 1, lEndCol) = strText
            .Cell(flexcpForeColor, .Rows - 1, lStartCol, .Rows - 1, lEndCol) = vbBlue
            .Cell(flexcpFontUnderline, .Rows - 1, lStartCol, .Rows - 1, lEndCol) = True
            .Cell(flexcpAlignment, .Rows - 1, lStartCol, .Rows - 1, lEndCol) = flexAlignLeftCenter
            
            .Redraw = nRedraw
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleEditor.AddClickHereManageLine"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RowIsClickHereLine
'' Description: Determine if the given row in the grid is the click here line
'' Inputs:      Row
'' Returns:     True if click here line, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RowIsClickHereLine(ByVal lRow As Long) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    Select Case m.nMode
        Case eGDManageMode_Ration
            bReturn = (fgObjects.TextMatrix(lRow, RationCol(eGDRationCols_IngredientID)) = "-1")
    
    End Select
    
    RowIsClickHereLine = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCattleEditor.RowIsClickHereLine"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RowIsClickHereManageLine
'' Description: Determine if the given row in the grid is the click here manage line
'' Inputs:      Row
'' Returns:     True if click here line, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RowIsClickHereManageLine(ByVal lRow As Long) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    Select Case m.nMode
        Case eGDManageMode_Ration
            bReturn = (fgObjects.TextMatrix(lRow, RationCol(eGDRationCols_IngredientID)) = "-2")
    
    End Select
    
    RowIsClickHereManageLine = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCattleEditor.RowIsClickHereManageLine"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LastNonClickHereLine
'' Description: Last row in the grid that is not a click-here line
'' Inputs:      None
'' Returns:     Last non click here line ( -1 if none )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LastNonClickHereLine() As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    lReturn = -1&
    With fgObjects
        For lIndex = .Rows - 1 To .FixedRows Step -1&
            If (RowIsClickHereLine(lIndex) = False) And (RowIsClickHereManageLine(lIndex) = False) Then
                lReturn = lIndex
                Exit For
            End If
        Next lIndex
    End With

    LastNonClickHereLine = lReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCattleEditor.LastNonClickHereLine"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ValidDataRow
'' Description: Determine if the given row is a valid row
'' Inputs:      Row
'' Returns:     True if valid, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidDataRow(Optional ByVal Row As Long = kNullData) As Boolean
On Error GoTo ErrSection:

    Dim lLastValidRow As Long           ' Last valid data row in the grid

    If Row = kNullData Then
        Row = fgObjects.Row
    End If
    
    lLastValidRow = LastNonClickHereLine
    
    ValidDataRow = ((ValidGridRow(fgObjects, Row) = True) And (Row <= lLastValidRow))

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCattleEditor.ValidDataRow"
    
End Function

