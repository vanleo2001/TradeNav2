VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmLinkedRules 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2490
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6135
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "frmMDIChild"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   6135
   Begin VSFlex7LCtl.VSFlexGrid vsGrid 
      Height          =   1935
      Left            =   75
      TabIndex        =   3
      Top             =   435
      Width           =   5970
      _cx             =   10530
      _cy             =   3413
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
Begin HexUniControls.ctlUniButtonImageXP cmdCancel
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   900
      Picture         =   "frmLinkedRules.frx":0000
      TabIndex        =   2
      Top             =   60
      Width           =   750
   End
Begin HexUniControls.ctlUniButtonImageXP Corner
      Caption         =   "Corner"
      Height          =   240
      Left            =   4365
      TabIndex        =   1
      Top             =   2250
      Visible         =   0   'False
      Width           =   720
   End
Begin HexUniControls.ctlUniButtonImageXP cmdOK
      Caption         =   "&OK"
      Height          =   315
      Left            =   75
      TabIndex        =   0
      Top             =   60
      Width           =   750
   End
Begin HexUniControls.ctlUniLabelXP Label1
      Caption         =   "Select Entry Rules that apply to this Exit Rule."
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Top             =   75
      Width           =   3255
   End
End
Attribute VB_Name = "frmLinkedRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmLinkedRules.frm
'' Description: Allow the user to link specific entries to an exit
''
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    bOK As Boolean
End Type
Private m As mPrivate

' Grid Columns
Private Enum eGDCols
    eGDCol_Select = 0
    eGDCol_RuleName = 1
    eGDCol_RuleID = 2
End Enum
Private Const kGridCols = 3

Private Function GDCol(ByVal lColumn As eGDCols) As Long
    GDCol = lColumn
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: If the user hits OK, return the selected items
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    Dim lIndex As Long
    Dim bSomeChecked As Boolean
    
    bSomeChecked = False
    For lIndex = vsGrid.FixedRows To vsGrid.Rows - 1
        If CheckedCell(vsGrid, lIndex, GDCol(eGDCol_Select)) Then
            bSomeChecked = True
            Exit For
        End If
    Next lIndex
    
    If Not bSomeChecked Then
        Err.Raise vbObjectError + 1000, , "You must select at least one entry"
    End If

    m.bOK = True
    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLinkedRules.cmdOK.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: If the user hits Cancel, unload the form
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
    RaiseError "frmLinkedRules.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Form_Activate()

    On Error Resume Next
    If vsGrid.Rows > vsGrid.FixedRows Then
        vsGrid.Row = vsGrid.FixedRows
        MoveFocus Me.vsGrid
    End If

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
    RaiseError "frmLinkedRules.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    vsGrid_BeforeEdit
'' Description: Only allow the user to edit the Select column
'' Inputs:      Row and Column of Edited Cell, Whether or not to Cancel
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub vsGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    If Col > GDCol(eGDCol_Select) Then Cancel = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLinkedRules.vsGrid.BeforeEdit", eGDRaiseError_Show
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

    Dim lRedraw As Long                 ' Current state of the grid redraw

    With vsGrid
        lRedraw = .Redraw
        .Redraw = flexRDNone
        .Editable = flexEDKbdMouse
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .ScrollBars = flexScrollBarBoth
        .ExtendLastCol = True
        .Rows = 1
        .Cols = kGridCols
        .FixedCols = 0
        .FormatString = "Sel|Rule"
        ''.Cell(flexcpFontUnderline, 0, 0, 0, 1) = True
        .ColDataType(GDCol(eGDCol_Select)) = flexDTBoolean
        .ColHidden(GDCol(eGDCol_RuleID)) = True
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = lRedraw
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmLinkedRules.InitGrid", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and show the form
'' Inputs:      Location to show form, Rules collection, Exit clicked on,
''              Linked Entries to send back
'' Returns:     True if OK, False if Cancel
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ByVal ptX As Double, ByVal ptY As Double, Rules As cRules, lExitID As Long, strLink As String) As Boolean
On Error GoTo ErrSection:

    Dim X As Integer                    ' Index into a for loop
    Dim Y As Integer                    ' Index into a for loop
    Dim bFound As Boolean               ' Was the rule found?
    Dim Rule As cRule                   ' Temporary Rule variable
    Dim strTemp As String               ' Temporary string variable
    Dim astrRuleIDs As New cGdArray     ' List of linked Rule ID's
    Dim bAllChecked As Boolean          ' Are all of the entries checked?
    
    astrRuleIDs.Create eGDARRAY_Strings
    
    Move ptX, ptY
    
    vsGrid.Redraw = flexRDNone
    InitGrid
    
    ' Get the current linked entries for the current exit
    Set Rule = Rules.Item(CStr(lExitID))
    If Rule.LinkedRules <> "" Then
        strTemp = Rule.LinkedRules
        If Left(strTemp, 1) = "," Then
            strTemp = Right(strTemp, Len(strTemp) - 1)
        End If
        If Right(strTemp, 1) = "," Then
            strTemp = Left(strTemp, Len(strTemp) - 1)
        End If
        astrRuleIDs.SplitFields strTemp, ","
    End If
    
    ' Load grid with the appropriate entry rules.  If Selected rule on System
    ' Manager is a Long Exit then only (non-late) BUY entries are loaded
    For X = 1 To Rules.Count
        With Rules.Item(X)
            If .RuleUse = 0 And .RuleID <> lExitID Then
                If .BuySell <> Rule.BuySell Then
                    vsGrid.Rows = vsGrid.Rows + 1
                    CheckedCell(vsGrid, vsGrid.Rows - 1, GDCol(eGDCol_Select)) = True
                    vsGrid.TextMatrix(vsGrid.Rows - 1, GDCol(eGDCol_RuleName)) = Rules.Item(X).Name
                    vsGrid.TextMatrix(vsGrid.Rows - 1, GDCol(eGDCol_RuleID)) = Rules.Item(X).RuleID
                End If
            End If
        End With
    Next X
    
    ' Check if there are any entries to link
    If vsGrid.Rows = vsGrid.FixedRows Then
        If Rule.BuySell Then
            Err.Raise vbObjectError + 1000, , "There are no Short Entry Rules for this Short Exit Rule"
        Else
            Err.Raise vbObjectError + 1000, , "There are no Long Entry Rules for this Long Exit Rule"
        End If
    End If
    
    With vsGrid
        ' If Linked RuleID's already exist, overlay the current values into grid
        If astrRuleIDs.Size > 0 Then
            For X = .FixedRows To .Rows - 1
                bFound = False
                For Y = 0 To astrRuleIDs.Size - 1
                    If astrRuleIDs(Y) = .TextMatrix(X, GDCol(eGDCol_RuleID)) Then
                        bFound = True
                        Exit For
                    End If
                Next Y
                
                CheckedCell(vsGrid, X, GDCol(eGDCol_Select)) = bFound
            Next X
        End If
    
        ' Sort the grid by Rule Name
        .Col = GDCol(eGDCol_RuleName)
        .Sort = flexSortGenericAscending
        .Redraw = flexRDBuffered
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With
    
    ShowForm Me, True, , , ALT_GRID_ROW_COLOR
    
    If m.bOK = True Then
        With vsGrid
            bAllChecked = True
            For X = .FixedRows To .Rows - 1
                If CheckedCell(vsGrid, X, GDCol(eGDCol_Select)) = False Then
                    bAllChecked = False
                    Exit For
                End If
            Next X
            
            ' If all of the entries are checked, send back an empty string since the
            ' exit is not linked to any particular entries...
            If bAllChecked Then
                strLink = ""
                
            ' Otherwise send back a comma delimited string of the Rule ID's of the
            ' entries that it is linked to...
            Else
                strLink = ","
                For X = .FixedRows To .Rows - 1
                    If CheckedCell(vsGrid, X, GDCol(eGDCol_Select)) Then
                        strLink = strLink & .TextMatrix(X, GDCol(eGDCol_RuleID)) & ","
                    End If
                Next X
            End If
        End With
    End If
    
ErrExit:
    Set Rule = Nothing
    ShowMe = m.bOK
    Unload Me
    Exit Function
    
ErrSection:
    RaiseError "frmLinkedRules.ShowMe", eGDRaiseError_Raise

End Function

