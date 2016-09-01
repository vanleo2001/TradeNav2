VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmSelect 
   Caption         =   "Symbol Groups: to store a list of specific symbols."
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4980
   Icon            =   "frmSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   4980
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniRadioXP optExisting 
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   2775
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
      Caption         =   "frmSelect.frx":0442
      Enabled         =   -1  'True
      Align           =   0
      RadioBackColor  =   -2147483643
      RadioForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   0   'False
      Tip             =   "frmSelect.frx":04A2
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmSelect.frx":04C2
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniRadioXP optNew 
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   2775
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
      Caption         =   "frmSelect.frx":04DE
      Enabled         =   -1  'True
      Align           =   0
      RadioBackColor  =   -2147483643
      RadioForeColor  =   -2147483640
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Pressed         =   -1  'True
      Tip             =   "frmSelect.frx":0530
      Style           =   -1
      MousePointer    =   0
      MouseIcon       =   "frmSelect.frx":0550
      ShowFocus       =   -1  'True
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   2715
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   1155
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
      Caption         =   "frmSelect.frx":056C
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmSelect.frx":05A0
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmSelect.frx":05C0
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdRename 
         Height          =   435
         Left            =   60
         TabIndex        =   5
         Top             =   1440
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
         Caption         =   "frmSelect.frx":05DC
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmSelect.frx":060A
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmSelect.frx":062A
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdNew 
         Height          =   435
         Left            =   60
         TabIndex        =   3
         Top             =   480
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
         Caption         =   "frmSelect.frx":0646
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmSelect.frx":066E
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmSelect.frx":068E
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdEdit 
         Default         =   -1  'True
         Height          =   435
         Left            =   60
         TabIndex        =   2
         Top             =   0
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
         Caption         =   "frmSelect.frx":06AA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmSelect.frx":06D4
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmSelect.frx":06F4
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdClose 
         Cancel          =   -1  'True
         Height          =   435
         Left            =   60
         TabIndex        =   6
         Top             =   2160
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
         Caption         =   "frmSelect.frx":0710
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmSelect.frx":073C
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmSelect.frx":075C
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdDelete 
         Height          =   435
         Left            =   60
         TabIndex        =   4
         Top             =   960
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
         Caption         =   "frmSelect.frx":0778
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmSelect.frx":07A6
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmSelect.frx":07C6
         RightToLeft     =   0   'False
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgList 
      Height          =   2235
      Left            =   60
      TabIndex        =   0
      Top             =   1080
      Width           =   1455
      _cx             =   2566
      _cy             =   3942
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
      GridColor       =   -2147483643
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   0
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
End
Attribute VB_Name = "frmSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const kExtendedCol = 1

Public Enum eSelectMode
    eSelectMode_Select = 0
    eSelectMode_Edit = 1
    eSelectMode_SendTo = 2
End Enum

Private Enum eGDCols
    eGDCol_Active = 0
    eGDCol_Name = 1
    eGDCol_ID = 2
    eGDCol_Desc = 3
    eGDCol_NumDays = 4
End Enum
Private Const kNumCols = 5

Private Type mPrivate
    strPath As String
    strType As String
    SelectMode As eSelectMode
    alSymbolIds As cGdArray
    strReturn As String
    lColSorted As Long
    lOrder As Long

    lPrevColWidth As Long               ' Used for Extend custom column
End Type
Private m As mPrivate

Private Function GDCol(ByVal Col As eGDCols) As Long
    GDCol = Col
End Function

Private Sub cmdClose_Click()
On Error GoTo ErrSection:

    Me.Hide
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSelect.cmdClose.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrSection:

    Dim strID$, strMsg$, SymbolGroup As New cSymbolGroup

    strMsg = fgList.TextMatrix(fgList.Row, GDCol(eGDCol_Name))
    strMsg = "Are you sure you want to delete:|" & strMsg
    If AskBox("i=? ; h=Confirm Delete ; b=+Delete|-Cancel ; " & strMsg) = "C" Then
        Exit Sub
    End If

    If fgList.Row >= fgList.FixedRows And fgList.Row < fgList.Rows Then
        strID = fgList.RowData(fgList.Row)
    End If
    If Len(strID) > 0 Then
        KillFile App.Path & "\Custom\" & strID

        With g.SymbolPool
            Select Case UCase(m.strType)
                Case "GRP":
                    Set SymbolGroup = .SymbolGroups(strID)
                    If Not SymbolGroup Is Nothing Then
                        If SymbolGroup.IsIndex Then
                            SU_DeleteComposite SymbolGroup.SymbolID, UCase("#" & SymbolGroup.Name)
                            g.SymbolPool.RemoveCustomIndex SymbolGroup.SymbolID
                            frmSymbolGrid.RefreshGrid
                        End If
                        .SymbolGroups.Remove strID
                    End If
                Case "FIL":
                    .Filters.Remove strID
                Case "SCN":
                    .Criterias.Remove strID
            End Select
            .RemoveOrphanedArraysFromTable
        End With

        FillList
        
        frmSymbolGrid.LoadCombo
    End If
    
ErrExit:
    Set SymbolGroup = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmSelect.cmdDelete.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cmdEdit_Click()
On Error GoTo ErrSection:

    Dim strID$
    
    If m.strType = "FIL" Then
        If fgList.Cell(flexcpChecked, fgList.Row, GDCol(eGDCol_Active)) = flexUnchecked Then Exit Sub
    End If
    
    ' Need to do this for "Default" buttons, since hitting 'Enter'
    ' from another control does not trigger it's 'LostFocus' event.
    MoveFocus cmdEdit

    If fgList.Row >= fgList.FixedRows And fgList.Row < fgList.Rows Then
        strID = fgList.RowData(fgList.Row)
    End If
    
    Select Case m.SelectMode
        Case eSelectMode_Select
            m.strReturn = strID
            Me.Hide
        Case eSelectMode_Edit
            If Len(strID) > 0 Then
                Edit strID
            Else
                Beep
            End If
        Case eSelectMode_SendTo
            If optNew.Value = True Then
                SendTo ""
            Else
                SendTo strID
            End If
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSelect.cmdEdit.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cmdNew_Click()
On Error GoTo ErrSection:

    Dim aFiles As New cGdArray
    
    Select Case UCase(m.strType)
        Case "SCN":
            If gdNumMatchingFiles(App.Path & "\Custom\Cus0*.SCN") >= 5 Then
                If Not HasGold(True, "Creating more custom Criteria") Then
                    Exit Sub
                End If
            End If
    End Select
    
    Edit ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSelect.cmdNew.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cmdRename_Click()
On Error GoTo ErrSection:

    Dim obj As Object
    Dim strReturn As String
    Dim strText As String
    
    With fgList
        Select Case m.strType
            Case "GRP"
                strText = "Rename current Symbol Group as..."
            Case "SCN"
                strText = "Rename current Criteria as..."
            Case "FIL"
                strText = "Rename current Filter as..."
        End Select
        
        strReturn = AskBox("h=Rename ; i=? ; g=string ; d=" & .Cell(flexcpText, .Row, GDCol(eGDCol_Name)) & " ; " & strText)
        If strReturn <> "" And strReturn <> .Cell(flexcpText, .Row, GDCol(eGDCol_Name)) Then
            Set obj = g.SymbolPool.PoolObject(m.strType & ":" & .RowData(.Row))
            If Not obj Is Nothing Then
                If m.strType = "GRP" Then obj.FromFile AddSlash(App.Path) & "Custom", obj.ID, True
                obj.Name = strReturn
                obj.ToFile
                obj.AddToPool ' to replace the name in the fields table
                frmSymbolGrid.RefreshGrid
                .Cell(flexcpText, .Row, GDCol(eGDCol_Name)) = strReturn
                Set obj = Nothing
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSelect.cmdRename.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    EnableButtons

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSelect.fgList.AfterEdit", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    EnableButtons
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSelect.fgList.AfterRowColChange", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgList_AfterSort(ByVal Col As Long, Order As Integer)
On Error GoTo ErrSection:

    m.lColSorted = Col
    m.lOrder = Order

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSelect.fgList.AfterSort", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim lWidth As Long                  ' Amount to adjust the column width
    Dim lIndex As Long                  ' Index into a for loop
    
    ' if column being resized is the extended column,
    ' then make the next column bigger (instead of adjusting
    ' the extended column)
    If Col >= kExtendedCol Then
        With fgList
            .Redraw = flexRDNone
            lWidth = .ColWidth(Col) - m.lPrevColWidth
            For lIndex = Col + 1 To .Cols - 1
                If Not .ColHidden(lIndex) Then
                    .ColWidth(lIndex) = .ColWidth(lIndex) - lWidth
                    Exit For
                End If
            Next
            m.lPrevColWidth = 0
            ExtendCustomColumn
            .Redraw = flexRDBuffered
        End With
   Else
        ExtendCustomColumn
   End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSelect.fgList.AfterUserResize", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    If Col <> GDCol(eGDCol_Active) Then
        Cancel = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSelect.fgList.BeforeEdit", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgList_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    ' if column being resized is the extended column, save size
    If Col >= kExtendedCol Then
        m.lPrevColWidth = fgList.ColWidth(Col)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSelect.fgList.BeforeUserResize", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgList_DblClick()
On Error GoTo ErrSection:

    Dim nRow&
    
    With fgList
        nRow = .MouseRow
        If nRow >= .FixedRows And nRow < .Rows Then
            cmdEdit_Click
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSelect.fgList.DblClick", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim nRow&
    
    With fgList
        nRow = .MouseRow
        If nRow >= .FixedRows And nRow < .Rows Then
            If .MouseCol = GDCol(eGDCol_NumDays) Then
                .ToolTipText = "Amount of data required to load when recalculate criteria"
            Else
                .ToolTipText = TipStr(.TextMatrix(nRow, GDCol(eGDCol_Desc)))
            End If
        Else
            .ToolTipText = ""
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSelect.fgList.MouseMove", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgList_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim astrInactiveCriteria As New cGdArray
    Dim Filter As New cFilter
    Dim Criteria As New cCriteria
    Dim strMessage As String
    Dim lIndex As Long
    Dim bAsked As Boolean
    Dim strReturn As String

    astrInactiveCriteria.Create eGDARRAY_Strings

    If m.strType = "FIL" Then
        With fgList
            Set Filter = g.SymbolPool.Filters(.Cell(flexcpText, Row, GDCol(eGDCol_ID)))
            If Val(.EditText) = flexChecked Then
                Set astrInactiveCriteria = Filter.InactiveCriteria
                If astrInactiveCriteria.Size > 0 Then
                    If astrInactiveCriteria.Size = 1 Then
                        strMessage = "The following criteria used by this filter|is currently inactive:||"
                    Else
                        strMessage = "The following criteria used by this filter|are currently inactive:||"
                    End If
                    For lIndex = 0 To astrInactiveCriteria.Size - 1
                        strMessage = strMessage & Parse(astrInactiveCriteria(lIndex), "|", 2) & "|"
                    Next lIndex
                    If astrInactiveCriteria.Size = 1 Then
                        strMessage = strMessage & "|This criteria will be activated by activating the filter.|"
                    Else
                        strMessage = strMessage & "|These criteria will be activated by activating the filter.|"
                    End If
                    
                    If InfBox(strMessage, "!", "+OK|-Cancel", "Warning") = "C" Then
                        Cancel = True
                        astrInactiveCriteria.Destroy
                        Exit Sub
                    Else
                        For lIndex = 0 To astrInactiveCriteria.Size - 1
                            Set Criteria = g.SymbolPool.Criterias(Parse(astrInactiveCriteria(lIndex), "|", 1))
                            Criteria.IsActive = True
                            Criteria.ToFile
                        Next lIndex
                    End If
                End If
            End If
            Filter.IsActive = (CLng(.EditText) = flexChecked)
            Filter.ToFile
        End With
    ElseIf m.strType = "SCN" Then
        With fgList
            Set Criteria = g.SymbolPool.Criterias(.Cell(flexcpText, Row, GDCol(eGDCol_ID)))
            If Val(.EditText) = flexUnchecked Then
                bAsked = False
                For Each Filter In g.SymbolPool.Filters
                    If Filter.IsActive Then
                        If Filter.CriteriaInFilter(Criteria.ID) = True Then
                            If bAsked = False Then
                                strReturn = AskBox("h=Criteria ; i=? ; b=+Yes|-No ; " & _
                                    "There are active filters that are using||" & Criteria.Name & _
                                    "||Deactivating this criteria will also" & _
                                    "|deactivate the filters that use this criteria.||" & _
                                    "Are you sure you want to do this?|")
                                bAsked = True
                                If strReturn = "N" Then
                                    Cancel = True
                                    Exit For
                                End If
                            End If
                            Filter.IsActive = False
                            Filter.ToFile
                        End If
                    End If
                Next Filter
            End If
        
            If Not Cancel Then
                Criteria.IsActive = (CLng(.EditText) = flexChecked)
                Criteria.ToFile
            End If
        End With
    End If
    
ErrExit:
    Set astrInactiveCriteria = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmSelect.fgList.ValidateEdit", eGDRaiseError_Show
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
    RaiseError "frmSelect.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    CenterTheForm Me
    
    g.Styler.StyleForm Me
    
    m.lColSorted = -1&
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSelect.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Public Function ShowMe(ByVal strType$, _
                        Optional SelectMode As eSelectMode = eSelectMode_Edit, _
                        Optional alSymbolIds As cGdArray = Nothing) As String
On Error GoTo ErrSection:

    Dim i&
    Dim Criteria As New cCriteria
    Dim Filter As New cFilter
    Dim bAsked As Boolean
    Dim strReturn As String
    
    Screen.MousePointer = vbHourglass

    ' Set the module level variables
    m.strType = UCase(strType)
    Set m.alSymbolIds = alSymbolIds
    m.SelectMode = SelectMode
    m.strPath = App.Path & "\Custom\"
    
    Select Case UCase(m.strType)
        Case "GRP":
            Me.Caption = "Custom SYMBOL GROUPS"
            Me.Icon = Picture16(ToolbarIcon("ID_SymbolGroups"), , True)
        Case "SCN":
            Me.Caption = "Custom CRITERIA"
            Me.Icon = Picture16(ToolbarIcon("ID_Criteria"), , True)
        Case "FIL":
            Me.Caption = "Custom FILTERS"
            Me.Icon = Picture16(ToolbarIcon("ID_Filters"), , True)
    End Select

    FillList

    ' Set up the user interface appropriately
    Select Case m.SelectMode
        Case eSelectMode_Edit
            optNew.Visible = False
            optExisting.Visible = False
        Case eSelectMode_Select
            optNew.Visible = False
            optExisting.Visible = False
            
            cmdDelete.Visible = False
            cmdRename.Visible = False
            cmdNew.Visible = False
            cmdClose.Top = cmdNew.Top
            cmdEdit.Caption = "&Select"
            cmdClose.Caption = "&Cancel"
        Case eSelectMode_SendTo
            cmdDelete.Visible = False
            cmdRename.Visible = False
            cmdNew.Visible = False
            cmdClose.Top = cmdNew.Top
            cmdEdit.Caption = "&OK" ' "&Send"
            cmdClose.Caption = "&Cancel"
            
            If fgList.Rows = 0 Then Disable optExisting
    End Select

    m.strReturn = ""
    Screen.MousePointer = vbDefault
    ShowForm Me, True
    
#If 0 Then
    If m.strType = "SCN" Then
        For i = fgList.FixedRows To fgList.Rows - 1
            Set Criteria = g.SymbolPool.Criterias(fgList.Cell(flexcpText, i, GDCol(eGDCol_ID)))
            Criteria.IsActive = (fgList.Cell(flexcpChecked, i, GDCol(eGDCol_Active)) = flexChecked)
            
            ' If the criteria was turned off, turn off any filters that use it
            bAsked = False
            If Criteria.IsActive = False Then
                For Each Filter In g.SymbolPool.Filters
                    If Filter.IsActive Then
                        If Filter.CriteriaInFilter(Criteria.ID) = True Then
                            If bAsked = False Then
                                strReturn = AskBox("h=Criteria ; i=? ; b=+Yes|-No ; " & _
                                    "There are active filters that are using||" & Criteria.Name & _
                                    "||Deactivating this criteria will also" & _
                                    "|deactivate the filters that use this criteria.||" & _
                                    "Are you sure you want to do this?|")
                                bAsked = True
                                If strReturn = "N" Then
                                    fgList.Cell(flexcpChecked, i, GDCol(eGDCol_Active)) = flexChecked
                                    Exit For
                                End If
                            End If
                            Filter.IsActive = False
                            Filter.ToFile
                        End If
                    End If
                Next Filter
            End If
            
            Criteria.IsActive = (fgList.Cell(flexcpChecked, i, GDCol(eGDCol_Active)) = flexChecked)
            Criteria.ToFile
        Next i
    End If
#End If
    
    'Me.Refresh
    'DoEvents
    
    ShowMe = m.strReturn

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmSelect.ShowMe", eGDRaiseError_Raise
    
End Function

Private Sub FillList(Optional ByVal strID_Select$ = "")
On Error GoTo ErrSection:

    Dim i&, strActive$, strID$, strName$, strDesc$, strNumDays$
    Dim iSelect&, iOldTop&
    Dim aNames As New cGdArray
    Dim Objects As cGdTree
    Dim bAdvanced As Boolean
    
    ' get the appropriate collection
    Select Case m.strType
        Case "GRP":
            Set Objects = g.SymbolPool.SymbolGroups
        Case "SCN":
            Set Objects = g.SymbolPool.Criterias
        Case "FIL":
            Set Objects = g.SymbolPool.Filters
    End Select
    
    ' get list of names and id's for each object
    For i = 1 To Objects.Count
        If Objects(i).Custom Or m.SelectMode = eSelectMode_Select Then
            If Objects(i).IsActive Then strActive = CStr(vbChecked) Else strActive = CStr(vbUnchecked)
            strID = Objects(i).ID
            strName = Objects(i).Name
            strDesc = Objects(i).Desc
            If m.strType = "SCN" Then
                If Objects(i).IsWeekly = True Then
                    strNumDays = Str(Objects(i).NumDays * 5)
                Else
                    strNumDays = Str(Objects(i).NumDays)
                End If
            Else
                strNumDays = ""
            End If
            If UCase(strID) <> "_FLAGS_.GRP" And UCase(strName) <> "ALL SYMBOLS" Then
                aNames.Add strActive & vbTab & strName & vbTab & strID & vbTab & strDesc & vbTab & strNumDays
            End If
        End If
    Next
    aNames.Sort eGdSort_IgnoreCase + 2
    
    ' add to grid
    With fgList
        .Redraw = flexRDNone
        iOldTop = .TopRow
        InitGrid

        If m.strType = "GRP" Then
            .ColHidden(GDCol(eGDCol_Active)) = True
        End If
        
        'bAdvanced = CBool(GetIniFileProperty("Advanced", vbUnchecked, "Criteria", App.Path & "\ChartNavigator.ini"))
        If m.strType <> "SCN" Then 'Or bAdvanced = False Then
            .ColHidden(GDCol(eGDCol_NumDays)) = True
        ElseIf m.strType = "SCN" Then
            .ColHidden(GDCol(eGDCol_NumDays)) = False
        End If
        
        iSelect = .FixedRows '(default)
        For i = 0 To aNames.Size - 1
            'show name
            fgList.AddItem aNames(i)
            'also put ID into RowData
            strID = Parse(aNames(i), vbTab, 3)
            If UCase(strID) = UCase(strID_Select) Then iSelect = fgList.Rows - 1
            fgList.RowData(fgList.Rows - 1) = strID
        Next
        If iSelect < .Rows Then
            ' try to set "the way it shows" back to the way it was
            If iOldTop < .Rows Then .TopRow = iOldTop
            .Row = iSelect
            .RowSel = iSelect
            .ShowCell iSelect, 0
        End If
                
        .ColPosition(GDCol(eGDCol_NumDays)) = 0
        .AutoSize 0, kNumCols - 1, False, 75
        i = .ColWidth(2)
        .ColPosition(0) = GDCol(eGDCol_NumDays)
        .ColWidth(GDCol(eGDCol_Name)) = i - 400
        ExtendCustomColumn
        
        If m.lColSorted <> -1& Then
            .ColSort(m.lColSorted) = m.lOrder
            .Select .FixedRows, 0, .Rows - 1, .Cols - 1
            .Sort = flexSortUseColSort
            For i = .FixedRows To .Rows - 1
                If .RowData(i) = UCase(strID_Select) Then
                    .RowSel = i
                    .Row = i
                    .ShowCell i, 0
                    Exit For
                End If
            Next i
        End If
        
        .Redraw = flexRDBuffered
    End With

    EnableButtons

ErrExit:
    Set Objects = Nothing
    Set aNames = Nothing
    Exit Sub
    
ErrSection:
    Set Objects = Nothing
    Set aNames = Nothing
    RaiseError "frmSelect.FillList", eGDRaiseError_Raise
    
End Sub

Private Sub InitGrid()
On Error GoTo ErrSection:

    With fgList
        .Redraw = flexRDNone
        .FixedCols = 0
        .FixedRows = 1
        .AllowUserResizing = flexResizeColumns
        .ExplorerBar = flexExSortShowAndMove
        .ScrollTrack = True
        '.SelectionMode = flexSelectionFree
        .SelectionMode = flexSelectionListBox
        .AllowSelection = True 'False
        '.AllowUserFreezing = flexFreezeColumns
        .SheetBorder = RGB(128, 128, 128)
        .ExtendLastCol = False 'True
        .Editable = flexEDKbdMouse

        If m.strType <> "GRP" Then
            .BackColorBkg = g.Styler.GetColor(eGrid_Background) 'RH override vbApplicationWorkspacevbApplicationWorkspace
            .GridLinesFixed = flexGridInset
        Else
            .RowHidden(0) = True
            .BackColorBkg = g.Styler.GetColor(eGrid_Background) 'RH override vbApplicationWorkspacevbWindowBackground
            .GridLinesFixed = flexGridNone
        End If
        
        .Rows = .FixedRows
        .Cols = kNumCols
        .ColDataType(GDCol(eGDCol_Active)) = flexDTBoolean
        
        .ColHidden(GDCol(eGDCol_ID)) = True
        .ColHidden(GDCol(eGDCol_Desc)) = True
        If m.SelectMode <> eSelectMode_Edit Then .ColHidden(GDCol(eGDCol_Active)) = True
        
        .Cell(flexcpText, 0, GDCol(eGDCol_Active)) = "Active"
        .Cell(flexcpText, 0, GDCol(eGDCol_Name)) = "Name"
        .Cell(flexcpText, 0, GDCol(eGDCol_NumDays)) = "Days"
        
        '.ColWidth(GDCol(eGDCol_Active)) = 600
        '.ColWidth(GDCol(eGDCol_NumDays)) = 600
        '.ColWidth(GDCol(eGDCol_Name)) = .Width - .ColWidth(GDCol(eGDCol_Active)) - .ColWidth(GDCol(eGDCol_NumDays)) - 4 * Screen.TwipsPerPixelX
        
        ''.TextMatrix(0, GDCol(eGDCol_Name)) = "Template"
        '.FillStyle = flexFillRepeat
        '.Select 0, 0, 0, .Cols - 1
        '.CellFontBold = True
        '.CellForeColor = fraAppearance.ForeColor
        '.AutoSize 1
        '.ExtendLastCol = True
        '.ColAlignment(GDCol(eGDCol_ID)) = flexAlignCenterCenter

        '.TextMatrix(1, GDCol(eGDCol_Name)) = "Testing"
        '.Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSelect.InitGrid", eGDRaiseError_Raise
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode = 0 Then
        Cancel = True
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSelect.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Resize()
On Error Resume Next

    Dim w&, h&

    'check minimum size
    w = fraButtons.Width * 3
    h = fraButtons.Top + fraButtons.Height + Me.Height - Me.ScaleHeight
    If LimitFormSize(Me, w, h) Then Exit Sub

    fraButtons.Left = Me.ScaleWidth - fraButtons.Width
    
    If optNew.Visible = True Then
        With fgList
            .Move fraButtons.Width, optExisting.Top + optExisting.Height, _
                            fraButtons.Left - fraButtons.Width - optNew.Left, _
                            Me.ScaleHeight - optExisting.Top - optExisting.Height - optNew.Top
            ExtendCustomColumn
        End With
    Else
        With fgList
            .Move .Left, fraButtons.Top, fraButtons.Left - .Left * 2, _
                            Me.ScaleHeight - fraButtons.Top - .Left
            ExtendCustomColumn
        End With
    End If

    Me.Refresh

End Sub

Private Sub EnableButtons()
On Error GoTo ErrSection:

    Dim strText$, bIsActive As Boolean
    Static bInHere As Boolean

    If bInHere Then Exit Sub
    bInHere = True
    
    bIsActive = True
    With fgList
        If m.SelectMode = eSelectMode_SendTo Then
            If .Rows = .FixedRows Then
                optNew = True
            End If
            If optExisting Then
                .Enabled = True
                .HighLight = flexHighlightAlways
                If .Row < .FixedRows Then
                    .Row = .FixedRows
                End If
            Else
                .Enabled = False
                .HighLight = flexHighlightNever
                .Row = -1
            End If
        End If
    
        If .Row >= .FixedRows And .Row < .Rows Then
            strText = Trim(.TextMatrix(.Row, GDCol(eGDCol_Name)))
            If .Cell(flexcpChecked, .Row, GDCol(eGDCol_Active)) = flexUnchecked Then
                bIsActive = False
            End If
        End If
        
        'If m.strType = "GRP" And .Rows > .FixedRows Then
        '    If g.SymbolPool.PoolObject(m.strType & ":" & .RowData(.Row)).GroupType = eGROUP_QuoteList Then
        '        Disable cmdRename
        '    Else
                Enable cmdRename
        '    End If
        'End If
    End With

    If Len(strText) = 0 Then
        Disable cmdDelete
        If m.SelectMode <> eSelectMode_SendTo Then Disable cmdEdit
        Disable cmdRename
    Else
        Enable cmdDelete
        If m.strType = "FIL" And bIsActive = False Then
            Disable cmdEdit
        Else
            Enable cmdEdit
        End If
    End If

ErrExit:
    bInHere = False
    Exit Sub
    
ErrSection:
    bInHere = False
    RaiseError "frmSelect.EnableButtons", eGDRaiseError_Raise
    
End Sub

Private Sub Edit(ByVal strID$)
On Error GoTo ErrSection:

    Dim strPath$
    Dim obj As Object

    Select Case m.strType
        Case "GRP":
            Set obj = New cSymbolGroup
        Case "SCN":
            Set obj = New cCriteria
        Case "FIL":
            Set obj = New cFilter
        Case Else:
            Beep
            Exit Sub
    End Select
    
    strPath = App.Path & "\Custom\"
    obj.Edit strPath, strID
    
    strID = obj.ID '(in case was a new one)
    Set obj = Nothing

    FillList strID
    frmSymbolGrid.LoadCombo

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSelect.Edit", eGDRaiseError_Raise
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    'frmSymbolGrid.LoadCombo

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSelect.Form.Unload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub SendTo(ByVal strID As String)
On Error GoTo ErrSection:

    Dim strReturn As String
    Dim lIndex As Long
    Dim frm As New frmSymbolGroup
    Dim bSaveNew As Boolean
    
    If m.strType <> "GRP" Then
        ' ??? Should never get here
        Exit Sub
    End If
    
    If strID <> "" Then
        strReturn = InfBox("h=Send To Group ; i=? ; b=+Append|Overwrite|-Cancel ; " & _
                            "Do you want to Append to this group or Overwrite it?")
        If strReturn = "C" Then Exit Sub
    End If
    
    ' DAJ 2/15/2002: Do not allow industry sectors to be added to the quote board
    If strID = "QUOTELIST.GRP" Then
        For lIndex = m.alSymbolIds.Size - 1 To 0 Step -1
            If Left(g.SymbolPool.SymbolForID(m.alSymbolIds(lIndex)), 2) = "$-" Then
                m.alSymbolIds.Remove lIndex
            End If
        Next lIndex
    End If
    
    Me.Hide
    Select Case UCase(strReturn)
        Case "A":
            ' Append to the current Symbol Group
            ''SymbolGroup.Edit AddSlash(App.Path) & "Custom", strID, , , m.alSymbolIDs, True, True
            frm.ShowMe AddSlash(App.Path) & "Custom\", strID, False, m.alSymbolIds, True
        Case "O":
            ' Overwrite the current Symbol Group
            ''SymbolGroup.Edit AddSlash(App.Path) & "Custom", strID, , , m.alSymbolIDs, False, True
            frm.ShowMe AddSlash(App.Path) & "Custom\", strID, False, m.alSymbolIds, False
        Case Else:
            ' Create new Symbol Group with the Symbol ID's passed in
            ''SymbolGroup.Edit AddSlash(App.Path) & "Custom", strID, , , m.alSymbolIDs
            bSaveNew = optNew.Value
            frm.ShowMe AddSlash(App.Path) & "Custom\", strID, False, m.alSymbolIds, False, , bSaveNew
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSelect.SendTo", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExtendCustomColumn
'' Description: Adjust all column widths to accomodate the custom "extend column"
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ExtendCustomColumn()
On Error GoTo ErrSection:

    Dim lTotal As Long                  ' New width of the extended column
    Dim lIndex As Long                  ' Index into a for loop
        
    With fgList
        .ColHidden(kExtendedCol) = True
        .Redraw = flexRDBuffered '(so .ClientWidth will be correct)
        .Redraw = flexRDNone
        lTotal = 0 * Screen.TwipsPerPixelX
        For lIndex = 0 To .Cols - 1
            If Not .ColHidden(lIndex) Then
                lTotal = lTotal + .ColWidth(lIndex)
            End If
        Next
        lTotal = .ClientWidth - lTotal
        If lTotal > 0 Then .ColWidth(kExtendedCol) = lTotal
        .ColHidden(kExtendedCol) = False
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSelect.ExtendCustomColumn", eGDRaiseError_Raise
    
End Sub

Private Sub optExisting_Click()
On Error GoTo ErrSection:
    
    EnableButtons
    
    Exit Sub
ErrSection:
    RaiseError Me.Name & ".optExisting_Click"
End Sub

Private Sub optNew_Click()
On Error GoTo ErrSection:
    
    EnableButtons
    
    Exit Sub
ErrSection:
    RaiseError Me.Name & ".optNew_Click"
End Sub

