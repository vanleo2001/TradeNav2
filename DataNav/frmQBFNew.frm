VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmQBFNew 
   Caption         =   "Add Criteria to the Quote Board"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4830
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniRichTextBoxXP rtbPreview 
      Height          =   795
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2400
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   1402
      BackColor       =   -2147483633
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmQBFNew.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   -1
      MultiLine       =   -1  'True
      Alignment       =   0
      ScrollBars      =   3
      PasswordChar    =   ""
      TrapTab         =   0   'False
      RaiseChangeEvent=   -1  'True
      RaiseUpdateEvent=   0   'False
      RaiseSelChangeEvent=   -1  'True
      Tip             =   "frmQBFNew.frx":0020
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmQBFNew.frx":0040
      ViewMode        =   0
      TextModeText    =   2
      TextModeUndoLevel=   8
      TextModeCodePage=   32
      AutoURLDetect   =   0   'False
      FileName        =   ""
      VerticalLayout  =   0   'False
      OnlyNumbers     =   0   'False
      NoIME           =   0   'False
      SelfIME         =   0   'False
      LanguageOptions =   150
      RaiseRequestResizeEvent=   0   'False
      RaiseMsgFilterEvent=   0   'False
      SubClassPaintMessage=   0   'False
      TabSize         =   4
      TypographyOptions=   0
      BlockAutoCopy   =   0   'False
      BlockAutoCut    =   0   'False
      BlockAutoPaste  =   0   'False
      BlockAutoUndo   =   0   'False
   End
   Begin VSFlex7LCtl.VSFlexGrid fgList 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      _cx             =   5106
      _cy             =   3201
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
      Height          =   1935
      Left            =   3180
      TabIndex        =   6
      Top             =   120
      Width           =   1335
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
      Caption         =   "frmQBFNew.frx":005C
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmQBFNew.frx":0088
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmQBFNew.frx":00A8
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdEdit 
         Height          =   435
         Left            =   0
         TabIndex        =   5
         Top             =   1500
         Width           =   1335
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
         Caption         =   "frmQBFNew.frx":00C4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmQBFNew.frx":0100
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmQBFNew.frx":0120
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdNew 
         Height          =   435
         Left            =   0
         TabIndex        =   4
         Top             =   1020
         Width           =   1335
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
         Caption         =   "frmQBFNew.frx":013C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmQBFNew.frx":0176
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmQBFNew.frx":0196
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Height          =   435
         Left            =   0
         TabIndex        =   3
         Top             =   480
         Width           =   1335
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
         Caption         =   "frmQBFNew.frx":01B2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmQBFNew.frx":01E0
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmQBFNew.frx":0200
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Height          =   435
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1335
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
         Caption         =   "frmQBFNew.frx":021C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmQBFNew.frx":0242
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmQBFNew.frx":0262
         RightToLeft     =   0   'False
      End
   End
End
Attribute VB_Name = "frmQBFNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmQBFNew.frm
'' Description: Allows the user to start a new Quote Board Field from scratch,
''              or from an existing criteria
''
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date      Author      Description
'' 08/03/01  D Jarmuth   Created
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Enum eGDCols
    eGDCol_ID = 0
    eGDCol_Name
    eGDCol_Preview
    eGDCol_NumCols
End Enum

Private Type mPrivate
    bOK As Boolean
End Type
Private m As mPrivate

Private Function GDCol(ByVal Col As eGDCols) As Long
    GDCol = Col
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: If the user clicks on the cancel button, set the OK flag to
''              False and hide the form to let the ShowMe finish
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
    RaiseError "frmQBFNew.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cmdEdit_Click()
On Error GoTo ErrSection:

    Dim strID As String
    Dim strNewID As String
    Dim Criteria As New cCriteria
    Dim lIndex As Long

    strID = fgList.TextMatrix(fgList.Row, GDCol(eGDCol_ID))
    If strID <> "" Then
        strNewID = frmCriteria.ShowMe(AddSlash(App.Path) & "Custom\", strID, True, eCriteria_FilterCriteria)
        If strNewID <> "" Then
            If Criteria.FromFile(AddSlash(App.Path) & "Custom\", strNewID) Then
                With fgList
                    ' Add or change the existing criteria...
                    If strNewID <> strID Then
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, GDCol(eGDCol_ID)) = strNewID
                        .TextMatrix(.Rows - 1, GDCol(eGDCol_Name)) = Criteria.Name
                        .TextMatrix(.Rows - 1, GDCol(eGDCol_Preview)) = Criteria.EnglishText
                    Else
                        .TextMatrix(.Row, GDCol(eGDCol_ID)) = strNewID
                        .TextMatrix(.Row, GDCol(eGDCol_Name)) = Criteria.Name
                        .TextMatrix(.Row, GDCol(eGDCol_Preview)) = Criteria.EnglishText
                        rtbPreview.Text = .TextMatrix(.Row, GDCol(eGDCol_Preview))
                    End If
                    
                    ' Resort the list...
                    .Col = GDCol(eGDCol_Name)
                    .Sort = flexSortStringAscending
                    
                    ' Find and select the edited criteria...
                    For lIndex = .FixedRows To .Rows - 1
                        If .TextMatrix(lIndex, GDCol(eGDCol_ID)) = strNewID Then
                            .Row = lIndex
                            .RowSel = lIndex
                            .ShowCell lIndex, GDCol(eGDCol_Name)
                            Exit For
                        End If
                    Next lIndex
                End With
            End If
        End If
    End If

ErrExit:
    Set Criteria = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmQBFNew.cmdEdit.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cmdNew_Click()
On Error GoTo ErrSection:

    Dim strID As String
    Dim Criteria As New cCriteria
    Dim lIndex As Long
    
    ' TLB: pass "-" to create a new Inactive criteria (so won't ask to recalc)
    strID = frmCriteria.ShowMe(AddSlash(App.Path) & "Custom\", "-", True, eCriteria_FilterCriteria)
    If strID <> "" Then
        If Criteria.FromFile(AddSlash(App.Path) & "Custom\", strID) Then
            With fgList
                ' Add the new criteria to the list...
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, GDCol(eGDCol_ID)) = strID
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Name)) = Criteria.Name
                
                ' Resort the list...
                .Col = GDCol(eGDCol_Name)
                .Sort = flexSortStringAscending
                
                ' Find and show the new criteria...
                For lIndex = .FixedRows To .Rows - 1
                    If .TextMatrix(lIndex, GDCol(eGDCol_ID)) = strID Then
                        .Row = lIndex
                        .RowSel = lIndex
                        .ShowCell lIndex, GDCol(eGDCol_Name)
                        Exit For
                    End If
                Next lIndex
            End With
        End If
    End If

ErrExit:
    Set Criteria = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmQBFNew.cmdNew.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: If the user clicks on the OK button, bring up the Criteria
''              editor with the proper criteria or blank
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    m.bOK = True
    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQBFNew.cmdOK.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    rtbPreview.Text = fgList.TextMatrix(NewRow, GDCol(eGDCol_Preview))
    EnableButtons

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQBFNew.fgList.AfterRowColChange", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgList_DblClick()
On Error GoTo ErrSection:

    fgList.Row = fgList.MouseRow
    fgList.RowSel = fgList.Row
    m.bOK = True
    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQBFNew.DblClick", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgList_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    If KeyAscii = vbKeyReturn Then
        fgList.Row = fgList.MouseRow
        fgList.RowSel = fgList.Row
        m.bOK = True
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQBFNew.KeyPress", eGDRaiseError_Show
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
    RaiseError "frmQBFNew.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: When the form gets loaded, center it and load up the list box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    CenterTheForm Me
    Icon = Picture16(ToolbarIcon("kSelect"))
    
    g.Styler.StyleForm Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQBFNew.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user closes the form with the control menu, set the OK
''              flag to False and hide the form to allow the ShowMe to finish
'' Inputs:      Whether or not to cancel the unload, Unload Mode
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode = 0 Then
        m.bOK = False
        Cancel = True
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQBFNew.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Show the form modally, then return the results
'' Inputs:      None
'' Returns:     ID of the QBF to add, or blank if cancelled
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ByVal strFields As String) As cGdArray
On Error GoTo ErrSection:

    Dim astrReturn As New cGdArray      ' Array to return
    Dim lIndex As Long                  ' Index into a for loop
    Dim lRow As Long                    ' Row selected

    fgList.Redraw = flexRDNone
    InitGrid
    LoadGrid strFields
    fgList.Redraw = flexRDBuffered

    ShowForm Me, eForm_ActModal
    
    If m.bOK Then
        astrReturn.Size = 0
        For lIndex = 0 To fgList.SelectedRows - 1
            lRow = fgList.SelectedRow(lIndex)
            If fgList.TextMatrix(lRow, GDCol(eGDCol_ID)) = "" Then
                astrReturn.Add "*" & fgList.TextMatrix(lRow, GDCol(eGDCol_Name))
            Else
                astrReturn.Add fgList.TextMatrix(lRow, GDCol(eGDCol_ID))
            End If
        Next lIndex
        Set ShowMe = astrReturn
    Else
        Set ShowMe = Nothing
    End If
    
ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmQBFNew.ShowMe", eGDRaiseError_Raise
    
End Function

Private Sub Form_Resize()
On Error Resume Next

    If LimitFormSize(Me, fraButtons.Width * 3, fraButtons.Height + rtbPreview.Height + (fraButtons.Top * 3)) Then
        Exit Sub
    End If
    
    With fraButtons
        .Move ScaleWidth - .Width - fgList.Left, fgList.Top
    End With
    
    With rtbPreview
        .Move fgList.Left, ScaleHeight - .Height - fraButtons.Top, ScaleWidth - (fgList.Left * 2)
    End With
    
    With fgList
        .Move .Left, .Top, fraButtons.Left - (.Left * 2), ScaleHeight - rtbPreview.Height - (.Top * 3)
    End With

End Sub

Private Sub InitGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw
    
    With fgList
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = True
        .ScrollTrack = True
        .ExplorerBar = flexExSortShow
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        .ExtendLastCol = True
        .Editable = flexEDNone
        
        .BackColorBkg = g.Styler.GetColor(eGrid_Background) 'RH override vbApplicationWorkspacevbWindowBackground
        .GridLines = flexGridNone
        .GridLinesFixed = flexGridNone
        
        .Cols = GDCol(eGDCol_NumCols)
        .Rows = 0
        .FixedCols = 0
        .FixedRows = 0
        
        .ColHidden(GDCol(eGDCol_ID)) = True
        .ColHidden(GDCol(eGDCol_Preview)) = True
        
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQBFNew.InitGrid", eGDRaiseError_Raise
    
End Sub

Private Sub LoadGrid(ByVal strFields As String)
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim astrFields As New cGdArray      ' Fields passed in
    Dim lIndex As Long                  ' Index into a for loop
    Dim Criteria As New cCriteria       ' Criteria object
    Dim strName As String               ' Name of the field
    
    With fgList
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        ' First, add the standard quote board fields that are not being shown...
        astrFields.SplitFields strFields, "|"
        For lIndex = 0 To astrFields.Size - 1
            'If Parse(astrFields(lIndex), ";", 3) = "" And Len(astrFields(lIndex)) > 0 Then
            If Len(astrFields(lIndex)) > 0 Then
                If Val(Parse(astrFields(lIndex), ";", 1)) <> 0 Then
                    strName = Parse(astrFields(lIndex), ";", 2)
                    If UCase(strName) <> "SYMBOLID" And UCase(strName) <> "SECTYPE" Then
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, GDCol(eGDCol_ID)) = "" 'Parse(astrFields(lIndex), ";", 3)
                        .TextMatrix(.Rows - 1, GDCol(eGDCol_Name)) = strName
                        .TextMatrix(.Rows - 1, GDCol(eGDCol_Preview)) = ""
                    End If
                End If
            End If
        Next lIndex
        
        ' Second, add any criteria that are not being shown...
        For Each Criteria In g.SymbolPool.Criterias
            If Criteria.Custom Then
                If InStr(UCase(strFields), UCase(Criteria.ID)) = 0 Then
                    If HasModule(Criteria.Required) Then
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, GDCol(eGDCol_ID)) = Criteria.ID
                        .TextMatrix(.Rows - 1, GDCol(eGDCol_Name)) = Criteria.Name
                        .TextMatrix(.Rows - 1, GDCol(eGDCol_Preview)) = Criteria.EnglishText
                    End If
                End If
            End If
        Next Criteria
        
        If .Rows > .FixedRows Then
            .Row = .FixedRows
            .RowSel = .Row
            
            .Col = GDCol(eGDCol_Name)
            .Sort = flexSortStringAscending
            
            .ShowCell .Row, GDCol(eGDCol_Name)
            rtbPreview.Text = .TextMatrix(.Row, GDCol(eGDCol_Preview))
            MoveFocus fgList
            EnableButtons
        End If
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQBFNew.LoadGrid", eGDRaiseError_Raise
    
End Sub

Private Sub EnableButtons()
On Error GoTo ErrSection:

    cmdEdit.Enabled = (fgList.TextMatrix(fgList.Row, GDCol(eGDCol_ID)) <> "")

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmQBFNew.EnableButtons", eGDRaiseError_Raise
    
End Sub

