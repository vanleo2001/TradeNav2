VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmExport 
   Caption         =   "Data Exporting"
   ClientHeight    =   3795
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   9285
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   9285
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   2715
      Left            =   6360
      TabIndex        =   1
      Top             =   240
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
      Caption         =   "frmExport.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmExport.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmExport.frx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Default         =   -1  'True
         Height          =   375
         Left            =   0
         TabIndex        =   7
         Top             =   2340
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
         Caption         =   "frmExport.frx":0068
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmExport.frx":0096
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmExport.frx":00B6
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Height          =   375
         Left            =   0
         TabIndex        =   6
         Top             =   1920
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
         Caption         =   "frmExport.frx":00D2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmExport.frx":00F8
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmExport.frx":0118
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdExport 
         Height          =   375
         Left            =   0
         TabIndex        =   5
         Top             =   1260
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
         Caption         =   "frmExport.frx":0134
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmExport.frx":016A
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmExport.frx":018A
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdRemove 
         Height          =   375
         Left            =   0
         TabIndex        =   4
         Top             =   840
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
         Caption         =   "frmExport.frx":01A6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmExport.frx":01D4
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmExport.frx":01F4
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdEdit 
         Height          =   375
         Left            =   0
         TabIndex        =   3
         Top             =   420
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
         Caption         =   "frmExport.frx":0210
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmExport.frx":023A
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmExport.frx":025A
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdAdd 
         Height          =   375
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
         Caption         =   "frmExport.frx":0276
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmExport.frx":029E
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmExport.frx":02BE
         RightToLeft     =   0   'False
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgExport 
      Height          =   2475
      Left            =   240
      TabIndex        =   0
      Top             =   900
      Width           =   5775
      _cx             =   10186
      _cy             =   4366
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
   Begin HexUniControls.ctlUniLabelXP lblDesc 
      Height          =   435
      Left            =   240
      Top             =   300
      Width           =   5775
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
      Caption         =   "frmExport.frx":02DA
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   2
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmExport.frx":0428
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmExport.frx":0448
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add Export Group"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit Export Group"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "&Delete Export Group"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "E&xport"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangeFont 
         Caption         =   "&Change Font"
      End
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmExport.frm
'' Description: Form to allow the user to select the symbol groups that they
''              would like to export to another format
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 06/21/2001   D Jarmuth   Created
'' 10/16/2014   DAJ         Replaced File System Object calls
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Enum eGDCols
    eGDCol_SymbolGroupID = 0
    eGDCol_AutoExport = 1
    eGDCol_SymbolGroup = 2
    eGDCol_Format = 3
    eGDCol_Path = 4
    eGDCol_Period = 5
    eGDCol_FromDate = 6
    eGDCol_ToDate = 7
End Enum
Private Const kNumCols = 8

Private Function GDCol(ByVal lColumn As eGDCols) As Long
    GDCol = lColumn
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdAdd_Click
'' Description: When the user clicks on the Add button, bring up the symbol
''              group export form to create a new symbol group export
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAdd_Click()
On Error GoTo ErrSection:

    Add

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExport.cmdAdd.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: If the user clicks on the cancel button, unload the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    Unload Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExport.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdEdit_Click
'' Description: When the user clicks on the edit button, bring up the symbol
''              group export form with the information from the currently
''              selected row in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdEdit_Click()
On Error GoTo ErrSection:

    Edit
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExport.cmdEdit.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdExport_Click
'' Description: If the user clicks on the Export button, export the data
''              according to the format/path in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdExport_Click()
On Error GoTo ErrSection:

    Export

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExport.cmdExport.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: If the user clicks on the OK button, save and exit the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    Save
    Unload Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExport.cmdOK.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdRemove_Click
'' Description: If the user clicks on the Remove button, remove the currently
''              selected row in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdRemove_Click()
On Error GoTo ErrSection:

    'aardvark 6493
    If InfBox("Delete the selected data export entry?", "?", "-Yes|+No", "Data Exporting Delete Confirmation") = "Y" Then
        Remove
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExport.cmdRemove.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgExport_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    With fgExport
        If KeyCode = vbKeyDelete Then
            Remove
        ElseIf KeyCode = vbKeyInsert Then
            Add
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExport.fgExport.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgExport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim lMouseRow As Long
    Dim lMouseCol As Long
    
    With fgExport
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If Button = vbRightButton Then
            If lMouseRow >= .FixedRows And lMouseRow < .Rows Then
                .RowSel = lMouseRow
                If Not .IsSelected(lMouseRow) Then .Row = lMouseRow
            End If
            
            mnuEdit.Enabled = (lMouseRow >= .FixedRows And lMouseRow < .Rows)
            mnuRemove.Enabled = (lMouseRow >= .FixedRows And lMouseRow < .Rows)
            mnuExport.Enabled = (lMouseRow >= .FixedRows And lMouseRow < .Rows)
            
            PopupMenu mnuPopUp
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExport.fgExport.MouseDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub fgExport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    GridTooltip fgExport
    
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
    RaiseError "frmExport.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: When the form is loaded, center it and set up the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strFont As String

    ' Center the form
    Me.Icon = Picture16(ToolbarIcon("ID_ExportData"), , True)
    CenterTheForm Me
    
    g.Styler.StyleForm Me

    ' Set up the grid
    With fgExport
        .Redraw = flexRDNone
        .ExtendLastCol = True
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExSortShow
        .SelectionMode = flexSelectionListBox
        .AllowUserResizing = flexResizeColumns
        .SheetBorder = RGB(128, 128, 128)
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        
        .FixedRows = 1
        .Rows = 1
        .FixedCols = 0
        .Cols = kNumCols
        
        .Cell(flexcpText, 0, GDCol(eGDCol_SymbolGroupID)) = "Symbol Group ID"
        .Cell(flexcpText, 0, GDCol(eGDCol_AutoExport)) = "Auto Export"
        .Cell(flexcpText, 0, GDCol(eGDCol_SymbolGroup)) = "Symbol Group"
        .Cell(flexcpText, 0, GDCol(eGDCol_Format)) = "Format"
        .Cell(flexcpText, 0, GDCol(eGDCol_Path)) = "Path"
        .Cell(flexcpText, 0, GDCol(eGDCol_Period)) = "Period"
        .Cell(flexcpText, 0, GDCol(eGDCol_FromDate)) = "From Date"
        .Cell(flexcpText, 0, GDCol(eGDCol_ToDate)) = "To Date"
        
        .ColDataType(GDCol(eGDCol_AutoExport)) = flexDTBoolean
        
        .ColHidden(GDCol(eGDCol_SymbolGroupID)) = True
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With
    
    Load
    EnableButtons
    
    If fgExport.Rows = fgExport.FixedRows Then cmdAdd_Click
    
    mnuPopUp.Visible = False
    strFont = GetIniFileProperty("Export", "", "Fonts", g.strIniFile)
    If strFont <> "" Then FontFromString fgExport.Font, strFont
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExport.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: As the user resizes the form, resize and move the controls on
''              the form as necessary
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    If LimitFormSize(Me, fraButtons.Width * 7, fraButtons.Height + fraButtons.Top * 2) Then Exit Sub
    
    fraButtons.Left = Me.ScaleWidth - (fgExport.Left + fraButtons.Width)
    
    With fgExport
        .Move .Left, lblDesc.Top + lblDesc.Height, fraButtons.Left - (fgExport.Left * 2), _
                Me.ScaleHeight - lblDesc.Height - (lblDesc.Top * 2)
    End With

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgExport_AfterEdit
'' Description: If the user changes the value of the auto export, save it to
''              the class that is stored in the rowdata
'' Inputs:      Row and Column of the change
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgExport_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    If Col = GDCol(eGDCol_AutoExport) Then
        If fgExport.Cell(flexcpChecked, Row, GDCol(eGDCol_AutoExport)) = flexChecked Then
            fgExport.RowData(Row).AutoExport = True
        Else
            fgExport.RowData(Row).AutoExport = False
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExport.fgExport.AfterEdit", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgExport_BeforeEdit
'' Description: If the user tries to edit the grid, only let them edit the
''              first column
'' Inputs:      Row, Column, Whether or not to cancel the edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgExport_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    If Col > GDCol(eGDCol_AutoExport) Then
        Cancel = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExport.fgExport.BeforeEdit", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableButtons
'' Description: Only allow the Edit, Remove, and Export buttons to be enabled
''              if there are rows in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableButtons()
On Error GoTo ErrSection:

    Dim bEnable As Boolean

    bEnable = fgExport.Rows > fgExport.FixedRows
    
    cmdEdit.Enabled = bEnable
    cmdRemove.Enabled = bEnable
    cmdExport.Enabled = bEnable

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExport.EnableButtons", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgExport_DblClick
'' Description: If the user double clicks in the grid, simulate an edit button
''              click
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgExport_DblClick()
On Error GoTo ErrSection:

    Edit

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExport.fgExport.DblClick", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Save the information in the grid to a file for later retrieval
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Save()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim astrLines As New cGdArray       ' Lines from the grid
    Dim strFileName As String           ' Name of the export file
    
    With fgExport
        For lIndex = .FixedRows To .Rows - .FixedRows
            astrLines.Add .RowData(lIndex).ToString
        Next lIndex
    End With
    
    strFileName = AddSlash(App.Path) & "Custom\Export.TXT"
    If astrLines.Size = 0 Then
        If FileExist(strFileName) Then
            mGenesis.KillFile strFileName, True
        End If
    ElseIf Not astrLines.ToFile(strFileName) Then
        Err.Raise vbObjectError + 1000, , "Could not create Export.TXT"
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExport.Save", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Load
'' Description: Load the saved information into the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Load()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into the grid
    Dim ExportGroup As cExportGroup     ' Export group object
    Dim aStrings As New cGdArray
    Dim i&
    
    If FileExist(App.Path & "\Custom\Export.TXT") Then
        aStrings.FromFile App.Path & "\Custom\Export.TXT"
        lIndex = fgExport.FixedRows
        With fgExport
            .Redraw = flexRDNone
            For i = 0 To aStrings.Size - 1
                Set ExportGroup = New cExportGroup
                ExportGroup.FromString aStrings(i)
                If g.SymbolPool.FieldNumForID(ExportGroup.SymbolGroupID) <> -1 Then
                    .Rows = .Rows + 1
                    '.Cell(flexcpChecked, lIndex, GDCol(eGDCol_AutoExport)) = ExportGroup.AutoExport
                    If ExportGroup.AutoExport = True Then
                        .Cell(flexcpChecked, lIndex, GDCol(eGDCol_AutoExport)) = flexChecked
                    Else
                        .Cell(flexcpChecked, lIndex, GDCol(eGDCol_AutoExport)) = flexUnchecked
                    End If
                    .Cell(flexcpText, lIndex, GDCol(eGDCol_SymbolGroupID)) = ExportGroup.SymbolGroupID
                    .Cell(flexcpText, lIndex, GDCol(eGDCol_SymbolGroup)) = g.SymbolPool.PoolObject(ExportGroup.SymbolGroupID).Name
                    .Cell(flexcpText, lIndex, GDCol(eGDCol_Format)) = ExportGroup.Format
                    .Cell(flexcpText, lIndex, GDCol(eGDCol_Path)) = ExportGroup.Path
                    .Cell(flexcpText, lIndex, GDCol(eGDCol_Period)) = ExportGroup.Period
                    .Cell(flexcpText, lIndex, GDCol(eGDCol_FromDate)) = DisplayDate(ExportGroup.StartDate)
                    .Cell(flexcpText, lIndex, GDCol(eGDCol_ToDate)) = DisplayDate(ExportGroup.EndDate)
                    .RowData(lIndex) = ExportGroup
                    lIndex = lIndex + 1&
                End If
            Next i
            .AutoSize 0, .Cols - 1, False, 75
            If .Rows > .FixedRows Then .Select .FixedRows, 0, .FixedRows, .Cols - 1
            .Redraw = flexRDBuffered
        End With
    End If
    
ErrExit:
    Set aStrings = Nothing
    Exit Sub
    
ErrSection:
    Set aStrings = Nothing
    RaiseError "frmExport.Load", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DisplayDate
'' Description: Format a string to display in the grid for a specific date
'' Inputs:      Date to display
'' Returns:     Formatted Date
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DisplayDate(ByVal dDate As Double) As String
On Error GoTo ErrSection:

    If dDate = 0 Then
        DisplayDate = DateFormat(Date)
    Else
        DisplayDate = DateFormat(dDate)
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmExport.DisplayDate", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Edit
'' Description: Allow the user to edit a export group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Edit()
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Row in the grid to edit
    
    With fgExport
        lRow = .RowSel
        If lRow >= .FixedRows Then
            If frmExportGroup.ShowMe(.RowData(lRow)) = True Then
                .Cell(flexcpText, .Rows - 1, GDCol(eGDCol_SymbolGroupID)) = .RowData(lRow).SymbolGroupID
                .Cell(flexcpText, lRow, GDCol(eGDCol_SymbolGroup)) = .RowData(lRow).SymbolGroup
                .Cell(flexcpText, lRow, GDCol(eGDCol_Format)) = .RowData(lRow).Format
                .Cell(flexcpText, lRow, GDCol(eGDCol_Path)) = .RowData(lRow).Path
                .Cell(flexcpText, lRow, GDCol(eGDCol_Period)) = .RowData(lRow).Period
                .Cell(flexcpText, lRow, GDCol(eGDCol_FromDate)) = DisplayDate(.RowData(lRow).StartDate)
                .Cell(flexcpText, lRow, GDCol(eGDCol_ToDate)) = DisplayDate(.RowData(lRow).EndDate)
                .AutoSize 0, .Cols - 1, False, 75
            End If
        End If
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExport.Edit", eGDRaiseError_Raise
    
End Sub

Private Sub Add()
On Error GoTo ErrSection:

    Dim ExportGroup As New cExportGroup ' Export group object
    
    If frmExportGroup.ShowMe(ExportGroup) = True Then
        With fgExport
            .Rows = .Rows + 1
            .Cell(flexcpText, .Rows - 1, GDCol(eGDCol_SymbolGroupID)) = ExportGroup.SymbolGroupID
            .Cell(flexcpText, .Rows - 1, GDCol(eGDCol_SymbolGroup)) = ExportGroup.SymbolGroup
            .Cell(flexcpText, .Rows - 1, GDCol(eGDCol_Format)) = ExportGroup.Format
            .Cell(flexcpText, .Rows - 1, GDCol(eGDCol_Path)) = ExportGroup.Path
            .Cell(flexcpText, .Rows - 1, GDCol(eGDCol_Period)) = ExportGroup.Period
            .Cell(flexcpText, .Rows - 1, GDCol(eGDCol_FromDate)) = DisplayDate(ExportGroup.StartDate)
            .Cell(flexcpText, .Rows - 1, GDCol(eGDCol_ToDate)) = DisplayDate(ExportGroup.EndDate)
            .RowData(.Rows - 1) = ExportGroup
            .Select .Rows - 1, 0
            .AutoSize 0, .Cols - 1, False, 75
        End With
        
        EnableButtons
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExport.Add", eGDRaiseError_Raise
    
End Sub

Private Sub Remove()
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Row to delete out of the grid

    With fgExport
        lRow = .RowSel
        If lRow > 0 Then
            .RemoveItem lRow
            
            If lRow - 1 > 0 Then
                .Select lRow - 1, 0
            ElseIf lRow > 0 And .Rows > .FixedRows Then
                .Select lRow, 0
            End If
            
            EnableButtons
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExport.Remove", eGDRaiseError_Raise
    
End Sub

Private Sub Export()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index for a for loop
    Dim lRow As Long                    ' Index for a for loop
    Dim strReturn As String             ' Return from an ask box
    Dim lFieldNum As Long               ' Field number for the symbol group
    Dim Bars As New cGdBars             ' Object that holds the data
    Dim lSymbolID As Long               ' Symbol ID to get the data for
    Dim strFormat As String             ' Format to export to
    Dim bOnlyIfSelected As Boolean      ' Export group only if selected

    Save

    If fgExport.Rows > fgExport.FixedRows + 1 Then
        strReturn = AskBox("h=Export ; i=? ; b=+Selected|All|-Cancel ; Do you want to export all|" & _
                            "groups at this time or only the|selected ones?")
        If UCase(strReturn) = "C" Then Exit Sub
        If UCase(strReturn) = "A" Then bOnlyIfSelected = False Else bOnlyIfSelected = True
    End If
    
    Me.Hide
    DoEvents

    With fgExport
        For lRow = .FixedRows To .Rows - .FixedRows
            If .IsSelected(lRow) = True Or Not bOnlyIfSelected Then
                .RowData(lRow).Export
            End If
        Next lRow
    End With
    If frmStatus.Visible Then frmStatus.AddDetail "Finished"
    
    ShowForm Me, True, , , ALT_GRID_ROW_COLOR

ErrExit:
    Set Bars = Nothing
    Exit Sub
    
ErrSection:
    Set Bars = Nothing
    RaiseError "frmExport.Export", eGDRaiseError_Raise
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    SetIniFileProperty "Export", FontToString(fgExport.Font), "Fonts", g.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExport.Form.Unload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub mnuAdd_Click()
On Error GoTo ErrSection:

    Add

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExport.mnuAdd.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub mnuEdit_Click()
On Error GoTo ErrSection:

    Edit

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExport.mnuEdit.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub mnuExport_Click()
On Error GoTo ErrSection:

    Export

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExport.mnuExport.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub mnuRemove_Click()
On Error GoTo ErrSection:

    Remove

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmExport.mnuRemove.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

