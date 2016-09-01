VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmTradeRules 
   Caption         =   "Form1"
   ClientHeight    =   4065
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   7485
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   2595
      Left            =   6120
      TabIndex        =   4
      Top             =   120
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
      Caption         =   "frmTradeRules.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTradeRules.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTradeRules.frx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdClose 
         Cancel          =   -1  'True
         Height          =   495
         Left            =   0
         TabIndex        =   5
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
         Caption         =   "frmTradeRules.frx":0068
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTradeRules.frx":0094
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTradeRules.frx":00B4
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdDelete 
         Height          =   495
         Left            =   0
         TabIndex        =   0
         Top             =   2100
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
         Caption         =   "frmTradeRules.frx":00D0
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTradeRules.frx":0108
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTradeRules.frx":0128
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdEdit 
         Height          =   495
         Left            =   0
         TabIndex        =   2
         Top             =   1560
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
         Caption         =   "frmTradeRules.frx":0144
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTradeRules.frx":0178
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTradeRules.frx":0198
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdNew 
         Height          =   495
         Left            =   0
         TabIndex        =   6
         Top             =   1020
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
         Caption         =   "frmTradeRules.frx":01B4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTradeRules.frx":01E6
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTradeRules.frx":0206
         RightToLeft     =   0   'False
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgExitRules 
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   2340
      Width           =   5775
      _cx             =   10186
      _cy             =   2778
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
   Begin VSFlex7LCtl.VSFlexGrid fgEntryRules 
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   5775
      _cx             =   10186
      _cy             =   2778
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
   Begin HexUniControls.ctlUniLabelXP lblExitRules 
      Height          =   195
      Left            =   120
      Top             =   2100
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
      Caption         =   "frmTradeRules.frx":0222
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmTradeRules.frx":0258
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTradeRules.frx":0278
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblEntryRules 
      Height          =   195
      Left            =   120
      Top             =   120
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
      Caption         =   "frmTradeRules.frx":0294
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmTradeRules.frx":02CC
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTradeRules.frx":02EC
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Begin VB.Menu mnuNewRule 
         Caption         =   "New Rule"
      End
      Begin VB.Menu mnuEditRule 
         Caption         =   "Edit Rule"
      End
      Begin VB.Menu mnuDeleteRule 
         Caption         =   "Delete Rule"
      End
   End
End
Attribute VB_Name = "frmTradeRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmTradeRules.frm
'' Description: Allow the user to manage their custom trade rules
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Enum eGDCols
    eGDCol_TradeRuleID = 0
    eGDCol_Abbreviation
    eGDCol_Name
    eGDCol_Description
    eGDCol_NumCols
End Enum

Private Type mPrivate
    nLastFocus As eGDTradeRuleTypes     ' Last grid to have the focus
    bChanged As Boolean                 ' Did something change?
End Type
Private m As mPrivate

Private Function GDCol(ByVal nCol As eGDCols) As Long
    GDCol = nCol
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Setup and show the form
'' Inputs:      None
'' Returns:     True if something changed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe() As Boolean
On Error GoTo ErrSection:

    InitGrid fgEntryRules
    LoadGrid fgEntryRules
    
    InitGrid fgExitRules
    LoadGrid fgExitRules
    
    EnableControls

    m.bChanged = False
    MoveFocus fgEntryRules
    ShowForm Me, eForm_ActModal, frmMain, , ALT_GRID_ROW_COLOR
    
    ShowMe = m.bChanged

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmTradeRules.ShowMe"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdClose_Click
'' Description: Hide the form and allow the ShowMe routine to unload it
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdClose_Click()
On Error GoTo ErrSection:

    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeRules.cmdClose_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdDelete_Click
'' Description: Allow the user to delete an existing rule
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdDelete_Click()
On Error GoTo ErrSection:

    RemoveRule m.nLastFocus

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeRules.cmdDelete_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdEdit_Click
'' Description: Allow the user to edit an existing rule
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdEdit_Click()
On Error GoTo ErrSection:

    EditRule m.nLastFocus

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeRules.cmdEdit_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdNew_Click
'' Description: Allow the user to create a new rule
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdNew_Click()
On Error GoTo ErrSection:

    If InfBox("Would you like to create|an Entry Rule or an Exit Rule?", "?", "+Entry|-E&xit", "New Trade Rule") = "E" Then
        NewRule eGDTradeRuleType_Entry
    Else
        NewRule eGDTradeRuleType_Exit
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeRules.cmdNew_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgEntryRules_AfterRowColChange
'' Description: After a row change, enable or disable controls
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgEntryRules_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    If OldRow <> NewRow Then
        EnableControls
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeRules.fgEntryRules_AfterRowColChange"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgEntryRules_BeforeMouseDown
'' Description: If user right clicks, show the popup menu
'' Inputs:      Button Pressed, Shift/Ctrl/Alt Status, Location, Cancel?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgEntryRules_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    MoveFocus fgEntryRules

    ' Set the current row...
    fgEntryRules.Row = fgEntryRules.MouseRow

    If Button = vbRightButton Then
        mnuPopUp.Tag = "ENTRY"
        EnableControls
        PopupMenu mnuPopUp
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeRules.fgEntryRules_BeforeMouseDown"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgEntryRules_DblClick
'' Description: When user double clicks on an item, edit it
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgEntryRules_DblClick()
On Error GoTo ErrSection:

    ' Set the current row...
    fgEntryRules.Row = fgEntryRules.MouseRow

    If ValidRowSelected(fgEntryRules) Then
        EditRule eGDTradeRuleType_Entry
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeRules.fgEntryRules_DblClick"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgEntryRules_GotFocus
'' Description: Keep track of which grid was last to get the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgEntryRules_GotFocus()
On Error GoTo ErrSection:

    m.nLastFocus = eGDTradeRuleType_Entry
    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeRules.fgEntryRules_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgEntryRules_KeyDown
'' Description: Allow the user to add or delete an entry rule with the Insert and
''              Delete keys
'' Inputs:      Code of the Key Pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgEntryRules_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyDelete Then
        If ValidRowSelected(fgEntryRules) Then
            RemoveRule eGDTradeRuleType_Entry
        End If
    ElseIf KeyCode = vbKeyInsert Then
        NewRule eGDTradeRuleType_Entry
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeRules.fgEntryRules_KeyDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgEntryRules_KeyPress
'' Description: If the user hits Enter on an entry rule, allow them to edit it
'' Inputs:      Key Pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgEntryRules_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    If KeyAscii = vbKeyReturn Then
        If ValidRowSelected(fgEntryRules) Then
            EditRule eGDTradeRuleType_Entry
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeRules.fgEntryRules_KeyPress"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgExitRules_AfterRowColChange
'' Description: After a row change, enable or disable controls
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgExitRules_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    If OldRow <> NewRow Then
        EnableControls
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeRules.fgExitRules_AfterRowColChange"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgExitRules_BeforeMouseDown
'' Description: If user right clicks, show the popup menu
'' Inputs:      Button Pressed, Shift/Ctrl/Alt Status, Location, Cancel?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgExitRules_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    MoveFocus fgExitRules

    ' Set the current row...
    fgExitRules.Row = fgExitRules.MouseRow

    If Button = vbRightButton Then
        mnuPopUp.Tag = "EXIT"
        EnableControls
        PopupMenu mnuPopUp
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeRules.fgExitRules_BeforeMouseDown"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgExitRules_DblClick
'' Description: When user double clicks on an item, edit it
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgExitRules_DblClick()
On Error GoTo ErrSection:

    ' Set the current row...
    fgExitRules.Row = fgExitRules.MouseRow

    If ValidRowSelected(fgExitRules) Then
        EditRule eGDTradeRuleType_Exit
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeRules.fgExitRules_DblClick"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgExitRules_GotFocus
'' Description: Keep track of which grid was last to get the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgExitRules_GotFocus()
On Error GoTo ErrSection:

    m.nLastFocus = eGDTradeRuleType_Exit
    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeRules.fgExitRules_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgExitRules_KeyDown
'' Description: Allow the user to add or delete an exit rule with the Insert and
''              Delete keys
'' Inputs:      Code of the Key Pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgExitRules_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyDelete Then
        If ValidRowSelected(fgExitRules) Then
            RemoveRule eGDTradeRuleType_Exit
        End If
    ElseIf KeyCode = vbKeyInsert Then
        NewRule eGDTradeRuleType_Exit
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeRules.fgExitRules_KeyDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgExitRules_KeyPress
'' Description: If the user hits Enter on an exit rule, allow them to edit it
'' Inputs:      Key Pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgExitRules_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    If KeyAscii = vbKeyReturn Then
        If ValidRowSelected(fgExitRules) Then
            EditRule eGDTradeRuleType_Exit
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeRules.fgExitRules_KeyPress"

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

    Dim strPlacement As String          ' Placement of the form
    
    strPlacement = GetIniFileProperty("frmTradeRules", "", "Placement", g.strIniFile)
    If Len(strPlacement) > 0 Then
        SetFormPlacement Me, strPlacement, "LTHW"
    Else
        CenterTheForm Me
    End If
    
    g.Styler.StyleForm Me
    
    mnuPopUp.Visible = False
    
    Caption = "Custom Trade Rules"
    Icon = Picture16("kBlank")

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeRules.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user clicks on the 'X', allow the ShowMe to unload form
'' Inputs:      Cancel the Unload?, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        Cancel = True
        Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeRules.Form_QueryUnload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Clean up and save settings when the form is unloaded
'' Inputs:      Cancel the Unload?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    SetIniFileProperty "frmTradeRules", GetFormPlacement(Me), "Placement", g.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeRules.Form_Unload"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuDeleteRule_Click
'' Description: Allow the user to delete the selected rule
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuDeleteRule_Click()
On Error GoTo ErrSection:

    If UCase(mnuPopUp.Tag) = "ENTRY" Then
        RemoveRule eGDTradeRuleType_Entry
    ElseIf UCase(mnuPopUp.Tag) = "EXIT" Then
        RemoveRule eGDTradeRuleType_Exit
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeRules.mnuDeleteRule_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuEditRule_Click
'' Description: Allow the user to edit the selected rule
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuEditRule_Click()
On Error GoTo ErrSection:

    If UCase(mnuPopUp.Tag) = "ENTRY" Then
        EditRule eGDTradeRuleType_Entry
    ElseIf UCase(mnuPopUp.Tag) = "EXIT" Then
        EditRule eGDTradeRuleType_Exit
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeRules.mnuEditRule_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuNewRule_Click
'' Description: Allow the user to create a new rule
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuNewRule_Click()
On Error GoTo ErrSection:

    If UCase(mnuPopUp.Tag) = "ENTRY" Then
        NewRule eGDTradeRuleType_Entry
    ElseIf UCase(mnuPopUp.Tag) = "EXIT" Then
        NewRule eGDTradeRuleType_Exit
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeRules.mnuNewRule_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitGrid
'' Description: Initialize the grid
'' Inputs:      Grid
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitGrid(Grid As VSFlexGrid)
On Error GoTo ErrSection:

    With Grid
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = True
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDNone
        .ExplorerBar = flexExNone
        .ExtendLastCol = True
        .HighLight = flexHighlightWithFocus
        .ScrollTrack = True
        .SelectionMode = flexSelectionByRow
        .SheetBorder = RGB(128, 128, 128)
        
        .Rows = 1
        .FixedRows = 1
        .Cols = GDCol(eGDCol_NumCols)
        .FixedCols = 0
        
        .TextMatrix(0, GDCol(eGDCol_TradeRuleID)) = "ID"
        .TextMatrix(0, GDCol(eGDCol_Abbreviation)) = "Abbrev"
        .TextMatrix(0, GDCol(eGDCol_Name)) = "Name"
        .TextMatrix(0, GDCol(eGDCol_Description)) = "Description"
        
        .ColHidden(GDCol(eGDCol_TradeRuleID)) = True
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeRules.InitGrid"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGrid
'' Description: Load the grid
'' Inputs:      Grid
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGrid(Grid As VSFlexGrid)
On Error GoTo ErrSection:

    Dim strFile As String               ' Name of the file to get data from
    Dim astrFile As cGdArray            ' File read into an array
    Dim lIndex As Long                  ' Index into a for loop
    Dim nRuleType As eGDTradeRuleTypes  ' Rule type for the trade rule
    Dim TradeRule As cTradeRule         ' Trade rule object
    
    If Grid.Name = "fgEntryRules" Then
        strFile = AddSlash(App.Path) & "Custom\ErFilter.TXT"
        nRuleType = eGDTradeRuleType_Entry
    Else
        strFile = AddSlash(App.Path) & "Custom\XrFilter.TXT"
        nRuleType = eGDTradeRuleType_Exit
    End If
    
    Set astrFile = New cGdArray
    
    With Grid
        .Redraw = flexRDNone
        
        If astrFile.FromFile(strFile) Then
            For lIndex = 0 To astrFile.Size - 1
                Set TradeRule = New cTradeRule
                TradeRule.FromString astrFile(lIndex), False, nRuleType
                TradeRuleToGrid TradeRule, Grid, , False
            Next lIndex
        End If
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTradeRules.LoadGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TradeRuleToGrid
'' Description: Display the trade rule in the grid
'' Inputs:      Trade Rule, Grid, Row, Size Grid?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TradeRuleToGrid(TradeRule As cTradeRule, Grid As VSFlexGrid, Optional ByVal lRow As Long = -1, Optional ByVal bSizeGrid As Boolean = True)
On Error GoTo ErrSection:

    With Grid
        If lRow = -1 Then
            .Rows = .Rows + 1
            lRow = .Rows - 1
        End If
    
        .RowData(lRow) = TradeRule
        .TextMatrix(lRow, GDCol(eGDCol_TradeRuleID)) = Str(TradeRule.ID)
        .TextMatrix(lRow, GDCol(eGDCol_Abbreviation)) = TradeRule.Abbreviation
        .TextMatrix(lRow, GDCol(eGDCol_Name)) = TradeRule.Name
        .TextMatrix(lRow, GDCol(eGDCol_Description)) = TradeRule.Description
        
        If bSizeGrid Then
            .AutoSize 0, .Cols - 1, False, 75
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeRules.TradeRuleToGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveGrid
'' Description: Save the grid
'' Inputs:      Grid
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveGrid(Grid As VSFlexGrid)
On Error GoTo ErrSection:

    Dim strFile As String               ' Name of the file to get data from
    Dim astrFile As cGdArray            ' File read into an array
    Dim lIndex As Long                  ' Index into a for loop
    Dim TradeRule As cTradeRule         ' Trade rule object
    
    If Grid.Name = "fgEntryRules" Then
        strFile = AddSlash(App.Path) & "Custom\ErFilter.TXT"
    Else
        strFile = AddSlash(App.Path) & "Custom\XrFilter.TXT"
    End If
    
    Set astrFile = New cGdArray
    astrFile.Create eGDARRAY_Strings
    
    With Grid
        For lIndex = .FixedRows To .Rows - 1
            If TypeOf .RowData(lIndex) Is cTradeRule Then
                Set TradeRule = .RowData(lIndex)
                astrFile.Add TradeRule.ToString
            End If
        Next lIndex
        
        astrFile.ToFile strFile
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeRules.SaveGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NewRule
'' Description: Allow the user to create a new trade rule
'' Inputs:      Rule Type
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub NewRule(ByVal nRuleType As eGDTradeRuleTypes)
On Error GoTo ErrSection:

    Dim TradeRule As cTradeRule         ' Trade rule object
    
    Set TradeRule = New cTradeRule
    TradeRule.Provided = False
    TradeRule.RuleType = nRuleType
    
    If frmTradeRule.ShowMe(TradeRule) Then
        If nRuleType = eGDTradeRuleType_Entry Then
            TradeRuleToGrid TradeRule, fgEntryRules
            SaveGrid fgEntryRules
        Else
            TradeRuleToGrid TradeRule, fgExitRules
            SaveGrid fgExitRules
        End If
        
        m.bChanged = True
        EnableControls
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeRules.NewRule"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EditRule
'' Description: Allow the user to edit an existing trade rule
'' Inputs:      Rule Type
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EditRule(ByVal nRuleType As eGDTradeRuleTypes)
On Error GoTo ErrSection:

    Dim Grid As VSFlexGrid              ' Grid that we are working with
    Dim TradeRule As cTradeRule         ' Trade rule object to edit

    If nRuleType = eGDTradeRuleType_Entry Then
        Set Grid = fgEntryRules
    Else
        Set Grid = fgExitRules
    End If
    
    If ValidRowSelected(Grid) Then
        Set TradeRule = Grid.RowData(Grid.Row)
        If frmTradeRule.ShowMe(TradeRule) Then
            TradeRuleToGrid TradeRule, Grid, Grid.Row
            
            SaveGrid Grid
            
            m.bChanged = True
            EnableControls
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeRules.EditRule"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveRule
'' Description: Allow the user to remove an existing trade rule
'' Inputs:      Rule Type
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RemoveRule(ByVal nType As eGDTradeRuleTypes)
On Error GoTo ErrSection:

    Dim strRule As String               ' Rule to delete
    Dim strType As String               ' Entry or Exit?
    Dim lID As Long                     ' ID of the rule to delete
    Dim rs As Recordset                 ' Recordset into the database
    Dim Grid As VSFlexGrid              ' Grid for the given type
    Dim Fill As cPtFill                 ' Fill object
    Dim TradeLine As cTradeLine         ' Trade line object

    If nType = eGDTradeRuleType_Entry Then
        strRule = fgEntryRules.TextMatrix(fgEntryRules.Row, GDCol(eGDCol_Name))
        strType = "entry"
        Set Grid = fgEntryRules
    Else
        strRule = fgExitRules.TextMatrix(fgExitRules.Row, GDCol(eGDCol_Name))
        strType = "exit"
        Set Grid = fgExitRules
    End If

    If ValidRowSelected(Grid) Then
        If InfBox("Are you sure you want to|delete the " & strType & " rule||'" & strRule & "'?", "?", "+Yes|-No", "Delete Confirmation") = "Y" Then
            If nType = eGDTradeRuleType_Entry Then
                lID = CLng(Val(fgEntryRules.TextMatrix(fgEntryRules.Row, GDCol(eGDCol_TradeRuleID))))
                
                ' Attempt to retrieve tradelines from the appropriate broker info object that had
                ' the entry rule set to the one that was just deleted, clear the entry rule, then
                ' resend the trade line to the appropriate broker info object.  Also, clear the
                ' entry rule on the appropriate fill and resave it. (12/18/2008 DAJ)...
                Set rs = g.dbPaper.OpenRecordset("SELECT tblAccountPositionTrades.*, tblAccountPositions.AccountID " & _
                            "FROM [tblAccountPositionTrades] INNER JOIN [tblAccountPositions] ON tblAccountPositionTrades.AccountPositionID=tblAccountPositions.AccountPositionID " & _
                            "WHERE [EntryRuleID]=" & Str(lID) & ";", dbOpenDynaset)
                Do While Not rs.EOF
                    Set TradeLine = g.Broker.GetTradeLine(rs!AccountID, rs!AccountPositionID, rs!TradeNumber)
                    If Not TradeLine Is Nothing Then
                        TradeLine.EntryRuleID = 0&
                        TradeLine.Save
                        
                        Set Fill = New cPtFill
                        If Fill.Load(TradeLine.EntryFillID) Then
                            Fill.EntryRuleIdCategory = 0&
                            Fill.Save
                            g.Broker.RefreshFill Fill
                        End If
                        
                        g.Broker.RefreshTradeLine TradeLine
                    End If
                    
                    rs.MoveNext
                Loop
                            
                fgEntryRules.RemoveItem fgEntryRules.Row
                SaveGrid fgEntryRules
            Else
                lID = CLng(Val(fgExitRules.TextMatrix(fgExitRules.Row, GDCol(eGDCol_TradeRuleID))))
                
                ' Attempt to retrieve tradelines from the appropriate broker info object that had
                ' the exit rule set to the one that was just deleted, clear the exit rule, then
                ' resend the trade line to the appropriate broker info object.  Also, clear the
                ' exit rule on the appropriate fill and resave it. (12/18/2008 DAJ)...
                Set rs = g.dbPaper.OpenRecordset("SELECT tblAccountPositionTrades.*, tblAccountPositions.AccountID " & _
                            "FROM [tblAccountPositionTrades] INNER JOIN [tblAccountPositions] ON tblAccountPositionTrades.AccountPositionID=tblAccountPositions.AccountPositionID " & _
                            "WHERE [ExitRuleID]=" & Str(lID) & ";", dbOpenDynaset)
                Do While Not rs.EOF
                    Set TradeLine = g.Broker.GetTradeLine(rs!AccountID, rs!AccountPositionID, rs!TradeNumber)
                    If Not TradeLine Is Nothing Then
                        TradeLine.ExitRuleID = 0&
                        TradeLine.Save
                        
                        Set Fill = New cPtFill
                        If Fill.Load(TradeLine.ExitFillID) Then
                            Fill.ExitRuleIdCategory = 0&
                            Fill.Save
                            g.Broker.RefreshFill Fill
                        End If
                        
                        g.Broker.RefreshTradeLine TradeLine
                    End If
                    
                    rs.MoveNext
                Loop
                            
                fgExitRules.RemoveItem fgExitRules.Row
                SaveGrid fgExitRules
            End If
            
            m.bChanged = True
            EnableControls
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeRules.RemoveRule"
    
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

    Dim bEnable As Boolean              ' Enable the control?

    If m.nLastFocus = eGDTradeRuleType_Entry Then
        bEnable = ValidRowSelected(fgEntryRules)
    ElseIf m.nLastFocus = eGDTradeRuleType_Exit Then
        bEnable = ValidRowSelected(fgExitRules)
    End If
    
    Enable cmdEdit, bEnable
    Enable mnuEditRule, bEnable
    Enable cmdDelete, bEnable
    Enable mnuDeleteRule, bEnable

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeRules.EnableControls"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ValidRowSelected
'' Description: Determine if the selected row for the given grid is valid
'' Inputs:      Grid
'' Returns:     True if Valid, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidRowSelected(Grid As VSFlexGrid) As Boolean
On Error GoTo ErrSection:

    ValidRowSelected = ((Grid.Row >= Grid.FixedRows) And (Grid.Row < Grid.Rows))

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTradeRules.ValidRowSelected"
    
End Function

