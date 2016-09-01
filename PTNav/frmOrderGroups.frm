VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmOrderGroups 
   Caption         =   "Form1"
   ClientHeight    =   3660
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   3375
      Left            =   3240
      TabIndex        =   1
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
      Caption         =   "frmOrderGroups.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmOrderGroups.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmOrderGroups.frx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdPark 
         Height          =   495
         Left            =   0
         TabIndex        =   6
         Top             =   2250
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
         Caption         =   "frmOrderGroups.frx":0068
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmOrderGroups.frx":0092
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmOrderGroups.frx":00B2
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdNew 
         Height          =   495
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
         Caption         =   "frmOrderGroups.frx":00CE
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmOrderGroups.frx":00F6
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmOrderGroups.frx":0116
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdEdit 
         Height          =   495
         Left            =   0
         TabIndex        =   3
         Top             =   540
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
         Caption         =   "frmOrderGroups.frx":0132
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmOrderGroups.frx":015C
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmOrderGroups.frx":017C
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdDelete 
         Height          =   495
         Left            =   0
         TabIndex        =   4
         Top             =   1080
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
         Caption         =   "frmOrderGroups.frx":0198
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmOrderGroups.frx":01C6
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmOrderGroups.frx":01E6
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdSubmit 
         Height          =   495
         Left            =   0
         TabIndex        =   5
         Top             =   1710
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
         Caption         =   "frmOrderGroups.frx":0202
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmOrderGroups.frx":0230
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmOrderGroups.frx":0250
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdClose 
         Height          =   495
         Left            =   0
         TabIndex        =   7
         Top             =   2880
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
         Caption         =   "frmOrderGroups.frx":026C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmOrderGroups.frx":0298
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmOrderGroups.frx":02B8
         RightToLeft     =   0   'False
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgOrderGroups 
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
      Begin VB.Menu mnuNew 
         Caption         =   "New Order Group"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit Order Group"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Order Group"
      End
      Begin VB.Menu mnuSubmit 
         Caption         =   "Submit Order Group"
      End
      Begin VB.Menu mnuPark 
         Caption         =   "Park Order Group"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangeFont 
         Caption         =   "Change Font"
      End
   End
End
Attribute VB_Name = "frmOrderGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmOrderGroups.frm
'' Description: Allow the user to manage their order groups
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 09/01/2009   DAJ         Use new Parked order status
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Enum eGDCols
    eGDCol_ID = 0
    eGDCol_Name
    eGDCol_Desc
    eGDCol_NumCols
End Enum

Private Enum eGDOrderFields
    eGDOrderField_Buy = 0
    eGDOrderField_Quantity
    eGDOrderField_Symbol
    eGDOrderField_OrderType
    eGDOrderField_StopPrice
    eGDOrderField_LimitPrice
    eGDOrderField_Account
    eGDOrderField_Expiration
End Enum

Private Function GDCol(ByVal Col As eGDCols) As Long
    GDCol = Col
End Function

Private Function OrderField(ByVal field As eGDOrderFields) As Long
    OrderField = field
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Set up the controls and show the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMe()
On Error GoTo ErrSection:

    InitGrid
    LoadGrid
    
    EnableControls

    ' This form calls frmOrderGroup which then calls frmTTEditOrder which gets shown
    ' in an "Act Modal" state, so this must be "Act Modal" as well...
    ShowForm Me, eForm_ActModal, frmMain, , ALT_GRID_ROW_COLOR

ErrExit:
    Unload Me
    Exit Sub

ErrSection:
    Unload Me
    RaiseError "frmOrderGroups.ShowMe", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdClose_Click
'' Description: Close the form
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
    RaiseError "frmOrderGroups.cmdClose.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdDelete_Click
'' Description: Allow the user to delete an existing order group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdDelete_Click()
On Error GoTo ErrSection:

    DeleteOrderGroup

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderGroups.cmdDelete.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdEdit_Click
'' Description: Allow the user to edit an existing order group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdEdit_Click()
On Error GoTo ErrSection:

    EditOrderGroup

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderGroups.cmdEdit.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdNew_Click
'' Description: Allow the user to create a new order group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdNew_Click()
On Error GoTo ErrSection:

    NewOrderGroup

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderGroups.cmdNew.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdPark_Click
'' Description: Allow the user to park an existing order group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdPark_Click()
On Error GoTo ErrSection:
    
    ParkOrderGroup

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderGroups.cmdPark.Click", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdSubmit_Click
'' Description: Allow the user to submit an existing order group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdSubmit_Click()
On Error GoTo ErrSection:

    SubmitOrderGroup

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderGroups.cmdSubmit.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgOrderGroups_BeforeMouseDown
'' Description: Show the popup menu if the user right clicks in the grid
'' Inputs:      Button pressed, Shift/Ctrl/Alt status, Location of mouse
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgOrderGroups_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Current mouse row in the grid

    With fgOrderGroups
        lMouseRow = .MouseRow
        If lMouseRow >= .FixedRows And lMouseRow < .Rows Then
            .Row = lMouseRow
            .RowSel = lMouseRow
        End If
    End With
    
    EnableControls
    
    If Button = vbRightButton Then
        PopupMenu mnuPopUp
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderGroups.fgOrderGroups.BeforeMouseDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgOrderGroups_DblClick
'' Description: Allow the user to edit an item by double clicking on it
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgOrderGroups_DblClick()
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Current mouse row in the grid

    With fgOrderGroups
        lMouseRow = .MouseRow
        If lMouseRow >= .FixedRows And lMouseRow < .Rows Then
            .Row = lMouseRow
            .RowSel = lMouseRow
        End If
    End With
    
    EditOrderGroup

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderGroups.fgOrderGroups.DblClick", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgOrderGroups_KeyUp
'' Description: Allow the user to do certain things with keystrokes
'' Inputs:      Code of the Key Pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgOrderGroups_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    Select Case KeyCode
        Case vbKeyDelete
            DeleteOrderGroup
            
        Case vbKeyInsert
            NewOrderGroup
            
        Case vbKeyReturn
            EditOrderGroup
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderGroups.fgOrderGroups.KeyUp", eGDRaiseError_Show
    Resume ErrExit
    
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

    Dim strPlacement As String          ' Placement for the form
    Dim strFont As String               ' Font from the ini file
    
    g.Styler.StyleForm Me
    
    Icon = Picture16(ToolbarIcon("ID_TradeTracker"), , True)
    Caption = "Order Groups"

    strPlacement = GetIniFileProperty("frmOrderGroups", "", "Placement", g.strIniFile)
    If Len(strPlacement) = 0 Then
        CenterTheForm Me
    Else
        SetFormPlacement Me, strPlacement, "LHTW"
    End If
    
    strFont = GetIniFileProperty("frmOrderGroups", "", "Fonts", g.strIniFile)
    If Len(strFont) > 0 Then FontFromString fgOrderGroups.Font, strFont
    
    mnuPopUp.Visible = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderGroups.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Allow only the ShowMe to unload the form
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
    RaiseError "frmOrderGroups.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Move and resize controls as the form is resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    Dim lMinWidth As Long               ' Minimum scale width
    Dim lMinHeight As Long              ' Minimum scale height
    
    lMinWidth = fraButtons.Width * 5
    lMinHeight = fraButtons.Height + 120
    
    If LimitFormSize(Me, lMinWidth, lMinHeight) Then Exit Sub
    
    With fraButtons
        .Move ScaleWidth - .Width - 60, 60
    End With
    
    With fgOrderGroups
        .Move 60, 60, ScaleWidth - fraButtons.Width - 180, ScaleHeight - 120
    End With

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Save settings and clean up after ourseleves on an unload
'' Inputs:      Cancel the Unload?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    SetIniFileProperty "frmOrderGroups", GetFormPlacement(Me), "Placement", g.strIniFile
    SetIniFileProperty "frmOrderGroups", FontToString(fgOrderGroups.Font), "Fonts", g.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderGroups.Form.Unload", eGDRaiseError_Show
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

    With fgOrderGroups
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .AllowUserResizing = flexResizeNone
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDNone
        .ExplorerBar = flexExSortShow
        .ExtendLastCol = True
        .HighLight = flexHighlightAlways
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        
        .Rows = 1
        .FixedRows = 1
        .Cols = GDCol(eGDCol_NumCols)
        .FixedCols = 0
        
        .TextMatrix(0, GDCol(eGDCol_ID)) = "ID"
        .TextMatrix(0, GDCol(eGDCol_Name)) = "Name"
        .TextMatrix(0, GDCol(eGDCol_Desc)) = "Description"
        
        .ColHidden(GDCol(eGDCol_ID)) = True
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderGroups.InitGrid", eGDRaiseError_Raise
    
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

    Dim rs As Recordset                 ' Recordset into the database

    With fgOrderGroups
        .Redraw = flexRDNone
        
        .Rows = .FixedRows
        
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrderGroups];", dbOpenDynaset)
        Do While Not rs.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, GDCol(eGDCol_ID)) = Str(rs!OrderGroupID)
            .TextMatrix(.Rows - 1, GDCol(eGDCol_Name)) = rs!Name
            .TextMatrix(.Rows - 1, GDCol(eGDCol_Desc)) = rs!Description
            
            rs.MoveNext
        Loop
        
        If .Rows > .FixedRows Then
            .Row = .FixedRows
            .RowSel = .FixedRows
        End If
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderGroups.LoadGrid", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableControls
'' Description: Enable/Disable the controls appropriately
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableControls()
On Error GoTo ErrSection:

    Dim bEnable As Boolean              ' Enable or disable the controls?
    
    With fgOrderGroups
        If .Row >= .FixedRows And .Row < .Rows Then bEnable = True Else bEnable = False
    End With
    
    Enable cmdEdit, bEnable
    Enable mnuEdit, bEnable
    Enable cmdDelete, bEnable
    Enable mnuDelete, bEnable
    Enable cmdSubmit, bEnable
    Enable mnuSubmit, bEnable
    Enable cmdPark, bEnable
    Enable mnuPark, bEnable

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderGroups.EnableControls", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NewOrderGroup
'' Description: Allow the user to create a new order group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub NewOrderGroup()
On Error GoTo ErrSection:

    Dim lOrderGroupID As Long           ' Order Group ID returned from edit form
    Dim rs As Recordset                 ' Recordset into the database
    
    lOrderGroupID = frmOrderGroup.ShowMe(0&)
    If lOrderGroupID <> 0 Then
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrderGroups] WHERE [OrderGroupID]=" & Str(lOrderGroupID) & ";", dbOpenDynaset)
        If Not (rs.BOF And rs.EOF) Then
            With fgOrderGroups
                .Redraw = flexRDNone
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, GDCol(eGDCol_ID)) = Str(rs!OrderGroupID)
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Name)) = rs!Name
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Desc)) = rs!Description
                
                .Row = .Rows - 1
                .RowSel = .Rows - 1
                
                .AutoSize 0, .Cols - 1, False, 75
                .Redraw = flexRDBuffered
            End With
            
            EnableControls
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderGroups.NewOrderGroup", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EditOrderGroup
'' Description: Allow the user to edit an existing order group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EditOrderGroup()
On Error GoTo ErrSection:

    Dim lOrderGroupID As Long           ' Order Group ID returned from edit form
    Dim rs As Recordset                 ' Recordset into the database
    Dim lRow As Long                    ' Row in the grid to edit
    
    With fgOrderGroups
        lRow = .Row
        
        If lRow >= .FixedRows And lRow < .Rows Then
            lOrderGroupID = CLng(Val(.TextMatrix(lRow, GDCol(eGDCol_ID))))
            lOrderGroupID = frmOrderGroup.ShowMe(lOrderGroupID)
            If lOrderGroupID <> 0 Then
                Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrderGroups] WHERE [OrderGroupID]=" & Str(lOrderGroupID) & ";", dbOpenDynaset)
                If Not (rs.BOF And rs.EOF) Then
                    .Redraw = flexRDNone
                    .TextMatrix(lRow, GDCol(eGDCol_ID)) = Str(rs!OrderGroupID)
                    .TextMatrix(lRow, GDCol(eGDCol_Name)) = rs!Name
                    .TextMatrix(lRow, GDCol(eGDCol_Desc)) = rs!Description
                    
                    .Row = lRow
                    .RowSel = lRow
                    
                    .AutoSize 0, .Cols - 1, False, 75
                    .Redraw = flexRDBuffered
                
                    EnableControls
                End If
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderGroups.EditOrderGroup", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeleteOrderGroup
'' Description: Allow the user to delete an existing order group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DeleteOrderGroup()
On Error GoTo ErrSection:

    Dim lOrderGroupID As Long           ' Order Group ID returned from edit form
    Dim rs As Recordset                 ' Recordset into the database
    Dim lRow As Long                    ' Row in the grid to edit
    
    With fgOrderGroups
        lRow = .Row
        
        If lRow >= .FixedRows And lRow < .Rows Then
            lOrderGroupID = CLng(Val(.TextMatrix(lRow, GDCol(eGDCol_ID))))
            If InfBox("Are you sure that you want to delete|" & .TextMatrix(lRow, GDCol(eGDCol_Name)) & "?|", "?", "+Yes|-No", "Confirmation") = "Y" Then
                Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrderGroups] WHERE [OrderGroupID]=" & Str(lOrderGroupID) & ";", dbOpenDynaset)
                If Not (rs.BOF And rs.EOF) Then
                    rs.Delete
                    .RemoveItem lRow
                    
                    If lRow - 1 >= .FixedRows And lRow - 1 < .Rows Then
                        .Row = lRow - 1
                        .RowSel = lRow - 1
                    End If
                    EnableControls
                End If
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderGroups.DeleteOrderGroup", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SubmitOrderGroup
'' Description: Allow the user to submit an existing order group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SubmitOrderGroup()
On Error GoTo ErrSection:

    Dim lOrderGroupID As Long           ' Order Group ID returned from edit form
    Dim rs As Recordset                 ' Recordset into the database
    Dim lRow As Long                    ' Row in the grid to edit
    Dim Orders As New cGdTree           ' Collection of orders to submit
    Dim Order As New cPtOrder           ' Order to add to the collection
    
    With fgOrderGroups
        lRow = .Row
        
        If lRow >= .FixedRows And lRow < .Rows Then
            lOrderGroupID = CLng(Val(.TextMatrix(lRow, GDCol(eGDCol_ID))))
            Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrderGroupItems] WHERE [OrderGroupID]=" & Str(lOrderGroupID) & ";", dbOpenDynaset)
            Do While Not rs.EOF
                Set Order = OrderFromOrderText(rs!OrderText)
                If Len(Order.GenesisOrderID) = 0 Then
                    Order.GenesisOrderID = NextGenesisOrderID(g.Broker.AccountNumberForID(Order.AccountID))
                End If
                Order.Save
                
                Orders.Add Order
                
                rs.MoveNext
            Loop
            
            SubmitMultipleOrders Orders
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderGroups.SubmitOrderGroup", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ParkOrderGroup
'' Description: Allow the user to park an existing order group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ParkOrderGroup()
On Error GoTo ErrSection:

    Dim lOrderGroupID As Long           ' Order Group ID returned from edit form
    Dim rs As Recordset                 ' Recordset into the database
    Dim lRow As Long                    ' Row in the grid to edit
    Dim Order As New cPtOrder           ' Order to add to the collection
    
    With fgOrderGroups
        lRow = .Row
        
        If lRow >= .FixedRows And lRow < .Rows Then
            lOrderGroupID = CLng(Val(.TextMatrix(lRow, GDCol(eGDCol_ID))))
            Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrderGroupItems] WHERE [OrderGroupID]=" & Str(lOrderGroupID) & ";", dbOpenDynaset)
            Do While Not rs.EOF
                Set Order = OrderFromOrderText(rs!OrderText)
                If Len(Order.GenesisOrderID) = 0 Then
                    Order.GenesisOrderID = NextGenesisOrderID(g.Broker.AccountNumberForID(Order.AccountID))
                End If
                Order.Status = eTT_OrderStatus_Parked
                Order.Save
                
                g.Broker.AddOrder Order
                OrderCallback Order
                
                rs.MoveNext
            Loop
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderGroups.ParkOrderGroup", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuChangeFont_Click
'' Description: Allow the user to change the font of the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuChangeFont_Click()
On Error GoTo ErrSection:

    ChangeGridFont fgOrderGroups

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderGroups.mnuChangeFont.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuDelete_Click
'' Description: Allow the user to delete an existing order group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuDelete_Click()
On Error GoTo ErrSection:

    DeleteOrderGroup

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderGroups.mnuDelete.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuEdit_Click
'' Description: Allow the user to edit an existing order group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuEdit_Click()
On Error GoTo ErrSection:

    EditOrderGroup

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderGroups.mnuEdit.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuNew_Click
'' Description: Allow the user to create a new order group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuNew_Click()
On Error GoTo ErrSection:

    NewOrderGroup

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderGroups.mnuNew.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuPark_Click
'' Description: Allow the user to park an existing order group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuPark_Click()
On Error GoTo ErrSection:
    
    ParkOrderGroup

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderGroups.mnuPark.Click", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuSubmit_Click
'' Description: Allow the user to submit an existing order group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuSubmit_Click()
On Error GoTo ErrSection:

    SubmitOrderGroup

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOrderGroups.mnuSubmit.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OrderFromOrderText
'' Description: Create an order from a string of order text
'' Inputs:      Order Text
'' Returns:     Order
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OrderFromOrderText(ByVal strOrder As String) As cPtOrder
On Error GoTo ErrSection:

    Dim Order As New cPtOrder           ' Order to fill in
    Dim astrOrder As New cGdArray       ' Array of order text information
    
    astrOrder.SplitFields strOrder, "|"
    With Order
        .OrderID = -1&
        .Buy = Val(astrOrder(OrderField(eGDOrderField_Buy)))
        .Quantity = CLng(Val(astrOrder(OrderField(eGDOrderField_Quantity))))
        If Val(astrOrder(OrderField(eGDOrderField_Symbol))) = 0 Then
            .SymbolOrSymbolID = astrOrder(OrderField(eGDOrderField_Symbol))
        Else
            .SymbolOrSymbolID = CLng(Val(astrOrder(OrderField(eGDOrderField_Symbol))))
        End If
        .OrderType = CLng(Val(astrOrder(OrderField(eGDOrderField_OrderType))))
        .StopPrice = Val(astrOrder(OrderField(eGDOrderField_StopPrice)))
        .LimitPrice = Val(astrOrder(OrderField(eGDOrderField_LimitPrice)))
        .AccountID = CLng(Val(astrOrder(OrderField(eGDOrderField_Account))))
        .Expiration = CLng(Val(astrOrder(OrderField(eGDOrderField_Expiration))))
    End With
    
    Set OrderFromOrderText = Order

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmOrderGroups.OrderTextFromOrder", eGDRaiseError_Raise
    
End Function

