VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmTradeSenseOrders 
   Caption         =   "Form1"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniRichTextBoxXP rtbPreview 
      Height          =   615
      Left            =   180
      TabIndex        =   4
      Top             =   3060
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
      BackColor       =   -2147483643
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmTradeSenseOrders.frx":0000
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
      Tip             =   "frmTradeSenseOrders.frx":0020
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTradeSenseOrders.frx":0040
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
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   1695
      Left            =   2640
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
      Caption         =   "frmTradeSenseOrders.frx":005C
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmTradeSenseOrders.frx":0088
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmTradeSenseOrders.frx":00A8
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdDelete 
         Height          =   495
         Left            =   0
         TabIndex        =   5
         Top             =   1200
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
         Caption         =   "frmTradeSenseOrders.frx":00C4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTradeSenseOrders.frx":00F2
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTradeSenseOrders.frx":0112
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
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
         Caption         =   "frmTradeSenseOrders.frx":012E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTradeSenseOrders.frx":015C
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTradeSenseOrders.frx":017C
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
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
         Caption         =   "frmTradeSenseOrders.frx":0198
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmTradeSenseOrders.frx":01BE
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmTradeSenseOrders.frx":01DE
         RightToLeft     =   0   'False
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgOrders 
      Height          =   2835
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   2235
      _cx             =   3942
      _cy             =   5001
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
End
Attribute VB_Name = "frmTradeSenseOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmTradeSenseOrders.frm
'' Description: Form that handles selecting a Trade Sense order
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 06/16/2010   DAJ         Allow for adding same order multiple times (#5800)
'' 08/12/2010   DAJ         Allow for deleting TradeSense orders
'' 08/12/2010   DAJ         Possible fixes for TN not shutting down correctly
'' 08/23/2010   DAJ         Added required module flag for TradeSense orders/groups
'' 09/09/2010   DAJ         Fix With Block error in EnableControls (#5913)
'' 09/16/2010   DAJ         Enable OK button when Clear is selected
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    bOK As Boolean                      ' Did the user click on OK?
    tsOrders As cTradeSenseOrders       ' Collection of TradeSense orders
    tsGroups As cTradeSenseOrderGroups  ' Collection of TradeSense order groups
    
    bPreviewOnActivate As Boolean       ' Should we update the preview in Form_Activate?
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Setup and show the form
'' Inputs:      Trade Sense Order, Trade Sense Order Collection, Show Clear?,
''              Show Order Number?
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(tsOrder As cTradeSenseOrder, Optional ByVal tsOrders As cTradeSenseOrders = Nothing, Optional ByVal bShowClear As Boolean = False, Optional ByVal bShowNumber As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim iMousePointer As Integer        ' Current state of the screen's mouse pointer

    iMousePointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass

    If tsOrders Is Nothing Then
        Set m.tsOrders = New cTradeSenseOrders
        m.tsOrders.Load
        
        Set m.tsGroups = New cTradeSenseOrderGroups
        m.tsGroups.Load
        
        cmdDelete.Visible = True
    Else
        Set m.tsOrders = tsOrders
        Set m.tsGroups = Nothing
        cmdDelete.Visible = False
    End If

    InitGrid
    LoadGrid bShowClear, bShowNumber
    
    Screen.MousePointer = iMousePointer
    If fgOrders.Rows = 0 Then
        InfBox "There are currently no Trade Sense orders to choose from", "!", , "Trade Sense Orders"
        m.bOK = False
        Set tsOrder = Nothing
    Else
        m.bPreviewOnActivate = True
        
        EnableControls
        ShowForm Me, eForm_Modal, frmMain
        
        If m.bOK Then
            If fgOrders.TextMatrix(fgOrders.Row, 0) = "(Clear)" Then
                Set tsOrder = Nothing
            Else
                Set tsOrder = SelectedOrder
            End If
        End If
    End If
    
    ShowMe = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmTradeSenseOrders.ShowMe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMeOCO
'' Description: Setup and show the form
'' Inputs:      Trade Sense Orders, Trade Sense Order Collection, Show Clear?,
''              Show Order Number?
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMeOCO(tsReturnOrders As cTradeSenseOrders, Optional ByVal tsOrders As cTradeSenseOrders = Nothing, Optional ByVal bShowClear As Boolean = False, Optional ByVal bShowNumber As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim tsOrder As cTradeSenseOrder     ' Trade Sense order object

    If tsOrders Is Nothing Then
        Set m.tsOrders = New cTradeSenseOrders
        m.tsOrders.Load
    Else
        Set m.tsOrders = tsOrders
    End If

    InitGrid True
    LoadGrid bShowClear, bShowNumber
    
    cmdDelete.Visible = False

    If fgOrders.Rows = 0 Then
        InfBox "There are currently no Trade Sense orders to choose from", "!", , "Trade Sense Orders"
        m.bOK = False
        Set tsReturnOrders = Nothing
    Else
        m.bPreviewOnActivate = True
        ShowForm Me, eForm_Modal, frmMain
        
        If m.bOK Then
            If fgOrders.TextMatrix(fgOrders.Row, 0) = "(Clear)" Then
                Set tsReturnOrders = Nothing
            Else
                Set tsReturnOrders = New cTradeSenseOrders
                With fgOrders
                    For lIndex = 0 To .SelectedRows - 1
                        If TypeOf .RowData(.SelectedRow(lIndex)) Is cTradeSenseOrder Then
                            Set tsOrder = .RowData(.SelectedRow(lIndex))
                            tsReturnOrders.Add tsOrder
                        End If
                    Next lIndex
                End With
            End If
        End If
    End If
    
    ShowMeOCO = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmTradeSenseOrders.ShowMeOCO"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Allow the user to cancel the dialog
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    m.bOK = False
    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrders.cmdCancel_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdDelete_Click
'' Description: Allow the user to delete a TradeSense order
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdDelete_Click()
On Error GoTo ErrSection:

    DeleteOrder

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTradeSenseOrders.cmdDelete_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: Allow the user to OK the dialog
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    m.bOK = True
    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrders.cmdOK_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgOrders_AfterRowColChange
'' Description: When the row changes, update the preview
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgOrders_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    If Visible Then
        If NewRow <> OldRow Then
            UpdatePreview
            EnableControls
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrders.fgOrders_AfterRowColChange"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgOrders_DblClick
'' Description: Allow user to select order by using a double click in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgOrders_DblClick()
On Error GoTo ErrSection:

    fgOrders.Row = fgOrders.MouseRow
    m.bOK = True
    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrders.fgOrders_DblClick"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgOrders_KeyDown
'' Description: Allow user to delete an order with the delete key
'' Inputs:      Key Pressed, Shift/Ctrl/Alt Status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgOrders_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyDelete Then
        If cmdDelete.Visible And cmdDelete.Enabled Then
            DeleteOrder
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrders.fgOrders_KeyDown"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgOrders_KeyPress
'' Description: Allow user to select order by pressing Enter in the grid
'' Inputs:      Key Pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgOrders_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    If KeyAscii = vbKeyReturn Then
        m.bOK = True
        Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrders.fgOrders_KeyPress"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: The first time the form is activated, update the preview
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:

    If m.bPreviewOnActivate = True Then
        m.bPreviewOnActivate = False
        UpdatePreview
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrders.Form_Activate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize things when the form is loaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strPlacement As String          ' Placement of the form on the screen
    
    g.Styler.StyleForm Me
    
    strPlacement = GetIniFileProperty("frmTradeSenseOrders", "", "Placement", g.strIniFile)
    If Len(strPlacement) = 0 Then
        CenterTheForm Me
    Else
        SetFormPlacement Me, strPlacement
    End If
    
    rtbPreview.Locked = True
    rtbPreview.BackColor = &H80000000
    
    Caption = "Trade Sense Orders"
    Me.Icon = Picture16(ToolbarIcon("ID_Rules"), , True)
    
    cmdCancel.Cancel = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrders.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user clicks on the X, allow ShowMe to unload the form
'' Inputs:      Cancel Unload?, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        Cancel = True
        m.bOK = False
        Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrders.Form_QueryUnload"
    
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

    Dim lMinScaleWidth As Long          ' Minimum allowed scale width
    Dim lMinScaleHeight As Long         ' Minimum allowed scale height
    Dim lVertSpace As Long              ' Vertical space between controls
    Dim lHorzSpace As Long              ' Horizontal space between controls
    
    lVertSpace = 120
    lHorzSpace = 120
    
    lMinScaleWidth = (fraButtons.Width * 3) + (lHorzSpace * 2)
    lMinScaleHeight = (fraButtons.Height * 2) + rtbPreview.Height + (lVertSpace * 3)
    
    If LimitFormSize(Me, lMinScaleWidth, lMinScaleHeight) = False Then
        With fraButtons
            .Move ScaleWidth - .Width - lHorzSpace, lVertSpace
        End With
        
        With rtbPreview
            .Move lHorzSpace, ScaleHeight - .Height - lVertSpace, ScaleWidth - (lHorzSpace * 2)
        End With
        
        With fgOrders
            .Move lHorzSpace, lVertSpace, ScaleWidth - fraButtons.Width - (lHorzSpace * 3), ScaleHeight - rtbPreview.Height - (lVertSpace * 3)
        End With
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Clean up when the form is unloaded
'' Inputs:      Cancel Unload?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    SetIniFileProperty "frmTradeSenseOrders", GetFormPlacement(Me), "Placement", g.strIniFile

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTradeSenseOrders.Form_Unload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitGrid
'' Description: Initialize the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitGrid(Optional ByVal bMultiSelect As Boolean = False)
On Error GoTo ErrSection:

    With fgOrders
        .Redraw = flexRDNone
        
        SetupGrid fgOrders, eGridMode_List
        .AllowSelection = bMultiSelect
        
        .Rows = 0
        .FixedRows = 0
        .Cols = 2
        .FixedCols = 0
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrders.InitGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGrid
'' Description: Load the grid
'' Inputs:      Show Clear?, Show Order Number?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGrid(Optional ByVal bShowClear As Boolean = False, Optional ByVal bShowNumber As Boolean = False)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim tsOrder As cTradeSenseOrder     ' Trade Sense order object
    
    With fgOrders
        .Redraw = flexRDNone
        
        .Rows = .FixedRows
        
        If bShowClear Then
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = "(Clear)"
            .TextMatrix(.Rows - 1, 1) = ""
        End If
        
        For lIndex = 1 To m.tsOrders.Count
            Set tsOrder = m.tsOrders(lIndex)
            If Not tsOrder Is Nothing Then
                If HasModule(tsOrder.RequiredMod) Then
                    .Rows = .Rows + 1
                    .RowData(.Rows - 1) = tsOrder
                    If bShowNumber Then
                        .TextMatrix(.Rows - 1, 0) = tsOrder.Name & " (" & Str(tsOrder.OrderNumber) & ")"
                    Else
                        .TextMatrix(.Rows - 1, 0) = tsOrder.Name
                    End If
                    .TextMatrix(.Rows - 1, 1) = tsOrder.Action
                End If
            End If
        Next lIndex
        
        .Col = 0
        .Sort = flexSortStringAscending
        
        If .Rows > 0 Then
            .Row = .FixedRows
        End If
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrders.LoadGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SelectedRowValid
'' Description: Is the currently selected row in the grid valid?
'' Inputs:      None
'' Returns:     True if valid row, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SelectedRowValid() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    With fgOrders
        bReturn = (.Row >= .FixedRows) And (.Row < .Rows)
    End With
    
    SelectedRowValid = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTradeSenseOrders.SelectedRowValid"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SelectedOrder
'' Description: Selected order in the grid
'' Inputs:      None
'' Returns:     Order (Nothing if not valid row)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SelectedOrder() As cTradeSenseOrder
On Error GoTo ErrSection:

    Dim tsOrder As cTradeSenseOrder     ' Order to return from the function

    Set tsOrder = Nothing
    If SelectedRowValid Then
        With fgOrders
            If TypeOf .RowData(.Row) Is cTradeSenseOrder Then
                Set tsOrder = .RowData(.Row)
            End If
        End With
    End If
    
    Set SelectedOrder = tsOrder

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTradeSenseOrders.SelectedOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdatePreview
'' Description: Update the preview
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdatePreview()
On Error GoTo ErrSection:

    Dim tsOrder As cTradeSenseOrder     ' Currently selected order
    
    Set tsOrder = SelectedOrder
    If tsOrder Is Nothing Then
        rtbPreview.Text = ""
    Else
        rtbPreview.TextRTF = tsOrder.PreviewRTF
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrders.UpdatePreview"
    
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

    Dim tsOrder As cTradeSenseOrder     ' Currently selected TradeSense order
    
    Enable cmdOK
    
    Set tsOrder = SelectedOrder
    If tsOrder Is Nothing Then
        Disable cmdDelete
    Else
        If m.tsGroups Is Nothing Then
            cmdDelete.Visible = False
        Else
            Enable cmdDelete, (tsOrder.Custom = True) And (m.tsGroups.OrderInGroup(tsOrder.ID) = False)
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrders.EnableControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeleteOrder
'' Description: Delete the selected order if possible
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DeleteOrder()
On Error GoTo ErrSection:

    Dim tsOrder As cTradeSenseOrder     ' TradeSense order
    Dim lRow As Long                    ' Currently selected row

    Set tsOrder = SelectedOrder
    If Not tsOrder Is Nothing Then
        If m.tsGroups.OrderInGroup(tsOrder.ID) Then
            InfBox "You cannot delete '" & tsOrder.Name & "' because it exists in one or more groups", "!", , "Delete Error"
        Else
            If InfBox("Are you sure you want to delete '" & tsOrder.Name & "'?", "?", "+Yes|-No", "Delete Order Confirmation") = "Y" Then
                KillFile tsOrder.FileName
                
                With fgOrders
                    lRow = .Row
                    .RemoveItem .Row
                    
                    If lRow < .Rows Then
                        .Row = lRow
                    ElseIf lRow - 1 >= .FixedRows Then
                        .Row = lRow - 1
                    End If
                End With
                
                EnableControls
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTradeSenseOrders.DeleteOrder"
    
End Sub

