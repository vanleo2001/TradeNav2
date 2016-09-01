VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmStrategyBasket 
   ClientHeight    =   4125
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   8580
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   8580
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraRequiredModule 
      Height          =   255
      Left            =   4680
      TabIndex        =   1
      Top             =   120
      Width           =   2655
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
      Caption         =   "frmStrategyBasket.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmStrategyBasket.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmStrategyBasket.frx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtRequiredModule 
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   0
         Width           =   1215
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmStrategyBasket.frx":0068
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   0
         MultiLine       =   0   'False
         Alignment       =   0
         ScrollBars      =   0
         PasswordChar    =   ""
         TrapTab         =   0   'False
         EnableContextMenu=   -1  'True
         RaiseChangeEvent=   -1  'True
         Tip             =   "frmStrategyBasket.frx":0088
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmStrategyBasket.frx":00A8
      End
      Begin HexUniControls.ctlUniLabelXP lblRequiredModule 
         Height          =   195
         Left            =   0
         Top             =   45
         Width           =   1275
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
         Caption         =   "frmStrategyBasket.frx":00C4
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmStrategyBasket.frx":0106
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmStrategyBasket.frx":0126
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin ActiveToolBars.SSActiveToolBars tbToolbar 
      Left            =   8040
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131083
      ToolBarsCount   =   1
      ToolsCount      =   10
      DisplayContextMenu=   0   'False
      Tools           =   "frmStrategyBasket.frx":0142
      ToolBars        =   "frmStrategyBasket.frx":1176
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   435
      Left            =   360
      TabIndex        =   5
      Top             =   3540
      Width           =   5535
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
      Caption         =   "frmStrategyBasket.frx":13AB
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmStrategyBasket.frx":13D7
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmStrategyBasket.frx":13F7
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdSetAllDates 
         Height          =   435
         Left            =   4260
         TabIndex        =   0
         Top             =   0
         Width           =   1275
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
         Caption         =   "frmStrategyBasket.frx":1413
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmStrategyBasket.frx":144F
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmStrategyBasket.frx":146F
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdRemoveItem 
         Height          =   435
         Left            =   2820
         TabIndex        =   2
         Top             =   0
         Width           =   1275
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
         Caption         =   "frmStrategyBasket.frx":148B
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmStrategyBasket.frx":14C3
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmStrategyBasket.frx":14E3
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdEditItem 
         Height          =   435
         Left            =   1440
         TabIndex        =   7
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
         Caption         =   "frmStrategyBasket.frx":14FF
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmStrategyBasket.frx":1533
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmStrategyBasket.frx":1553
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdNewItem 
         Height          =   435
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   1275
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
         Caption         =   "frmStrategyBasket.frx":156F
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmStrategyBasket.frx":15A1
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmStrategyBasket.frx":15C1
         RightToLeft     =   0   'False
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgStrategyBasketItems 
      Height          =   2895
      Left            =   120
      TabIndex        =   4
      Top             =   420
      Width           =   4575
      _cx             =   8070
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
   Begin HexUniControls.ctlUniLabelXP lblBasket 
      Height          =   195
      Left            =   180
      Top             =   150
      Width           =   4275
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
      Caption         =   "frmStrategyBasket.frx":15DD
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmStrategyBasket.frx":1675
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmStrategyBasket.frx":1695
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Begin VB.Menu mnuNewItem 
         Caption         =   "New Item"
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "Edit Item"
      End
      Begin VB.Menu mnuRemoveItem 
         Caption         =   "Remove Item"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangeFont 
         Caption         =   "Change Font"
      End
   End
End
Attribute VB_Name = "frmStrategyBasket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmStrategyBasket.frm
'' Description: Allow the user to set up a strategy basket
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 11/15/2011   DAJ         Renamed the Strategy Basket stuff
'' 02/16/2012   DAJ         Moved some code to objects, added Set All Dates button
'' 02/21/2012   DAJ         Added contract multiplier column
'' 02/24/2012   DAJ         Fix for issue loading items with different strategy ID for the name
'' 03/09/2012   DAJ         Fix for the last item having a multiplier of zero locking up optimizer
'' 04/03/2013   DAJ         Move Strategy Baskets into the database
'' 05/01/2013   DAJ         Shadow Trading
'' 05/20/2013   DAJ         Fix for 'Set All Dates' changes not taking effect
'' 07/23/2013   DAJ         Add required module UI for strategy baskets
'' 08/05/2013   DAJ         Don't allow user to save strategy basket with an existing name
'' 08/05/2013   DAJ         Ignore case on existing strategy basket name check
'' 08/09/2013   DAJ         Initialize sort column and direction
'' 03/05/2014   DAJ         Removed unused reference to cLotColumn
'' 06/02/2014   DAJ         Fix for running a basket with symbol group with more than 500 items
'' 08/19/2014   DAJ         Fix for sorting in the grid; Expose Strategy Basket Item Inputs;
''                          Allow save when auto trade items in a position; Don't allow save if
''                          removing an item with an auto trade item in a position
'' 06/02/2015   DAJ         Change the 500 limit to kSN_BASKETITEMS ( 99 ); Fix for changing from
''                          a basket over the limit to a basket under the limit
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    bOK As Boolean                      ' Did the user click on OK?
    bStop As Boolean                    ' Stop the run?
    
    Basket As cStrategyBasket           ' Strategy basket object
    
    nPrevColWidth As Long               ' Previous column width
    bAutoSize As Boolean                ' Auto size the grid?

    lSortedCol As Long                  ' Sorted column
    nSortedDir As SortSettings          ' Sort direction for the column
End Type
Private m As mPrivate

Private Enum eGDCols
    eGDCol_System = 0
    eGDCol_SystemNumber
    eGDCol_Symbol
    eGDCol_SymbolGroupID
    eGDCol_Period
    eGDCol_FromDate
    eGDCol_ToDate
    eGDCol_ToEndOfData
    eGDCol_ToDateDisplay
    eGDCol_Multiplier
    eGDCol_SplitAdjust
    eGDCol_Overrides
    eGDCol_SymbolID
    eGDCol_Key
    
    eGDCol_SortKey
    eGDCol_AscSortKey
    eGDCol_DescSortKey
    eGDCol_OutlineLevel
    
    eGDCol_NumCols
End Enum

Const kExtendedCol = eGDCol_System

Private Function GDCol(ByVal lCol As eGDCols) As Long
    GDCol = lCol
End Function

Public Property Get ID() As String
    ID = Str(m.Basket.ID)
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and show the form
'' Inputs:      Strategy Basket ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMe(Optional ByVal lStrategyBasketID As Long = 0&)
On Error GoTo ErrSection:

    Dim bEnableSave As Boolean          ' Enable the save controls?
    Dim bIsOwner As Boolean             ' Is the current user an owner of this object?

    bEnableSave = False
    
    Set m.Basket = New cStrategyBasket
    m.Basket.LoadDb lStrategyBasketID

    If m.Basket.IsGuru = True Then
        bIsOwner = IsOwnerOfGuruObject(m.Basket.LibraryID)
    Else
        bIsOwner = True
    End If
    
    If bIsOwner Then
        ' Initialize the Grid
        fgStrategyBasketItems.Redraw = flexRDNone
        InitGrid
        If lStrategyBasketID > 0& Then
            bEnableSave = Load
        End If
        fgStrategyBasketItems.Redraw = flexRDBuffered
        
        SetEditorCaption Me, "Strategy Basket", m.Basket.Name
        
        ' Hide this button for now...
        tbToolbar.Tools("ID_RunSelection").Visible = False
        
        EnableControls bEnableSave
        m.bOK = False
        
        ShowForm Me, eForm_Nonmodal, frmMain, , ALT_GRID_ROW_COLOR
        
        If fgStrategyBasketItems.Rows = fgStrategyBasketItems.FixedRows Then
            NewBasketItem
        End If
    Else
        InfBox "You are not authorized to view this strategy basket", "!", , "Strategy Basket Error"
        Unload Me
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.ShowMe"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AskToSave
'' Description: Ask the user if they wish to save if changes have been made
'' Inputs:      None
'' Returns:     True if Cancelled, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AskToSave() As Boolean
On Error GoTo ErrSection:
    
    Dim strResponse As String
    
    If tbToolbar.Tools("ID_Save").Enabled Then
        strResponse = InfBox("Do you want to save your changes?", "?", "+Yes|No|-Cancel", Caption)
        Select Case strResponse
            Case "C"
                AskToSave = True
            Case "Y"
                Save "ID_Save"
        End Select
    End If
        
ErrExit:
    Exit Function

ErrSection:
    AskToSave = True
    RaiseError "frmStrategyBasket.AskToSave"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PrintMe
'' Description: Allow the user to print the basket run
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PrintMe()
On Error GoTo ErrSection:

    frmPrintPreview.ShowMe "CNV MultRun", Me, 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.PrintMe"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenerateReport
'' Description: Generate the information to print
'' Inputs:      Args passed through
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GenerateReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:

    Dim lRow As Long
    Dim lCol As Long
    Dim strText As String
    
    With frmPrintPreview.vp
        .StartDoc
        DoPrintHeader
        
        .Font.Name = "Times New Roman"
        .Font.Size = 14
        .Font.Bold = True
        .FontUnderline = True
        .Text = vbLf & "Strategy Basket:"
        .FontUnderline = False
        .Text = "    " & Trim(m.Basket.Name) & vbCrLf '& vbCrLf
        .Font.Bold = False
        .Font.Size = 12
        .Text = "Description: " & Trim(m.Basket.Description) & vbCrLf
        
        .Text = vbLf & vbLf
        
        If frmPrintPreview.GoingToFile Then
            With fgStrategyBasketItems
                For lRow = 0 To .Rows - 1
                    strText = ""
                    For lCol = 0 To .Cols - 1
                        If Not .ColHidden(lCol) Then
                            strText = strText & .Cell(flexcpTextDisplay, lRow, lCol) & vbTab
                        End If
                    Next lCol
                    strText = Left(strText, Len(strText) - 1) ' strip the trailing tab
                    frmPrintPreview.vp.Text = strText & vbCrLf
                Next lRow
            End With
        Else
            .RenderControl = fgStrategyBasketItems.hWnd
        End If

        .EndDoc
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.GenerateReport"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdEditItem_Click
'' Description: Allow the user to edit a new System/Symbol pairing
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdEditItem_Click()
On Error GoTo ErrSection:

    EditBasketItem

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.cmdEditItem_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdNewItem_Click
'' Description: Allow the user to create a new System/Symbol pairing
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdNewItem_Click()
On Error GoTo ErrSection:

    NewBasketItem

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.cmdNewItem_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdRemoveItem_Click
'' Description: Remove the System/Symbol pairing
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdRemoveItem_Click()
On Error GoTo ErrSection:
    
    RemoveBasketItem

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.cmdRemoveItem_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdSetAllDates_Click
'' Description: Allow the user to set the from/to dates the same for all items
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdSetAllDates_Click()
On Error GoTo ErrSection:

    SetAllDates

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.cmdSetAllDates_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgStrategyBasketItems_AfterCollapse
'' Description: After a collapse, make sure to reset the background colors
'' Inputs:      Row Expanded/Collapsed, New State
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgStrategyBasketItems_AfterCollapse(ByVal Row As Long, ByVal State As Integer)
On Error GoTo ErrSection:

    SetBackColors fgStrategyBasketItems

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.fgStrategyBasketItems_AfterCollapse"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgStrategyBasketItems_AfterEdit
'' Description: After user changes multiplier, set the dirty flag
'' Inputs:      Row and Column of the edit, Whether to Cancel the Edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgStrategyBasketItems_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.fgStrategyBasketItems_AfterEdit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgStrategyBasketItems_AfterRowColChange
'' Description: Enable/Disable controls as user changes rows in the grid
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgStrategyBasketItems_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    Dim bNewRowValid As Boolean         ' Is the new row valid?

    With fgStrategyBasketItems
        bNewRowValid = (NewRow >= .FixedRows And NewRow < .Rows)
    
        If .RowOutlineLevel(NewRow) > 0 Then
            Disable cmdRemoveItem
            Disable mnuRemoveItem
        Else
            Enable cmdRemoveItem, bNewRowValid
            Enable mnuRemoveItem, bNewRowValid
        End If
        
        If (NewCol = GDCol(eGDCol_Multiplier)) And (bNewRowValid = True) Then
            .EditCell
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.fgStrategyBasketItems_AfterRowColChange"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgStrategyBasketItems_AfterUserResize
'' Description: Make sure to resize the custom column after a user resize
'' Inputs:      Row and Column of resize
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgStrategyBasketItems_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:
    
    ExtendCustomColumn
    SaveColumns
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.fgStrategyBasketItems_AfterUserResize"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgStrategyBasketItems_BeforeEdit
'' Description: Only allow the user to edit the multiplier column
'' Inputs:      Row and Column of the edit, Whether to Cancel the Edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgStrategyBasketItems_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    Cancel = (Col <> GDCol(eGDCol_Multiplier))

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.fgStrategyBasketItems_BeforeEdit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgStrategyBasketItems_BeforeSort
'' Description: Handle the user sorting a column
'' Inputs:      Column, Sort Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgStrategyBasketItems_BeforeSort(ByVal Col As Long, Order As Integer)
On Error GoTo ErrSection:

    SortOnCol Col, Order

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.fgStrategyBasketItems_BeforeSort"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgStrategyBasketItems_BeforeUserResize
'' Description: Set up for resizing the custom column after the user resize
'' Inputs:      Row and Column of Resize, Whether to Cancel the Resize
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgStrategyBasketItems_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:
    
    ' save current size in case after custom extended column
    m.nPrevColWidth = fgStrategyBasketItems.ColWidth(Col)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.fgStrategyBasketItems_BeforeUserResize"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgStrategyBasketItems_Compare
'' Description: Perform a comparison for the two rows for sorting purposes
'' Inputs:      Row 1, Row 2, Compare Value
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgStrategyBasketItems_Compare(ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
On Error GoTo ErrSection:

    Dim strRow1 As String               ' Value for the first row
    Dim strRow2 As String               ' Value for the second row
    
    strRow1 = fgStrategyBasketItems.TextMatrix(Row1, GDCol(eGDCol_SortKey))
    strRow2 = fgStrategyBasketItems.TextMatrix(Row2, GDCol(eGDCol_SortKey))
    
    If strRow1 = strRow2 Then
        Cmp = 0
    ElseIf strRow1 < strRow2 Then
        Cmp = -1
    Else
        Cmp = 1
    End If
    
    If m.nSortedDir = flexSortStringDescending Then
        Cmp = Cmp * -1
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.fgStrategyBasketItems_Compare"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgStrategyBasketItems_DblClick
'' Description: When the user double clicks on a Run Setup, allow them to edit
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgStrategyBasketItems_DblClick()
On Error GoTo ErrSection:

    With fgStrategyBasketItems
        .Row = .MouseRow
        .RowSel = .Row
    End With
    EditBasketItem

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.fgStrategyBasketItems_DblClick"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgStrategyBasketItems_KeyDown
'' Description: Handle keystrokes in the grid
'' Inputs:      Code of the Key Pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgStrategyBasketItems_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    Select Case KeyCode
        Case vbKeyDelete
            RemoveBasketItem
        Case vbKeyInsert
            NewBasketItem
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.fgStrategyBasketItems_KeyDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgStrategyBasketItems_KeyPress
'' Description: When the user presses Enter on a Run Setup, allow them to edit
'' Inputs:      Key Pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgStrategyBasketItems_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    If KeyAscii = vbKeyReturn Then
        EditBasketItem
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.fgRunSeutp_KeyPress"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgStrategyBasketItems_MouseDown
'' Description: Show the popup menu on a right click in the grid
'' Inputs:      Button pressed, Shift/Ctrl/Alt Status, Mouse Location
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgStrategyBasketItems_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Current mouse row in the grid
    
    With fgStrategyBasketItems
        lMouseRow = .MouseRow
        If lMouseRow >= .FixedRows And lMouseRow < .Rows Then
            If Button = vbRightButton Then
                .Row = lMouseRow
                .RowSel = lMouseRow
                
                PopupMenu mnuPopUp
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.fgStrategyBasketItems_MouseDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgStrategyBasketItems_MouseMove
'' Description: Show appropriate tooltip as the mouse is moved over the grid
'' Inputs:      Button pressed, Shift/Ctrl/Alt status, Mouse Location
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgStrategyBasketItems_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    GridTooltip fgStrategyBasketItems
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgStrategyBasketItems_ValidateEdit
'' Description: After user changes multiplier, set the dirty flag
'' Inputs:      Row and Column of the edit, Whether to Cancel the Edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgStrategyBasketItems_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim strOldValue As String           ' Old value from the cell
    Dim strNewValue As String           ' New value from the cell
    Dim lChildRow As Long               ' Child row
    Dim Item As cStrategyBasketItem     ' Strategy basket item

    With fgStrategyBasketItems
        strOldValue = .TextMatrix(Row, Col)
        strNewValue = .EditText
        
        If Val(strNewValue) < 0 Then
            InfBox "You cannot enter in a multiplier less than zero", "!", , "Error"
            Cancel = True
        Else
            If strOldValue <> strNewValue Then
                EnableControls True
                
                Set Item = BasketItemFromGrid(Row)
                Item.ContractMultiplier = Val(strNewValue)
                .RowData(Row) = Item
                
                lChildRow = .GetNodeRow(Row, flexNTFirstChild)
                Do While lChildRow <> -1&
                    If .TextMatrix(lChildRow, Col) = strOldValue Then
                        .TextMatrix(lChildRow, Col) = strNewValue
                        
                        If TypeOf .RowData(lChildRow) Is cStrategyBasketItem Then
                            Set Item = .RowData(lChildRow)
                            Item.ContractMultiplier = Val(strNewValue)
                            .RowData(lChildRow) = Item
                        End If
                    End If
                    lChildRow = .GetNodeRow(lChildRow, flexNTNextSibling)
                Loop
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.fgStrategyBasketItems_ValidateEdit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_KeyDown
'' Description: Handle certain keystrokes at the form level (e.g. F1 for help)
'' Inputs:      Code of the Key Pressed, Shift/Ctrl/Alt status
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
    RaiseError "frmStrategyBasket.Form_KeyDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Place and initialize the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strPlace As String              ' Form Placement saved off in the ini file
    Dim strFont As String               ' Grid font saved off in the ini file
    
    g.Styler.StyleForm Me
    
    Me.Icon = Picture16(ToolbarIcon("ID_StrategyBaskets"), , True)
    With tbToolbar
        .Tools("ID_Description").Picture = Picture16(ToolbarIcon("ID_News"))
        .Tools("ID_Print").Picture = Picture16(ToolbarIcon("ID_Print"))
        .Tools("ID_Toolbox").Picture = Picture16(ToolbarIcon("ID_Toolbox"))
        .Tools("ID_Save").Picture = Picture16(ToolbarIcon("kSave"))
        .Tools("ID_SaveAs").Picture = Picture16(ToolbarIcon("kSaveAs"))
        .Tools("ID_Rename").Picture = Picture16(ToolbarIcon("kRename"))
        .Tools("ID_Run").Picture = Picture16(ToolbarIcon("ID_Performance"))
        .Tools("ID_RunSelection").Picture = Picture16(ToolbarIcon("ID_Performance"))
        .Tools("ID_Orders").Picture = Picture16(ToolbarIcon("ID_Orders"))
        .Tools("ID_Close").Picture = Picture16(ToolbarIcon("kCancel"))
    End With
    
    ' Place the form
    strPlace = GetIniFileProperty("MultRun", "", "Placement", g.strIniFile)
    If strPlace <> "" Then
        SetFormPlacement Me, strPlace ', "LT"
    Else
        Move Left, Top, 9045, 4815
        CenterTheForm Me
    End If
    
    mnuPopUp.Visible = False
    
    ' Set the Grid Font
    strFont = GetIniFileProperty("MultRun", "", "Fonts", g.strIniFile)
    If strFont <> "" Then FontFromString fgStrategyBasketItems.Font, strFont
    
    fraRequiredModule.Visible = IsIDE
    
    m.lSortedCol = -1&
    m.nSortedDir = -1&
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.Form_Load"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: When the user hits the 'X', unload without saving
'' Inputs:      Whether to Cancel the Unload, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        Cancel = AskToSave
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.Form_QueryUnload"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Resize and move the controls as the form is resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    Dim lMinScaleWidth As Long          ' Minimum scale width for the form
    
    If IsIDE Then
        lMinScaleWidth = lblBasket.Width + fraRequiredModule.Width + (120 * 2)
    Else
        lMinScaleWidth = fraButtons.Width + 120
    End If

    If Not LimitFormSize(Me, lMinScaleWidth, fraButtons.Height * 4) Then
        With fraRequiredModule
            .Move ScaleWidth - .Width
        End With
        
        With fraButtons
            .Move (ScaleWidth - .Width) / 2, ScaleHeight - .Height - 120
        End With
        
        With lblBasket
            '.Move .Left, .Top, ScaleWidth - (.Left * 2)
        End With
        
        With fgStrategyBasketItems
            .Move .Left, lblBasket.Height + (lblBasket.Top * 2), ScaleWidth - (.Left * 2), _
                    ScaleHeight - fraButtons.Height - fraRequiredModule.Height - (.Left * 4)
        End With
        
        ExtendCustomColumn
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: As the form unloads, save the placement and size
'' Inputs:      Whether to Cancel the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    SetIniFileProperty "MultRun", GetFormPlacement(Me), "Placement", g.strIniFile
    SetIniFileProperty "MultRun", FontToString(fgStrategyBasketItems.Font), "Fonts", g.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.Form_Unload"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuChangeFont_Click
'' Description: Allow the user to change the font in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuChangeFont_Click()
On Error GoTo ErrSection:

    ChangeGridFont fgStrategyBasketItems, True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmFunctionMgrCT.mnuChangeFont_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuEditItem_Click
'' Description: Allow the user to edit a basket item from the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuEditItem_Click()
On Error GoTo ErrSection:

    EditBasketItem

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.mnuEditItem_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuNewItem_Click
'' Description: Allow the user to create a new basket item
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuNewItem_Click()
On Error GoTo ErrSection:

    NewBasketItem

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.mnuNewItem_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuRemoveItem_Click
'' Description: Allow the user to remove a basket item from the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuRemoveItem_Click()
On Error GoTo ErrSection:

    RemoveBasketItem

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.mnuRemoveItem_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tbToolbar_ToolClick
'' Description: Handle the user's choice from the toolbar
'' Inputs:      Tool selected on the toolbar
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tbToolbar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
On Error GoTo ErrSection:

    ToggleFocus Me, Me.cmdNewItem
    
    Select Case Tool.ID
        Case "ID_Save", "ID_SaveAs", "ID_Rename"
            Save Tool.ID
        
        Case "ID_Run"
            RunSystems False, False
        
        Case "ID_RunSelection"
            RunSystems True, False
            
        Case "ID_Orders"
            If Tool.State = ssChecked Then
                Tool.Name = "St&op"
                m.bStop = False
                RunSystems False, True
            Else
                m.bStop = True
                Tool.Name = "&Orders"
            End If
            
        Case "ID_Description"
            m.Basket.Description = frmNotes.ShowMe(m.Basket.Description, "Description")
            EnableControls True
        
        Case "ID_Print"
            PrintMe
        
        Case "ID_Toolbox"
            If Not AskToSave Then
                Unload Me
                frmToolbox.ShowMe eTab_StrategyBaskets, m.Basket.Name
            End If
        
        Case "ID_Close"
            If Not AskToSave Then
                Unload Me
            End If
    
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.tbToolbar_ToolClick"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtRequiredModule_Change
'' Description: When the required module changes, set the dirty flag
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtRequiredModule_Change()
On Error GoTo ErrSection:

    EnableControls True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.txtRequiredModule_Change"
    
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
    
    With fgStrategyBasketItems
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExSortShow
        .AllowBigSelection = False
        .AllowSelection = False
        .AllowUserResizing = flexResizeColumns
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .ExtendLastCol = True
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        
        .Rows = 1
        .FixedRows = 1
        .Cols = GDCol(eGDCol_NumCols)
        .FixedCols = 0
        
        .TextMatrix(0, GDCol(eGDCol_System)) = "Strategy"
        .TextMatrix(0, GDCol(eGDCol_Symbol)) = "Symbol(s)"
        .TextMatrix(0, GDCol(eGDCol_Period)) = "Period"
        .TextMatrix(0, GDCol(eGDCol_FromDate)) = "From"
        .TextMatrix(0, GDCol(eGDCol_ToDate)) = "To Date"
        .TextMatrix(0, GDCol(eGDCol_ToEndOfData)) = "To End"
        .TextMatrix(0, GDCol(eGDCol_ToDateDisplay)) = "To"
        .TextMatrix(0, GDCol(eGDCol_Multiplier)) = "Mult"
        .TextMatrix(0, GDCol(eGDCol_SplitAdjust)) = "Split Adjust"
        .TextMatrix(0, GDCol(eGDCol_Overrides)) = "Custom Inputs"
        .TextMatrix(0, GDCol(eGDCol_SymbolID)) = "Symbol ID"
        .TextMatrix(0, GDCol(eGDCol_Key)) = "Key"
        .TextMatrix(0, GDCol(eGDCol_SortKey)) = "Sort Key"
        .TextMatrix(0, GDCol(eGDCol_AscSortKey)) = "Asc Sort Key"
        .TextMatrix(0, GDCol(eGDCol_DescSortKey)) = "Desc Sort Key"
        .TextMatrix(0, GDCol(eGDCol_OutlineLevel)) = "Outline Level"
        
        .ColHidden(GDCol(eGDCol_SymbolGroupID)) = True
        .ColHidden(GDCol(eGDCol_SystemNumber)) = True
        .ColHidden(GDCol(eGDCol_ToDate)) = True
        .ColHidden(GDCol(eGDCol_ToEndOfData)) = True
        .ColHidden(GDCol(eGDCol_SplitAdjust)) = True
        .ColHidden(GDCol(eGDCol_SymbolID)) = True
        .ColHidden(GDCol(eGDCol_Key)) = True
        .ColHidden(GDCol(eGDCol_SortKey)) = True
        .ColHidden(GDCol(eGDCol_AscSortKey)) = True
        .ColHidden(GDCol(eGDCol_DescSortKey)) = True
        .ColHidden(GDCol(eGDCol_OutlineLevel)) = True
        
        .ColDataType(GDCol(eGDCol_ToEndOfData)) = flexDTBoolean
        
        .OutlineBar = flexOutlineBarSimpleLeaf
        .OutlineCol = GDCol(eGDCol_System)
        
        SetUpColumns
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.InitGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RowIsValid
'' Description: Determines if the given row is valid in the grid
'' Inputs:      Row
'' Returns:     True if valid, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RowIsValid(ByVal lRow As Long) As Boolean
On Error GoTo ErrSection:

    RowIsValid = ((lRow >= fgStrategyBasketItems.FixedRows) And (lRow < fgStrategyBasketItems.Rows))

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmStrategyBasket.RowIsValid"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableControls
'' Description: Enable/Disable controls on the form as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableControls(ByVal bEnable As Boolean)
On Error GoTo ErrSection:

    Dim bHasRows As Boolean             ' Does the grid have rows?
    Dim bRowIsValid As Boolean          ' Is the currently selected row valid?

    With fgStrategyBasketItems
        bHasRows = (.Rows > .FixedRows)
        bRowIsValid = RowIsValid(.RowSel)
        
        With tbToolbar
            .Tools("ID_Save").Enabled = bEnable And bHasRows
            .Tools("ID_SaveAs").Enabled = (Len(Trim(m.Basket.Name)) > 0) And bHasRows
            .Tools("ID_Rename").Enabled = (Len(Trim(m.Basket.Name)) > 0) And bHasRows
            .Tools("ID_Run").Enabled = bHasRows
            .Tools("ID_RunSelection").Enabled = bHasRows
        End With
        
        Enable cmdEditItem, bRowIsValid
        Enable cmdRemoveItem, bRowIsValid
        
        If AllFutures Then
            If .TextMatrix(0, GDCol(eGDCol_Multiplier)) <> "# Contracts" Then
                .TextMatrix(0, GDCol(eGDCol_Multiplier)) = "# Contracts"
                SetUpColumns
            End If
        Else
            If .TextMatrix(0, GDCol(eGDCol_Multiplier)) <> "Mult" Then
                .TextMatrix(0, GDCol(eGDCol_Multiplier)) = "Mult"
                SetUpColumns
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.EnableControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Save the current Multiple Run setup
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Save(ByVal strButton As String)
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Index into a for loop
    Dim strText As String               ' Text to show in a message box
    Dim strOldName As String            ' Old name
    Dim strNewName As String            ' Return from the message box
    Dim bSaveAs As Boolean              ' Are we in SaveAs mode?
    Dim Item As cStrategyBasketItem     ' Strategy basket item
    Dim Baskets As cStrategyBaskets     ' Collection of strategy baskets
    Dim bNameExists As Boolean          ' Does the name exist?
    
    If (strButton <> "ID_SaveAs") And (g.TradingItems.IsStrategyBasketAutoTrading(m.Basket.ID) = True) Then
        InfBox "You cannot save this strategy basket because it is currently in an active automated trading item", "!", , "Error"
    ElseIf (strButton <> "ID_SaveAs") And (g.TradingItems.HasNonExistentItemsInPosition(m.Basket.ID, ListOfItemIds(False))) Then
        InfBox "You cannot save this strategy basket because it will remove currently active automated trading item(s) that are in a position", "!", , "Error"
    Else
        If (strButton <> "ID_SaveAs") And (g.TradingItems.IsStrategyBasketInPosition(m.Basket.ID) = True) Then
            InfBox "Saving this strategy basket may cause automated trading items to be in a different position when you reactive them", "i", , "Warning"
        End If
        
        Set Baskets = New cStrategyBaskets
        Baskets.LoadDb
        
        ' Handle Rename/Save As
        strOldName = m.Basket.Name
        strNewName = strOldName
        
        Do
            If (Len(strNewName) = 0) Or ((strButton = "ID_Save") And (bNameExists = True)) Then
                strText = "Save the current Strategy Basket as..."
                strNewName = AskBox("h=Save ; i=? ; g=string ; d=" & strNewName & " ; " & strText)
            ElseIf strButton = "ID_SaveAs" Then
                strText = "Save a copy of the current Strategy Basket as..."
                strNewName = AskBox("h=Save As ; i=? ; g=string ; d=" & strNewName & " ; " & strText)
                If Trim(UCase(strNewName)) <> UCase(strOldName) Then
                    bSaveAs = True
                End If
            ElseIf strButton = "ID_Rename" Then
                strText = "Rename the current Strategy Basket as..."
                strNewName = AskBox("h=Rename ; i=? ; g=string ; d=" & strNewName & " ; " & strText)
            End If
            
            bNameExists = Baskets.NameExists(strNewName, m.Basket.ID, True)
            If bNameExists Then
                InfBox "There is already a strategy basket named|" & strNewName & "|Please choose another name", "!", , "Error"
            End If
        Loop Until bNameExists = False
        
        If Len(Trim(strNewName)) > 0 Then
            If bSaveAs Then
                m.Basket.ClearID
            End If
            m.Basket.Name = Trim(strNewName)
            SetEditorCaption Me, "Strategy Basket", m.Basket.Name
            
            m.Basket.RequiredModule = txtRequiredModule.Text
            BasketItemsFromGrid False
            
            m.Basket.SaveDb
            EnableControls False
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.Save"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Load
'' Description: Load the current Multiple Run setup
'' Inputs:      None
'' Returns:     Did something change?
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Load() As Boolean
On Error GoTo ErrSection:
    
    Dim bReturn As Boolean              ' Return value for the function
    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim Item As cStrategyBasketItem     ' Strategy basket item
    Dim bStrategyChanged As Boolean     ' Stratgy name or ID changed
    Dim bAddedNewItems As Boolean       ' Did we add new items to the grid?
    Dim lRow As Long                    ' Row in the grid
    
    bReturn = False
    With fgStrategyBasketItems
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lIndex = 1 To m.Basket.Items.Count
            Set Item = m.Basket.Items(lIndex)
            If Item.Validate(bStrategyChanged) Then
                If CombinationExists(Item) Then
                    BasketItemToGrid Item
                    
                    If bReturn = False Then
                        bReturn = bStrategyChanged
                    End If
                Else
                    bReturn = True
                End If
            Else
                bReturn = True
            End If
        Next lIndex
                
        SortOnCol
                
        bAddedNewItems = False
        If .Rows > .FixedRows Then
            lRow = .FixedRows
            Do While lRow <> -1&
                Set Item = .RowData(lRow)
                
                If Len(Item.SymbolGroupID) > 0 Then
                    If FillInSymbolGroup(Item.SymbolGroupID, lRow, True) Then
                        bAddedNewItems = True
                    End If
                End If
                
                lRow = .GetNodeRow(lRow, flexNTNextSibling)
            Loop
        End If
        
        If bAddedNewItems Then
            SortOnCol
            bReturn = True
        End If
        
        .Redraw = nRedraw
    End With
    
    With fgStrategyBasketItems
        If .Rows > .FixedRows Then
            .Row = .FixedRows
            .RowSel = .Row
        End If
    End With
    
    txtRequiredModule = m.Basket.RequiredModule
    
    Load = bReturn

ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmStrategyBasket.Load"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BasketItemToGrid
'' Description: Send the given basket item to the grid
'' Inputs:      Strategy Basket Item, Row, Set row data object?
'' Returns:     Row for item
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BasketItemToGrid(ByVal basketItem As cStrategyBasketItem, Optional ByVal lRow As Long = -1&) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim bNewRow As Boolean              ' Is this a new row?
    Dim strPrefix As String             ' Prefix to the sort key

    lReturn = -1&
    With fgStrategyBasketItems
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        bNewRow = False
        If lRow = -1& Then
            lRow = RowForKey(basketItem.Key)
        End If
        
        If lRow = -1& Then
            .Rows = .Rows + 1
            lRow = .Rows - 1
            bNewRow = True
        End If
        
        .RowData(lRow) = basketItem
        
        .TextMatrix(lRow, GDCol(eGDCol_System)) = basketItem.StrategyName
        .TextMatrix(lRow, GDCol(eGDCol_SystemNumber)) = Str(basketItem.StrategyID)
        If Len(basketItem.Symbol) > 0 Then
            .TextMatrix(lRow, GDCol(eGDCol_Symbol)) = basketItem.Symbol
        Else
            .TextMatrix(lRow, GDCol(eGDCol_Symbol)) = basketItem.SymbolGroupName
        End If
        .TextMatrix(lRow, GDCol(eGDCol_SymbolGroupID)) = basketItem.SymbolGroupID
        .TextMatrix(lRow, GDCol(eGDCol_Period)) = basketItem.Period
        .TextMatrix(lRow, GDCol(eGDCol_FromDate)) = DateFormat(basketItem.FromDate)
        .TextMatrix(lRow, GDCol(eGDCol_ToDate)) = DateFormat(basketItem.ToDate)
        .TextMatrix(lRow, GDCol(eGDCol_ToEndOfData)) = Str(CLng(basketItem.ToEndOfData))
        .TextMatrix(lRow, GDCol(eGDCol_ToDateDisplay)) = basketItem.ToDateDisplay
        .TextMatrix(lRow, GDCol(eGDCol_Multiplier)) = Str(basketItem.ContractMultiplier)
        .TextMatrix(lRow, GDCol(eGDCol_SplitAdjust)) = basketItem.SplitDisplay
        .TextMatrix(lRow, GDCol(eGDCol_SymbolID)) = Str(basketItem.SymbolID)
        .TextMatrix(lRow, GDCol(eGDCol_Key)) = basketItem.Key
        .TextMatrix(lRow, GDCol(eGDCol_Overrides)) = basketItem.Overrides
        
        .TextMatrix(lRow, GDCol(eGDCol_SortKey)) = ""
        
        .IsSubtotal(lRow) = True
        .MergeRow(lRow) = False
        
        strPrefix = Pad(basketItem.StrategyName, 50, "L") & Pad(basketItem.SymbolGroupName, 50, "L") & Pad(basketItem.Period, 50, "L")
        If (Len(basketItem.SymbolGroupID) > 0) And (Len(basketItem.Symbol) > 0) Then
            .RowOutlineLevel(lRow) = 1
            .TextMatrix(lRow, GDCol(eGDCol_AscSortKey)) = strPrefix & "_"
            .TextMatrix(lRow, GDCol(eGDCol_DescSortKey)) = strPrefix & "_"
        Else
            .RowOutlineLevel(lRow) = 0
            .TextMatrix(lRow, GDCol(eGDCol_AscSortKey)) = strPrefix & Chr(13)
            .TextMatrix(lRow, GDCol(eGDCol_DescSortKey)) = strPrefix & "}"
        End If
        .TextMatrix(lRow, GDCol(eGDCol_OutlineLevel)) = Str(.RowOutlineLevel(lRow))
        
        lReturn = lRow
        .Redraw = nRedraw
    End With

    BasketItemToGrid = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmStrategyBasket.BasketItemToGrid"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BasketItemFromGrid
'' Description: Get a basket item object from the given row in the grid
'' Inputs:      Row
'' Returns:     Basket Item
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BasketItemFromGrid(ByVal lRow As Long) As cStrategyBasketItem
On Error GoTo ErrSection:

    Dim Item As cStrategyBasketItem     ' Strategy basket item to return
    Dim lParentRow As Long              ' Parent row in the grid
    
    Set Item = New cStrategyBasketItem
    If RowIsValid(lRow) Then
        With fgStrategyBasketItems
            If TypeOf .RowData(lRow) Is cStrategyBasketItem Then
                Set Item = .RowData(lRow)
            Else
                lParentRow = .GetNodeRow(lRow, flexNTParent)
                If lParentRow <> -1& Then
                    If TypeOf .RowData(lParentRow) Is cStrategyBasketItem Then
                        Set Item = .RowData(lParentRow).MakeCopy
                        
                        Item.StrategyBasketID = 0&
                        Item.SymbolOrSymbolID = .TextMatrix(lRow, GDCol(eGDCol_Symbol))
                        Item.ContractMultiplier = Val(.TextMatrix(lRow, GDCol(eGDCol_Multiplier)))
                    End If
                End If
            End If
        End With
    End If
    
    Set BasketItemFromGrid = Item

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmStrategyBasket.BasketItemFromGrid"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RunSystems
'' Description: Run the system/symbol pairs
'' Inputs:      Whether to run just the Selected Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RunSystems(ByVal bJustSelected As Boolean, ByVal bNextBarReport As Boolean)
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Currently selected row
    Dim Item As cStrategyBasketItem     ' Strategy basket item object from the grid
    Dim lJustSelectedID As Long         ' Just the selected ID

    lJustSelectedID = -1&
    If bJustSelected Then
        With fgStrategyBasketItems
            lRow = .RowSel
            If RowIsValid(lRow) Then
                Set Item = .RowData(lRow)
                lJustSelectedID = Item.ID
            End If
        End With
    End If

    BasketItemsFromGrid True
    
    m.Basket.Run bNextBarReport, lJustSelectedID
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.RunSystems"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EditBasketItem
'' Description: Allow the user to edit a new System/Symbol pairing
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EditBasketItem()
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Currently selected row
    Dim Item As cStrategyBasketItem     ' Strategy basket item object from the grid
    Dim OldItem As cStrategyBasketItem  ' Strategy basket item before modification
    Dim lRowOutlineLevel As Long        ' Row outline level for the row
    
    With fgStrategyBasketItems
        lRow = .RowSel
        If RowIsValid(lRow) Then
            Set Item = .RowData(lRow)
            Set OldItem = Item.MakeCopy
            lRowOutlineLevel = .RowOutlineLevel(lRow)
            
            If frmStrategyBasketItem.ShowMe(Item, lRowOutlineLevel) Then
                BasketItemToGrid Item, lRow
                
                If (lRowOutlineLevel = 0) And (Len(OldItem.SymbolGroupID) > 0) Then
                    If (Item.StrategyID <> OldItem.StrategyID) Then
                        RemoveChildren lRow
                        FillInSymbolGroup Item.SymbolGroupID, lRow
                    ElseIf (Item.SymbolGroupID <> OldItem.SymbolGroupID) Then
                        RemoveChildren lRow, Item.SymbolGroupID, Item.SymbolGroupName
                        FillInSymbolGroup Item.SymbolGroupID, lRow, True
                    Else
                        SyncChildren lRow, OldItem
                    End If
                End If
                
                EnableControls True
                
                SetUpColumns
            End If
        End If
    End With
    
    MoveFocus fgStrategyBasketItems

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.EditBasketItem"
    
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

    Dim i&, nTotal&, nDiff&

    With fgStrategyBasketItems
        ' if column being resized is after the extended column,
        ' then change the width of the next visible column instead
        If nResizeCol >= kExtendedCol Then
            .Redraw = flexRDNone
            nDiff = .ColWidth(nResizeCol) - m.nPrevColWidth
            For i = nResizeCol + 1 To .Cols - 1
                If Not .ColHidden(i) Then
                    .ColWidth(i) = .ColWidth(i) - nDiff
                    Exit For
                End If
            Next
            m.nPrevColWidth = 0
        End If
        
        ' size the custom extended column in order to fill the client width
        .ColHidden(kExtendedCol) = True
        .Redraw = flexRDBuffered '(must do this so .ClientWidth will be correct)
        .Redraw = flexRDNone
        nTotal = 0
        For i = 0 To .Cols - 1
            If Not .ColHidden(i) Then
                nTotal = nTotal + .ColWidth(i)
            End If
        Next
        nTotal = .ClientWidth - nTotal
        If nTotal > 0 Then .ColWidth(kExtendedCol) = nTotal
        .ColHidden(kExtendedCol) = False
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.ExtendCustomColumn"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NewBasketItem
'' Description: Allow the user to create a new basket item
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub NewBasketItem()
On Error GoTo ErrSection:
    
    Dim Item As cStrategyBasketItem     ' Strategy basket item
    Dim lRow As Long                    ' Row for the newly added basket item
    
    Set Item = New cStrategyBasketItem
    If frmStrategyBasketItem.ShowMe(Item, 0&) Then
        If RowForKey(Item.Key) = -1& Then
            lRow = BasketItemToGrid(Item)
            FillInSymbolGroup Item.SymbolGroupID, lRow
            
            EnableControls True
            
            SetUpColumns
        Else
            InfBox "That item already exists in the strategy basket", "!", , "Error"
        End If
    End If
    
    MoveFocus fgStrategyBasketItems

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.NewBasketItem"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveBasketItem
'' Description: Allow the user to remove a basket item
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RemoveBasketItem()
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Currently selected row in the grid
    Dim lChildRow As Long               ' Child row
    
    lRow = fgStrategyBasketItems.RowSel
    
    If InfBox("Are you sure you want to remove this strategy basket item?", "?", "+Yes|-No", "Confirmation") = "Y" Then
        With fgStrategyBasketItems
            .Redraw = flexRDNone
            
            lChildRow = .GetNodeRow(lRow, flexNTLastChild)
            Do While lChildRow <> -1&
                .RemoveItem lChildRow
                lChildRow = .GetNodeRow(lRow, flexNTLastChild)
            Loop
            
            .RemoveItem lRow
            
            If lRow >= .Rows Then
                .Row = .Rows - 1
                .RowSel = .Rows - 1
            Else
                .Row = lRow
                .RowSel = lRow
            End If
            
            .Redraw = flexRDBuffered
        End With
        
        EnableControls True
    End If
    
    MoveFocus fgStrategyBasketItems

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.RemoveBasketItem"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetUpColumns
'' Description: Set up the column order/width/visibility according to spec
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetUpColumns()
On Error GoTo ErrSection:

    Dim strFields As String             ' Fields from the ini file
    Dim lIndex As Long                  ' Index into a for loop
    Dim lCol As Long                    ' Index into a for loop
    Dim astrFields As New cGdArray      ' Array of field information
    Dim strColName As String            ' Column Name
    Dim strHidden As String             ' Is Column Hidden?
    Dim lColWidth As Long               ' Width of the column
    Dim lTotWidth As Long               ' Total width from the fields string
    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim lColPos As Long                 ' Column position

    strFields = GetIniFileProperty("Display", "", "StrategyBaskets", g.strIniFile)
    lTotWidth = 0&
    
    With fgStrategyBasketItems
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        astrFields.SplitFields strFields, ","
        For lIndex = 0 To astrFields.Size - 1
            strColName = Parse(astrFields(lIndex), ";", 1)
            strHidden = Parse(astrFields(lIndex), ";", 2)
            lColWidth = CLng(ValOfText(Parse(astrFields(lIndex), ";", 3)))
            
            For lCol = 0 To .Cols - 1
                If UCase(.TextMatrix(0, lCol)) = UCase(strColName) Then
                    .ColPosition(lCol) = lCol
                    lTotWidth = lTotWidth + lColWidth
                    If strHidden = "-1" Then
                        .ColHidden(lCol) = True
                    Else
                        .ColHidden(lCol) = False
                    End If
                    .ColWidth(lCol) = lColWidth
                    
                    Exit For
                End If
            Next lCol
        Next lIndex
        
        If lTotWidth = 0& Then
            .AutoSize 0, .Cols - 1, False, 75
            m.bAutoSize = True
        Else
            m.bAutoSize = False
        End If
        ExtendCustomColumn
        
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.SetUpColumns"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveColumns
'' Description: Save the column information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveColumns()
On Error GoTo ErrSection:

    Dim astrFields As New cGdArray      ' Array of field information
    Dim lIndex As Long                  ' Index into a for loop
    Dim strFields As String             ' Fields to save to ini file

    astrFields.Create eGDARRAY_Strings
    With fgStrategyBasketItems
        For lIndex = 0 To .Cols - 1
            If .ColHidden(lIndex) = True Then
                astrFields.Add .TextMatrix(0, lIndex) & ";-1;" & Str(.ColWidth(lIndex))
            Else
                astrFields.Add .TextMatrix(0, lIndex) & ";0;" & Str(.ColWidth(lIndex))
            End If
        Next lIndex
    End With
    
    strFields = astrFields.JoinFields(",")
    SetIniFileProperty "Display", strFields, "StrategyBaskets", g.strIniFile
    m.bAutoSize = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.SaveColumns"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillInSymbolGroup
'' Description: Fill in the symbols contained in a symbol group as leafs on
''              a tree branch.
'' Inputs:      Symbol Group or Filter ID, Row of Symbol Group or Filter, Keep Item?
'' Returns:     True if changed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FillInSymbolGroup(ByVal strGroupID As String, ByVal lRowOfParent As Long, Optional ByVal bKeepItem As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lFieldNum As Long               ' Field number for the group in the symbol pool
    Dim lIndex As Long                  ' Index into a for loop
    Dim aIndex As cGdArray              ' Indexed list of items in the group
    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim ParentItem As cStrategyBasketItem   ' Strategy basket item for the parent
    Dim ChildItem As cStrategyBasketItem    ' Strategy basket item for the child
    Dim lRow As Long                    ' Row of the child
        
    bReturn = False
    If Len(strGroupID) > 0 Then
        lFieldNum = g.SymbolPool.FieldNumForID(strGroupID)
        If lFieldNum >= 0 Then
            Set aIndex = g.SymbolPool.ArrayTable.CreateIndex(lFieldNum)
            With fgStrategyBasketItems
                lRedraw = .Redraw
                .Redraw = flexRDNone
                
                If aIndex.Size > kSN_BASKETLIMIT Then
                    .Rows = .Rows + 1
                    .RowOutlineLevel(.Rows - 1) = 1
                    .IsSubtotal(.Rows - 1) = True
                    .TextMatrix(.Rows - 1, GDCol(eGDCol_System)) = "This symbol group contains more than " & Str(kSN_BASKETLIMIT) & " symbols and cannot be further customized"
                    .MergeCells = flexMergeSpill
                    .MergeRow(.Rows - 1) = True
                    .RowPosition(.Rows - 1) = lRowOfParent + 1
                Else
                    For lIndex = 0 To aIndex.Size - 1
                        Set ParentItem = .RowData(lRowOfParent)
                        Set ChildItem = ParentItem.MakeCopy
                        ChildItem.SymbolOrSymbolID = g.SymbolPool.SymbolID(aIndex(lIndex))
                        
                        lRow = RowForKey(ChildItem.Key)
                        If (lRow = -1&) Or (bKeepItem = False) Then
                            lRow = BasketItemToGrid(ChildItem)
                            bReturn = True
                        End If
                        .RowPosition(lRow) = lRowOfParent + lIndex + 1
                    Next lIndex
                End If
                
                .IsCollapsed(lRowOfParent) = flexOutlineExpanded
                .Redraw = lRedraw
            End With
        End If
    End If

    FillInSymbolGroup = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmStrategyBasket.FillInSymbolGroup"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ValidSymbolGroup
'' Description: Determine whether the symbol group passed in is still valid
'' Inputs:      Group ID
'' Returns:     True if Valid, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidSymbolGroup(ByVal strGroupID As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    bReturn = True
    If Len(strGroupID) > 0 Then
        If g.SymbolPool.FieldNumForID(strGroupID) = -1 Then
            bReturn = False
        End If
    End If
    
    ValidSymbolGroup = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmStrategyBasket.ValidSymbolGroup"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetNextBarDate
'' Description: Ask the user for the date to use for next bar reports
'' Inputs:      Assume No Position, Ignore Next Bar Data
'' Returns:     Next Bar Date
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetNextBarDate(bAssumeNoPosition As Boolean, bIgnoreNextBarData As Boolean) As Double
On Error GoTo ErrSection:

    Dim dNewYorkTime As Double          ' Current date and time in New York
    Dim dNextBarDate As Double          ' Date (and time) of the next bar report
    Dim lMousePointer As Long           ' Current state of the mouse pointer

    lMousePointer = Screen.MousePointer
    Screen.MousePointer = vbDefault

    ' Come up with an educated guess as to the next bar date...
    dNewYorkTime = ConvertTimeZone(Now)
    If Hour(dNewYorkTime) < 14 Then
        dNextBarDate = Int(dNewYorkTime)
    Else
        dNextBarDate = Int(dNewYorkTime) + 1
    End If
    Do While Not IsWeekday(dNextBarDate)
        dNextBarDate = dNextBarDate + 1
    Loop
    
    ' Verify our educated guess with the user...
    If frmNextBarOpt.ShowMe(dNextBarDate, False, False, bAssumeNoPosition, bIgnoreNextBarData) Then
        GetNextBarDate = dNextBarDate
    Else
        GetNextBarDate = -99999#
    End If

ErrExit:
    Screen.MousePointer = lMousePointer
    Exit Function
    
ErrSection:
    Screen.MousePointer = lMousePointer
    RaiseError "frmStrategyBasket.GetNextBarDate"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SymbolOrSymbolIdForRow
'' Description: Determine the symbol or symbol ID for the given row
'' Inputs:      Row
'' Returns:     Symbol or Symbol ID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SymbolOrSymbolIdForRow(ByVal lRow As Long) As Variant
On Error GoTo ErrSection:

    Dim vReturn As Variant              ' Return value for the function
    Dim lSymbolID As Long               ' Symbol ID

    vReturn = ""
    With fgStrategyBasketItems
        If RowIsValid(lRow) Then
            lSymbolID = CLng(Val(.TextMatrix(lRow, GDCol(eGDCol_SymbolID))))
            If lSymbolID = 0 Then
                vReturn = .TextMatrix(lRow, GDCol(eGDCol_Symbol))
            Else
                vReturn = lSymbolID
            End If
        End If
    End With
    
    SymbolOrSymbolIdForRow = vReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmStrategyBasket.SymbolOrSymbolIdForRow"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetAllDates
'' Description: Allow the user to set the from/to dates the same for all items
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetAllDates()
On Error GoTo ErrSection:

    Dim dFromDate As Double             ' Date to run the strategies from
    Dim dToDate As Double               ' Date to run the strategies to
    Dim bToEndOfData As Boolean         ' Run the strategy through the end of the data?
    Dim nRedraw As RedrawSettings       ' Current state of the grids redraw
    Dim lIndex As Long                  ' Index into a for loop
    Dim Item As cStrategyBasketItem     ' Strategy basket item
    
    With fgStrategyBasketItems
        If RowIsValid(.Row) Then
            dFromDate = CLng(DateOf(.TextMatrix(.Row, GDCol(eGDCol_FromDate))))
            dToDate = CLng(DateOf(.TextMatrix(.Row, GDCol(eGDCol_ToDate))))
            bToEndOfData = CBool(Val(.TextMatrix(.Row, GDCol(eGDCol_ToEndOfData))))
        Else
            dFromDate = Date
            dToDate = Date
            bToEndOfData = True
        End If
    End With
    
    If frmStrategyDates.ShowMe(dFromDate, dToDate, bToEndOfData) = True Then
        With fgStrategyBasketItems
            nRedraw = .Redraw
            .Redraw = flexRDNone
            
            For lIndex = .FixedRows To .Rows - 1
                If TypeOf .RowData(lIndex) Is cStrategyBasketItem Then
                    Set Item = .RowData(lIndex)
                    
                    Item.FromDate = dFromDate
                    Item.ToDate = dToDate
                    Item.ToEndOfData = bToEndOfData
                    
                    BasketItemToGrid Item, lIndex
                End If
            Next lIndex
            
            .Redraw = nRedraw
        End With
        
        EnableControls True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.SetAllDates"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AllFutures
'' Description: Determine if all symbols in the basket are futures
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AllFutures() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    bReturn = True
    With fgStrategyBasketItems
        For lIndex = .FixedRows To .Rows - 1
            If .GetNodeRow(lIndex, flexNTFirstChild) = -1& Then
                If SecurityType(.TextMatrix(lIndex, GDCol(eGDCol_Symbol))) <> "F" Then
                    bReturn = False
                    Exit For
                End If
            End If
        Next lIndex
    End With
    
    AllFutures = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmStrategyBasket.AllFutures"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RowForKey
'' Description: Determine the row for the given key
'' Inputs:      Key
'' Returns:     Row ( -1 if not found )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RowForKey(ByVal strKey As String) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    lReturn = -1&
    With fgStrategyBasketItems
        For lIndex = .FixedRows To .Rows - 1
            If .TextMatrix(lIndex, GDCol(eGDCol_Key)) = strKey Then
                lReturn = lIndex
                Exit For
            End If
        Next lIndex
    End With

    RowForKey = lReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmStrategyBasket.RowForKey"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveChildren
'' Description: Remove the children for the given row
'' Inputs:      Parent Row, New Symbol Group ID, New Symbol Group Name
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RemoveChildren(ByVal lParentRow As Long, Optional ByVal strNewSymbolGroupID As String = "", Optional ByVal strNewSymbolGroupName As String = "")
On Error GoTo ErrSection:
    
    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim lChildRow As Long               ' Child row
    Dim alRowsToDelete As cGdArray      ' Array of rows to delete
    Dim lIndex As Long                  ' Index into a for loop
    Dim ChildItem As cStrategyBasketItem ' Child item from the grid
    Dim lNumInGroup As Long             ' Number of symbols in the new group
    
    Set alRowsToDelete = New cGdArray
    alRowsToDelete.Create eGDARRAY_Longs
    
    lNumInGroup = 0&
    If Len(strNewSymbolGroupID) > 0 Then
        lNumInGroup = g.SymbolPool.NumberRecordsForID(strNewSymbolGroupID)
    End If
        
    With fgStrategyBasketItems
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        lChildRow = .GetNodeRow(lParentRow, flexNTFirstChild)
        Do While lChildRow <> -1&
            If Len(strNewSymbolGroupID) = 0 Then
                alRowsToDelete.Add lChildRow
            ElseIf .MergeRow(lChildRow) = True Then
                alRowsToDelete.Add lChildRow
            Else
                If lNumInGroup > kSN_BASKETLIMIT Then
                    alRowsToDelete.Add lChildRow
                Else
                    Set ChildItem = .RowData(lChildRow)
                
                    If CombinationExists(ChildItem, strNewSymbolGroupID) = False Then
                        alRowsToDelete.Add lChildRow
                    Else
                        ChildItem.SymbolGroupID = strNewSymbolGroupID
                        ChildItem.SymbolGroupName = strNewSymbolGroupName
                        
                        BasketItemToGrid ChildItem, lChildRow
                    End If
                End If
            End If
            
            lChildRow = .GetNodeRow(lChildRow, flexNTNextSibling)
        Loop
        
        For lIndex = alRowsToDelete.Size - 1 To 0 Step -1
            .RemoveItem alRowsToDelete(lIndex)
        Next lIndex
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.RemoveChildren"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SyncChildren
'' Description: Synchronize the children for the given row
'' Inputs:      Parent Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SyncChildren(ByVal lParentRow As Long, ByVal OldParent As cStrategyBasketItem)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim lChildRow As Long               ' Child row
    Dim NewParent As cStrategyBasketItem ' New parent item
    Dim OldChild As cStrategyBasketItem ' Old child item
    Dim NewChild As cStrategyBasketItem ' New child item
    Dim lParm As Long                   ' Index into a for loop
    Dim Parm As cStrategyBasketItemParm ' Strategy basket item parameter
    
    With fgStrategyBasketItems
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        If TypeOf .RowData(lParentRow) Is cStrategyBasketItem Then
            Set NewParent = .RowData(lParentRow)
            
            lChildRow = .GetNodeRow(lParentRow, flexNTFirstChild)
            Do While lChildRow <> -1&
                Set NewChild = NewParent.MakeCopy(True)
                NewChild.SymbolOrSymbolID = .TextMatrix(lChildRow, GDCol(eGDCol_Symbol))
                
                If TypeOf .RowData(lChildRow) Is cStrategyBasketItem Then
                    Set OldChild = .RowData(lChildRow)
                    
                    NewChild.ID = OldChild.ID
                    NewChild.Parms.Clear
                    
                    If OldChild.ContractMultiplier <> OldParent.ContractMultiplier Then
                        NewChild.ContractMultiplier = OldChild.ContractMultiplier
                    End If
                    
                    For lParm = 1 To OldChild.Parms.Count
                        Set Parm = OldChild.Parms(lParm)
                        
                        If NewParent.Parms.Exists(Parm.Key) Then
                            Parm.IsExposed = NewParent.Parms(Parm.Key).IsExposed
                            
                            If Parm.Value = NewParent.Parms(Parm.Key).Value Then
                                ' Old child override value = New parent override value
                                NewChild.Parms.Add Parm, Parm.Key
                            Else
                                If OldParent.Parms.Exists(Parm.Key) Then
                                    If Parm.Value = OldParent.Parms(Parm.Key).Value Then
                                        ' Old child override value <> New parent override value
                                        ' Old child override value = Old parent override value
                                        Parm.Value = NewParent.Parms(Parm.Key).Value
                                        NewChild.Parms.Add Parm, Parm.Key
                                    Else
                                        ' Old child override value <> New parent override value
                                        ' Old child override value <> Old parent override value
                                        NewChild.Parms.Add Parm, Parm.Key
                                    End If
                                Else
                                    ' Old child override value <> New parent override value
                                    ' Old child override doesn't exist in old parent
                                    NewChild.Parms.Add Parm, Parm.Key
                                End If
                            End If
                        Else
                            If OldParent.Parms.Exists(Parm.Key) Then
                                If Parm.Value = OldParent.Parms(Parm.Key).Value Then
                                    ' Old child override doesn't exist in new parent
                                    ' Old child override value = Old parent override value
                                Else
                                    ' Old child override doesn't exist in new parent
                                    ' Old child override value <> Old parent override value
                                    NewChild.Parms.Add Parm, Parm.Key
                                End If
                            Else
                                ' Old child override doesn't exist in new parent
                                ' Old child override doesn't exist in old parent
                                NewChild.Parms.Add Parm, Parm.Key
                            End If
                        End If
                    Next lParm
                    
                    For lParm = 1 To NewParent.Parms.Count
                        Set Parm = NewParent.Parms(lParm)
                        
                        If NewChild.Parms.Exists(Parm.Key) = False Then
                            NewChild.Parms.Add Parm, Parm.Key
                        End If
                    Next lParm
                End If
                
                BasketItemToGrid NewChild, lChildRow
                
                lChildRow = .GetNodeRow(lChildRow, flexNTNextSibling)
            Loop
        End If
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.SyncChildren"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BasketItemsFromGrid
'' Description: Set the basket items from the grid
'' Inputs:      Create Items for Large Basket?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub BasketItemsFromGrid(ByVal bCreateItemsForLargeBasket As Boolean)
On Error GoTo ErrSection:
    
    Dim lRow As Long                    ' Index into a for loop
    Dim Item As cStrategyBasketItem     ' Strategy basket item
    Dim aIndex As cGdArray              ' Indexed list of items in the group
    Dim ParentItem As cStrategyBasketItem   ' Strategy basket item for the parent
    Dim lParentRow As Long              ' Parent row
    Dim lFieldNum As Long               ' Field number for the group in the symbol pool
    Dim lIndex As Long                  ' Index into a for loop

    m.Basket.Items.Clear
    
    With fgStrategyBasketItems
        For lRow = .FixedRows To .Rows - 1
            If TypeOf .RowData(lRow) Is cStrategyBasketItem Then
                Set Item = .RowData(lRow)
                m.Basket.Items.Add Item
            ElseIf (.MergeRow(lRow) = True) And (bCreateItemsForLargeBasket = True) Then
                lParentRow = .GetNodeRow(lRow, flexNTParent)
                If lParentRow <> -1& Then
                    If TypeOf .RowData(lParentRow) Is cStrategyBasketItem Then
                        Set ParentItem = .RowData(lParentRow)
                        lFieldNum = g.SymbolPool.FieldNumForID(ParentItem.SymbolGroupID)
                        If lFieldNum >= 0 Then
                            Set aIndex = g.SymbolPool.ArrayTable.CreateIndex(lFieldNum)
                            For lIndex = 0 To aIndex.Size - 1
                                Set Item = ParentItem.MakeCopy
                                Item.SymbolOrSymbolID = g.SymbolPool.SymbolID(aIndex(lIndex))
                                
                                m.Basket.Items.Add Item
                            Next lIndex
                        End If
                    End If
                End If
            End If
        Next lRow
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.BasketItemsFromGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SortOnCol
'' Description: Sort the grid for the given column number and order
'' Inputs:      Column, Sort Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SortOnCol(Optional ByVal lCol As Long = kNullData, Optional ByVal nOrder As SortSettings = kNullData)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim bAscending As Boolean           ' Do we want to sort ascending?
    Dim lIndex As Long                  ' Index into a for loop
    Dim strValue As String              ' Value for the column being sorted
    Dim strFormat As String             ' Format for the column
    Dim strValue2 As String             ' Value for the child

    If lCol = kNullData Then
        If m.lSortedCol = -1& Then
            lCol = GDCol(eGDCol_Symbol)
        Else
            lCol = m.lSortedCol
        End If
    End If
    
    strFormat = ""
    Select Case lCol
        Case GDCol(eGDCol_FromDate), GDCol(eGDCol_ToDate)
            strFormat = "Date"
            
        Case GDCol(eGDCol_Multiplier)
            strFormat = "Int"
        
    End Select
    
    If nOrder = kNullData Then
        If m.nSortedDir = -1& Then
            nOrder = flexSortStringAscending
        Else
            nOrder = m.nSortedDir
        End If
    End If

    If (nOrder = flexSortGenericAscending) Or (nOrder = flexSortNumericAscending) Or (nOrder = flexSortStringAscending) Or (nOrder = flexSortStringNoCaseAscending) Then
        bAscending = True
    Else
        bAscending = False
    End If

    With fgStrategyBasketItems
        If .Rows > .FixedRows Then
            nRedraw = .Redraw
            .Redraw = flexRDNone
            
            strValue = ""
            strValue2 = ""
            For lIndex = .FixedRows To .Rows - 1
                If .RowOutlineLevel(lIndex) = 0 Then
                    Select Case UCase(strFormat)
                        Case "BOOL"
                            strValue = Pad(Format(Val(CheckedCell(fgStrategyBasketItems, lIndex, lCol)), "0000000"), 50, "R")
                        Case "INT"
                            strValue = Pad(Format(Val(.TextMatrix(lIndex, lCol)), "0000000"), 50, "R")
                        Case "NUMBER", "DATE"
                            strValue = Pad(Format(Val(.TextMatrix(lIndex, lCol)), "#.0000000"), 50, "R")
                        Case "CURRENCY"
                            strValue = Pad(Format(Val(.TextMatrix(lIndex, lCol)), "#.00"), 50, "R")
                        Case Else
                            strValue = Pad(.TextMatrix(lIndex, lCol), 50, "L")
                    End Select
                Else
                    Select Case UCase(strFormat)
                        Case "BOOL"
                            strValue2 = Pad(Format(Val(CheckedCell(fgStrategyBasketItems, lIndex, lCol)), "0000000"), 50, "R")
                        Case "INT"
                            strValue2 = Pad(Format(Val(.TextMatrix(lIndex, lCol)), "0000000"), 50, "R")
                        Case "NUMBER", "DATE"
                            strValue2 = Pad(Format(Val(.TextMatrix(lIndex, lCol)), "#.0000000"), 50, "R")
                        Case "CURRENCY"
                            strValue2 = Pad(Format(Val(.TextMatrix(lIndex, lCol)), "#.00"), 50, "R")
                        Case Else
                            strValue2 = Pad(.TextMatrix(lIndex, lCol), 50, "L")
                    End Select
                End If
                
                If bAscending Then
                    .TextMatrix(lIndex, GDCol(eGDCol_SortKey)) = strValue & "_" & .TextMatrix(lIndex, GDCol(eGDCol_AscSortKey)) & strValue2
                Else
                    .TextMatrix(lIndex, GDCol(eGDCol_SortKey)) = strValue & "_" & .TextMatrix(lIndex, GDCol(eGDCol_DescSortKey)) & strValue2
                End If
                
                .RowOutlineLevel(lIndex) = 0
                .IsSubtotal(lIndex) = False
            Next lIndex
            
            .Select .FixedRows, GDCol(eGDCol_SortKey), .Rows - 1, GDCol(eGDCol_SortKey)
            If bAscending Then
                m.nSortedDir = flexSortStringAscending
            Else
                m.nSortedDir = flexSortStringDescending
            End If
            .Sort = flexSortCustom
            .Select .FixedRows, 0
            
            For lIndex = .FixedRows To .Rows - 1
                .RowOutlineLevel(lIndex) = CLng(Val(.TextMatrix(lIndex, GDCol(eGDCol_OutlineLevel))))
                .IsSubtotal(lIndex) = True
            Next lIndex
            
            If m.lSortedCol > -1& Then
                .Cell(flexcpPicture, 0, m.lSortedCol) = Nothing
            End If
            If bAscending Then
                .Cell(flexcpPicture, 0, lCol) = Picture16("kSortedUpRight")
            Else
                .Cell(flexcpPicture, 0, lCol) = Picture16("kSortedDownRight")
            End If
            
            .Cell(flexcpPictureAlignment, 0, lCol) = flexPicAlignRightTop
            .PicturesOver = True
            
            m.lSortedCol = lCol
            
            SetUpColumns
            .Redraw = nRedraw
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmStrategyBasket.SortOnCol"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CombinationExists
'' Description: Does the given strategy basket item combination still exist?
'' Inputs:      Basket Item, New Symbol Group ID
'' Returns:     True if exists, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CombinationExists(ByVal Item As cStrategyBasketItem, Optional ByVal strNewSymbolGroupID As String = "") As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lFieldNum As Long               ' Field number for the group in the symbol pool
    Dim lRecNum As Long                 ' Record number for the symbol in the symbol pool
    Dim strSymbolGroupID As String      ' Symbol Group ID to use
    
    If Len(strNewSymbolGroupID) = 0 Then
        strSymbolGroupID = Item.SymbolGroupID
    Else
        strSymbolGroupID = strNewSymbolGroupID
    End If
    
    bReturn = False
    If Len(strSymbolGroupID) = 0 Then
        bReturn = True
    Else
        lFieldNum = g.SymbolPool.FieldNumForID(strSymbolGroupID)
        If lFieldNum >= 0 Then
            If Len(Item.Symbol) = 0 Then
                bReturn = True
            Else
                lRecNum = g.SymbolPool.PoolRecForSymbolID(Item.SymbolID)
                If lRecNum >= 0 Then
                    bReturn = g.SymbolPool.SymbolInField(Item.SymbolOrSymbolID, strSymbolGroupID)
                End If
            End If
        End If
    End If
    
    CombinationExists = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmStrategyBasket.CombinationExists"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ListOfItemIds
'' Description: Build a list of ID's of the basket items in the grid
'' Inputs:      Include Parents?
'' Returns:     List of ID's
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ListOfItemIds(ByVal bIncludeParents As Boolean) As cGdArray
On Error GoTo ErrSection:

    Dim alItemIds As cGdArray           ' Array of strategy basket item ID's
    Dim lIndex As Long                  ' Index into a for loop
    Dim basketItem As cStrategyBasketItem ' Strategy basket item from the grid
    
    Set alItemIds = New cGdArray
    alItemIds.Create eGDARRAY_Longs
    
    With fgStrategyBasketItems
        For lIndex = .FixedRows To .Rows - 1
            If (.GetNodeRow(lIndex, flexNTFirstChild) = -1&) Or (bIncludeParents = True) Then
                If TypeOf .RowData(lIndex) Is cStrategyBasketItem Then
                    Set basketItem = .RowData(lIndex)
                    
                    alItemIds.Add basketItem.ID
                End If
            End If
        Next lIndex
    End With
    
    alItemIds.Sort
    
    Set ListOfItemIds = alItemIds

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmStrategyBasket.ListOfItemIds"
    
End Function

