VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmEditLotContentDetails 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrEditCell 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   660
      Top             =   2580
   End
   Begin VB.Timer tmrMenu 
      Left            =   120
      Top             =   2580
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   495
      Left            =   420
      TabIndex        =   1
      Top             =   2280
      Width           =   2535
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
      Caption         =   "frmEditLotContentDetails.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditLotContentDetails.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditLotContentDetails.frx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Height          =   495
         Left            =   0
         TabIndex        =   3
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
         Caption         =   "frmEditLotContentDetails.frx":0068
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmEditLotContentDetails.frx":008E
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmEditLotContentDetails.frx":00AE
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Height          =   495
         Left            =   1320
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
         Caption         =   "frmEditLotContentDetails.frx":00CA
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmEditLotContentDetails.frx":00F8
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmEditLotContentDetails.frx":0118
         RightToLeft     =   0   'False
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgDetails 
      Height          =   1995
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4395
      _cx             =   7752
      _cy             =   3519
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
      Begin VB.Menu mnuInsertRation 
         Caption         =   "Insert Ration"
      End
      Begin VB.Menu mnuManageIngredients 
         Caption         =   "Manage Ingredients"
      End
   End
End
Attribute VB_Name = "frmEditLotContentDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmEditLotContentDetails.frm
'' Description: Form for allowing user to edit lot content details
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 11/26/2013   DAJ         Tweaks to Turnkey detail editing
'' 12/03/2013   DAJ         Cell alignments
'' 12/04/2013   DAJ         Detail Options
'' 12/05/2013   DAJ         Fix for groups with multiple text fields
'' 12/05/2013   DAJ         Aggregate Column Mode tweaks
'' 12/19/2013   DAJ         "Lauren List" tweaks
'' 01/23/2014   DAJ         Multiple owners per lot
'' 01/31/2014   DAJ         Calculations for feed details
'' 02/25/2014   DAJ         Rations/Ingredients
'' 03/07/2014   DAJ         Moved into NavCattle.DLL
'' 03/14/2014   DAJ         Added support for Boolean lot column type
'' 03/19/2014   DAJ         Allow for default field on startup of the form
'' 03/20/2014   DAJ         Added a "Click Here" line
'' 03/21/2014   DAJ         Fix for goto field ending up on click here line;
''                          goto field when no non click-here lines
'' 04/08/2014   DAJ         Added Average Pay Weight and Cattle Cost per CWT
'' 04/15/2014   DAJ         Fix for automatic edit; new owner lookup form; allow
''                          user to manage ingredients from here
'' 04/24/2014   DAJ         Fix for automatic "goto" adding a new line each time
'' 04/28/2014   DAJ         Don't allow user to delete "Click Here" line
'' 05/22/2014   DAJ         Renamed frmTurnkeyLotContentDetails to frmEditLotContentDetails;
''                          Renamed frmTurnkeyManage to frmCattleManage
'' 05/22/2014   DAJ         Renamed g.Turnkey to g.Cattle
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Enum eGDCols
    eGDCol_FeedYardID = 0
    eGDCol_FeedYardLotID = 1
    eGDCol_Date = 2
    eGDCol_Notes = 3
    
    eGDCol_NumCols
End Enum

Private Type mPrivate
    bOK As Boolean                      ' Did the user click on OK?
    
    LotDetails As cGdTree               ' Collection of lot details
    DetailsByCoord As cGdTree           ' Details with the coordinates as a key
    LotColumn As cLotColumn             ' Lot column information
    LotColumns As cGdTree               ' Collection of lot columns
    lFeedYardID As Long                 ' Feed Yard ID
    lFeedYardLotID As Long              ' Feed Yard Lot ID
    bMult As Boolean                    ' Multiple mode?
    bOwner As Boolean                   ' Owner mode?
    bHasIngredient As Boolean           ' Does this set have ingredient as one of the fields?
    strDefaultKeyValueField As String   ' Default key value field
    bAlreadyDone As Boolean             ' Have we done one time stuff?
    iButton As Integer                  ' Mouse button pressed
    bSkipEditInRowColChange As Boolean  ' Skip the edit in the row/col change event?

    astrOwners As cGdArray              ' Array of owners to pass to dialog
    Owners As cGdTree                   ' Collection of owners
    FeedYard As cBrokerMessage          ' Feedyard object
End Type
Private m As mPrivate

Private Property Get GDCol(ByVal nCol As eGDCols) As Long
    GDCol = nCol
End Property

Private Property Get NotesCol() As Long
    NotesCol = fgDetails.Cols - 1
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ValueForKey
'' Description: Determine the value for the given row and key
'' Inputs:      Row, Key
'' Returns:     Value ( Blank if not found )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Get ValueForKey(ByVal lRow, ByVal strKeyValueField As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim lCol As Long                    ' Column for the given key value field
    
    strReturn = ""
    lCol = ColForKeyValueField(strKeyValueField)
    If (lCol > -1&) And (lRow >= fgDetails.FixedRows) And (lRow < fgDetails.Rows) Then
        strReturn = fgDetails.TextMatrix(lRow, lCol)
    End If
    
    ValueForKey = strReturn

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmEditLotContentDetails.ValueForKey.Get"
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ValueForKey
'' Description: Determine the value for the given row and key
'' Inputs:      Row, Key, Value
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Property Let ValueForKey(ByVal lRow, ByVal strKeyValueField As String, ByVal strValue As String)
On Error GoTo ErrSection:

    Dim lCol As Long                    ' Column for the given key value field
    
    lCol = ColForKeyValueField(strKeyValueField)
    If (lCol > -1&) And (lRow >= fgDetails.FixedRows) And (lRow < fgDetails.Rows) Then
        fgDetails.TextMatrix(lRow, lCol) = strValue
    End If
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmEditLotContentDetails.ValueForKey.Let"
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Setup and show the form
'' Inputs:      Feed Yard ID, Lot ID, Column, Lot Details, Default field
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ByVal lFeedYardID As Long, ByVal lFeedYardLotID As Long, ByVal LotColumn As cLotColumn, LotDetails As cGdTree, Optional ByVal strDefaultKeyValueField As String = "") As Boolean
On Error GoTo ErrSection:

    Set m.LotDetails = LotDetails
    Set m.DetailsByCoord = New cGdTree
    Set m.LotColumn = LotColumn
    Set m.LotColumns = New cGdTree
    m.LotColumns.Add LotColumn
    m.lFeedYardID = lFeedYardID
    Set m.FeedYard = g.Cattle.FeedYards(Str(m.lFeedYardID))
    m.lFeedYardLotID = lFeedYardLotID
    m.bMult = False
    m.bOwner = False
    m.strDefaultKeyValueField = strDefaultKeyValueField
    m.bSkipEditInRowColChange = False
    
    Caption = "Details for " & LotColumn.ColumnHeader
    
    InitGrid
    LoadGrid
    
    ShowForm Me, eForm_Modal, g.frmMain
    
    If m.bOK Then
        DetailsFromGrid
        Set LotDetails = m.LotDetails
    End If
    
    ShowMe = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmEditLotContentDetails.ShowMe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMeCategory
'' Description: Setup and show the form in multiple column mode
'' Inputs:      Feed Yard ID, Lot ID, Columns, Header, Owner mode?,
''              Lot Details, Default field
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMeCategory(ByVal lFeedYardID As Long, ByVal lFeedYardLotID As Long, ByVal LotColumns As cGdTree, ByVal strHeader As String, ByVal bOwner As Boolean, LotDetails As cGdTree, Optional ByVal strDefaultKeyValueField As String = "") As Boolean
On Error GoTo ErrSection:

    Set m.LotDetails = LotDetails
    Set m.DetailsByCoord = New cGdTree
    Set m.LotColumns = LotColumns
    m.lFeedYardID = lFeedYardID
    Set m.FeedYard = g.Cattle.FeedYards(Str(m.lFeedYardID))
    m.lFeedYardLotID = lFeedYardLotID
    m.bMult = True
    m.strDefaultKeyValueField = strDefaultKeyValueField
    
    m.bOwner = bOwner
    If m.bOwner Then
        LoadOwners
    End If
    
    Caption = "Details for " & strHeader
    
    InitGrid
    LoadGrid

    ShowForm Me, eForm_Modal, g.frmMain
    
    If m.bOK Then
        DetailsFromGrid
        Set LotDetails = m.LotDetails
    End If
    
    ShowMeCategory = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmEditLotContentDetails.ShowMeCategory"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Turnkey_Customer
'' Description: Handle a new turnkey customer being added
'' Inputs:      Customer information
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Turnkey_Customer(ByVal turnkeyMessage As cBrokerMessage)
On Error GoTo ErrSection:

    AddOwner turnkeyMessage

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLotContentDetails.Turnkey_Customer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Handle the user clicking on the Cancel button
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
    RaiseError "frmEditLotContentDetails.cmdCancel_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: Handle the user clicking on the OK button
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
    RaiseError "frmEditLotContentDetails.cmdOK_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgDetails_AfterEdit
'' Description: If the user is editing the final row, add another one
'' Inputs:      Index of Grid, Row, Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgDetails_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim LotColumn As cLotColumn         ' Lot column object

    If TypeOf fgDetails.ColData(Col) Is cLotColumn Then
        Set LotColumn = fgDetails.ColData(Col)
        
        Select Case UCase(LotColumn.KeyValueField)
            Case "POUNDSFED"
                If Len(ValueForKey(Row, "DryFeedPct")) = 0 Then
                    ValueForKey(Row, "DryFeedPct") = m.FeedYard("DryFeedPct")
                End If
                CalculateFeedStats Row
                
            Case "DRYFEEDPCT", "FEEDCOSTPERPOUND"
                CalculateFeedStats Row
                
            Case "HEADIN", "TOTALCOSTOFCATTLE", "TOTALPAYWEIGHT"
                CalculateCattleInStats Row
                
        End Select
    End If
    
    ' Do the above before calculating the totals so that the totals get
    ' calculated based ont he updated calculations on the row...
    CalcTotals

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLotContentDetails.fgDetails_AfterEdit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgDetails_AfterRowColChange
'' Description: Handle the user changing cells in the grid
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgDetails_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    If m.bSkipEditInRowColChange = False Then
        If NewCol <> GDCol(eGDCol_Date) Then
DebugLog "fgDetails_AfterRowColChange ( " & Str(NewCol) & " ) -> EditCell"
            'EditCell fgDetails
            tmrEditCell.Enabled = True
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLotContentDetails.fgDetails_AfterRowColChange"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgDetails_BeforeEdit
'' Description: Make sure the user can only edit appropriate cells
'' Inputs:      Row, Column, Cancel Edit?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgDetails_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim LotColumn As cLotColumn         ' Lot column object
    Dim strDetailOptions As String      ' Detail options
    Dim strComboList As String          ' Combo list
    
    If ValidGridRow(fgDetails, Row) Then
        If (RowIsClickHereLine(Row) = False) And (RowIsClickHereRationLine(Row) = False) And (RowIsClickHereIngredientLine(Row) = False) Then
            strComboList = ""
            
            If Col = GDCol(eGDCol_Date) Then
                strComboList = "..."
            ElseIf (Col > GDCol(eGDCol_Date)) And (Col <> NotesCol) Then
                Set LotColumn = fgDetails.ColData(Col)
                
                If (UCase(LotColumn.KeyValueField) = "OWNERNAME") Or (UCase(LotColumn.KeyValueField) = "OWNERNUMBER") Then
                    strComboList = "..."
                ElseIf UCase(LotColumn.Format) = "TEXT" Then
                    If UCase(LotColumn.KeyValueField) = "INGREDIENT" Then
                        strComboList = "|" & g.Cattle.IngredientList.JoinFields("|")
                    Else
                        strDetailOptions = g.Cattle.DetailOptions(Str(LotColumn.ID))
                        If Len(strDetailOptions) > 0 Then
                            strComboList = "|" & strDetailOptions
                        End If
                    End If
                End If
            End If
        
            fgDetails.ComboList = strComboList
        Else
            Cancel = True
        End If
    Else
        Cancel = True
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLotContentDetails.fgDetails_BeforeEdit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgDetails_BeforeMouseDown
'' Description: Bring up the context menu on a right-click in the grid
'' Inputs:      Mouse Button pressed, Shift/Ctrl/Alt status, Mouse location,
''              Cancel the Mouse Down?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgDetails_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    m.iButton = Button
    If (m.bHasIngredient = True) And (Button = vbRightButton) Then
        With fgDetails
            .Row = .MouseRow
            
            PopupMenu mnuPopup
        End With
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmEditLotContentDetails.fgDetails_BeforeMouseDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgDetails_CellButtonClick
'' Description: Handle the user clicking on the "..." button in the cell
'' Inputs:      Row, Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgDetails_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim pt As POINTAPI                  ' Mouse location point
    Dim dDate As Double                 ' Date to send to the edit date form
    Dim bCancelled As Boolean           ' Was the dialog cancelled?
    Dim LotColumn As cLotColumn         ' Lot column object
    Dim strOwner As String              ' Owner

    With fgDetails
        If Col = GDCol(eGDCol_Date) Then
            pt.X = .ColPos(Col) / Screen.TwipsPerPixelX
            pt.Y = (.RowPos(Row) + .RowHeight(Row)) / Screen.TwipsPerPixelY
            ClientToScreen .hWnd, pt
            
            pt.X = pt.X * Screen.TwipsPerPixelX
            pt.Y = pt.Y * Screen.TwipsPerPixelY
            dDate = .Cell(flexcpValue, Row, Col)
            If dDate = 0 Then dDate = Date
            
            frmEditDate.BackColor = BackColor
            dDate = frmEditDate.ShowMe(pt.X, pt.Y, dDate, Me, , , , , , , , , , , bCancelled)
            If bCancelled = False Then
                .TextMatrix(Row, Col) = dDate
            End If
        Else
            Set LotColumn = fgDetails.ColData(Col)
            
            If UCase(LotColumn.KeyValueField) = "OWNERNAME" Then
                'strOwner = g.AppBridge.AccountLookup(m.astrOwners, , .TextMatrix(Row, Col), True)
                strOwner = frmOwnerLookup.ShowMe(m.astrOwners, , .TextMatrix(Row, Col), True)
                If Len(strOwner) > 0 Then
                    SetOwner strOwner, Row
                End If
            ElseIf UCase(LotColumn.KeyValueField) = "OWNERNUMBER" Then
                'strOwner = g.AppBridge.AccountLookup(m.astrOwners, .TextMatrix(Row, Col), , True)
                strOwner = frmOwnerLookup.ShowMe(m.astrOwners, .TextMatrix(Row, Col), , True)
                If Len(strOwner) > 0 Then
                    SetOwner strOwner, Row
                End If
            End If
        End If
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLotContentDetails.fgDetails_CellButtonClick"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgDetails_Click
'' Description: Handle a user click in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgDetails_Click()
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Mouse row in the grid
    Dim lMouseCol As Long               ' Mouse column in the grid
    Dim LotColumn As cLotColumn         ' Lot column object
    
    With fgDetails
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If m.iButton = vbLeftButton Then
            If ValidGridRow(fgDetails, lMouseRow) And ValidGridCol(fgDetails, lMouseCol) Then
                If RowIsClickHereLine(lMouseRow) Then
                    AddRow
                ElseIf RowIsClickHereRationLine(lMouseRow) Then
                    InsertRation
                ElseIf RowIsClickHereIngredientLine(lMouseRow) Then
                    frmCattleManage.ShowMeIngredients
                Else
                    If TypeOf .ColData(lMouseCol) Is cLotColumn Then
                        Set LotColumn = .ColData(lMouseCol)
                        
                        If UCase(LotColumn.Format) = "BOOLEAN" Then
                            CheckedCell(fgDetails, lMouseRow, lMouseCol) = Not CheckedCell(fgDetails, lMouseRow, lMouseCol)
                        End If
                    End If
                End If
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLotContentDetails.fgDetails_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgDetails_KeyUp
'' Description: Allow the user to delete a row
'' Inputs:      Key Code, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgDetails_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    Select Case KeyCode
        Case vbKeyDelete
            With fgDetails
                If ValidDataRow Then
                    If InfBox("Are you sure you want to delete the current row?", "?", "+Yes|-No", "Confirmation") = "Y" Then
                        .RemoveItem .Row
                    End If
                End If
            End With
            
        Case vbKeyInsert
            If m.bHasIngredient = True Then
                InsertRation
            End If
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLotContentDetails.fgDetails_KeyUp"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgDetails_ValidateEdit
'' Description: Validate the user input
'' Inputs:      Row, Column, Cancel Edit?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgDetails_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim LotColumn As cLotColumn         ' Lot column object

    If (Col > GDCol(eGDCol_Date)) And (Col <> NotesCol) Then
        Set LotColumn = fgDetails.ColData(Col)
        
        If (UCase(LotColumn.Format) = "NUMBER") Or (UCase(LotColumn.Format) = "CURRENCY") Then
            If (Len(fgDetails.EditText) > 0) And (IsNumeric(fgDetails.EditText) = False) Then
                InfBox "Please enter in a number as the value", "i", , "Invalid Value"
                Cancel = True
            End If
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLotContentDetails.fgDetails_ValidateEdit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: Handle the form being activated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:

DebugLog "Form_Activate() -- Already Done = " & Str(m.bAlreadyDone)
    If m.bAlreadyDone = False Then
        m.bAlreadyDone = True
        If Len(m.strDefaultKeyValueField) > 0 Then
            GotoField m.strDefaultKeyValueField, True
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmEditLotContentDetails.Form_Activate"
    
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

DebugLog "Form_Load() -- Already Done = " & Str(m.bAlreadyDone)
    Icon = g.AppBridge.Picture16(g.Cattle.IconName)
    
    g.Styler.StyleForm Me
    
    PlaceForm Me
    
    mnuPopup.Visible = False
    
    tmrMenu.Interval = 10
    tmrMenu.Enabled = False

    m.bAlreadyDone = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLotContentDetails.Form_Load"
    
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
        m.bOK = False
        Cancel = True
        Hide
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmEditLotContentDetails.Form_QueryUnload"
    
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

    Dim lMinScaleWidth As Long          ' Minimum scale width
    Dim lMinScaleHeight As Long         ' Minimum scale height
    Dim lSpace As Long                  ' Space between controls
    
    lSpace = 60
    lMinScaleWidth = fraButtons.Width + (lSpace * 2)
    lMinScaleHeight = (fraButtons.Height * 3) + (lSpace * 3)
    
    If LimitFormSize(Me, lMinScaleWidth, lMinScaleHeight) = False Then
        With fraButtons
            .Move (ScaleWidth - .Width) / 2, ScaleHeight - .Height - lSpace
        End With
        With fgDetails
            .Move lSpace, lSpace, ScaleWidth - (lSpace * 2), fraButtons.Top - (lSpace * 2)
        End With
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
    tmrMenu.Enabled = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLotContentDetails.Form_Unload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuInsertRation_Click
'' Description: Insert a ration into the feed grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuInsertRation_Click()
On Error GoTo ErrSection:

    tmrMenu.Tag = "INSERTRATION"
    tmrMenu.Enabled = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLotContentDetails.mnuInsertRation_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuManageIngredients_Click
'' Description: Allow the user to manage ingredients
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuManageIngredients_Click()
On Error GoTo ErrSection:
    
    tmrMenu.Tag = "MANAGEINGREDIENTS"
    tmrMenu.Enabled = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLotContentDetails.mnuManageIngredients_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tmrEditCell_Timer
'' Description: Handle a edit cell
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrEditCell_Timer()
On Error GoTo ErrSection:

    tmrEditCell.Enabled = False
    EditCell fgDetails

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLotContentDetails.tmrEditCell_Timer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tmrMenu_Timer
'' Description: Handle a menu item off the timer
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tmrMenu_Timer()
On Error GoTo ErrSection:

    Dim strTag As String                ' Tag of the timer control
    Static bInProgress As Boolean       ' Are we currently performing a command?

    g.AppBridge.TimerStart "frmEditLotContentDetails.tmrMenu"
    If bInProgress = False Then
        bInProgress = True
        
        strTag = tmrMenu.Tag
        tmrMenu.Tag = ""
        tmrMenu.Enabled = False
        
        Select Case UCase(Parse(strTag, vbTab, 1))
            Case "INSERTRATION"
                InsertRation
                
            Case "MANAGEINGREDIENTS"
                frmCattleManage.ShowMeIngredients
                
        End Select
        
        bInProgress = False
    End If
    g.AppBridge.TimerEnd "frmEditLotContentDetails.tmrMenu", tmrMenu.Interval

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLotContentDetails.tmrMenu_Timer"
    
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
    Dim lIndex As Long                  ' Index into a for loop
    Dim LotColumn As cLotColumn         ' Lot column object
    Dim lCol As Long                    ' Column in the grid
    
    With fgDetails
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExNone
        .ExtendLastCol = True
        .MergeCells = flexMergeFree
        .OutlineBar = flexOutlineBarNone
        .ScrollTrack = True
        .SelectionMode = flexSelectionFree
        .SheetBorder = RGB(128, 128, 128)
        .TabBehavior = flexTabCells
        .WordWrap = False
        
        .FixedRows = 2
        .Rows = 2
        .FixedCols = 0
        .Cols = GDCol(eGDCol_NumCols) + m.LotColumns.Count
        
        .TextMatrix(0, GDCol(eGDCol_FeedYardID)) = "Feed Yard ID"
        .TextMatrix(0, GDCol(eGDCol_FeedYardLotID)) = "Feed Yard Lot ID"
        .TextMatrix(0, GDCol(eGDCol_Date)) = "Date"
        .TextMatrix(1, GDCol(eGDCol_Date)) = "Totals"
        
        m.bHasIngredient = False
        For lIndex = 1 To m.LotColumns.Count
            Set LotColumn = m.LotColumns(lIndex)
            
            lCol = GDCol(eGDCol_Date) + lIndex
            .ColData(lCol) = LotColumn
            .TextMatrix(0, lCol) = LotColumn.ColumnHeader
            .ColAlignment(lCol) = flexAlignRightCenter
            
            If UCase(LotColumn.Format) = "DATE" Then
                .ColHidden(lCol) = True
                .ColAlignment(lCol) = flexAlignCenterCenter
            ElseIf UCase(LotColumn.Format) = "TEXT" Then
                .ColAlignment(lCol) = flexAlignLeftCenter
            Else
                .ColAlignment(lCol) = flexAlignRightCenter
            End If
            If UCase(LotColumn.Format) = "BOOLEAN" Then
                .ColDataType(lCol) = flexDTBoolean
            Else
                .ColFormat(lCol) = LotColumn.DisplayFormat
            End If
            
            .ColHidden(lCol) = LotColumn.AlwaysHidden Or LotColumn.FeedyardHidden
            
            If UCase(LotColumn.KeyValueField) = "INGREDIENT" Then
                m.bHasIngredient = True
            End If
        Next lIndex
        
        .TextMatrix(0, NotesCol) = "Notes"
        
        .ColAlignment(GDCol(eGDCol_Date)) = flexAlignCenterCenter
        .ColAlignment(NotesCol) = flexAlignLeftCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftCenter
        
        .ColFormat(GDCol(eGDCol_Date)) = DateFormat("Format", MM_DD_YYYY)
        
        .ColHidden(GDCol(eGDCol_Date)) = m.bOwner
        .ColHidden(GDCol(eGDCol_FeedYardID)) = True
        .ColHidden(GDCol(eGDCol_FeedYardLotID)) = True
        .ColHidden(NotesCol) = m.bOwner
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLotContentDetails.InitGrid"
    
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

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim lIndex As Long                  ' Index into a for loop
    Dim Detail As cBrokerMessage        ' Lot content detail object
    Dim lRow As Long                    ' Row in the grid for the date
    Dim lCol As Long                    ' Column in the grid for the lot column id
    Dim LotColumn As cLotColumn         ' Lot column object
    
    With fgDetails
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Rows = .FixedRows
        
        For lIndex = 1 To m.LotDetails.Count
            Set Detail = m.LotDetails(lIndex)
            
            lRow = RowForDate(Detail("Date"))
            lCol = ColForLotColumnID(Detail("LotColumnID"))
            
            If lCol >= 1& Then
                Set LotColumn = .ColData(lCol)
                
                If lRow = -1& Then
                    .Rows = .Rows + 1
                    lRow = .Rows - 1
                    
                    .MergeRow(lRow) = False
                                        
                    .TextMatrix(lRow, GDCol(eGDCol_FeedYardID)) = Detail("FeedYardID")
                    .TextMatrix(lRow, GDCol(eGDCol_FeedYardLotID)) = Detail("FeedYardLotID")
                    .TextMatrix(lRow, GDCol(eGDCol_Date)) = Detail("Date")
                End If
                
                m.DetailsByCoord.Add Detail, Str(lRow) & ";" & Str(lCol)
                
                g.Cattle.GridValue(fgDetails, lRow, lCol, LotColumn) = Detail("Value")
                
                .TextMatrix(lRow, NotesCol) = Detail("Notes")
            End If
        Next lIndex
        
        If .Rows > .FixedRows Then
            .Select .FixedRows, GDCol(eGDCol_Date), .Rows - 1, GDCol(eGDCol_Date)
            .Sort = flexSortNumericAscending
            .Select 0, 0
        End If
        
        CalcTotals
        
        AddClickHereLine
        AddClickHereRationLine
        AddClickHereIngredientLine
        
        .AutoSize 0, .Cols - 1, False, 250
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLotContentDetails.LoadGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DetailsFromGrid
'' Description: Rebuild the details collection from the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DetailsFromGrid()
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Index into a for loop
    Dim Detail As cBrokerMessage        ' Lot content detail object
    Dim lCol As Long                    ' Column in the grid
    Dim LotColumn As cLotColumn         ' Lot column object
    Dim OldDetail As cBrokerMessage     ' Old detail record
    
    m.LotDetails.Clear
    With fgDetails
        For lRow = .FixedRows To .Rows - 1
            For lCol = GDCol(eGDCol_Date) + 1 To .Cols - 2
                If (RowIsClickHereLine(lRow) = False) And (RowIsClickHereRationLine(lRow) = False) And (RowIsClickHereIngredientLine(lRow) = False) Then
                    If Len(Trim(.TextMatrix(lRow, lCol))) > 0 Then
                        If TypeOf .ColData(lCol) Is cLotColumn Then
                            Set LotColumn = .ColData(lCol)
                            
                            If UCase(LotColumn.Format) <> "DATE" Then
                                Set Detail = New cBrokerMessage
                                
                                If m.DetailsByCoord.Exists(Str(lRow) & ";" & Str(lCol)) Then
                                    Set OldDetail = m.DetailsByCoord(Str(lRow) & ";" & Str(lCol))
                                    
                                    Detail.Add "ID", OldDetail("ID")
                                    Detail.Add "LotContentID", OldDetail("LotContentID")
                                Else
                                    Detail.Add "ID", ""
                                    Detail.Add "LotContentID", ""
                                End If
                                
                                Detail.Add "FeedYardID", .TextMatrix(lRow, GDCol(eGDCol_FeedYardID))
                                Detail.Add "FeedYardLotID", .TextMatrix(lRow, GDCol(eGDCol_FeedYardLotID))
                                Detail.Add "LotColumnID", Str(LotColumn.ID)
                                If m.bOwner Then
                                    Detail.Add "Date", Str(0 + (CDbl(lRow) / 1440#))
                                Else
                                    Detail.Add "Date", Str(CDbl(Int(.Cell(flexcpValue, lRow, GDCol(eGDCol_Date)))) + (CDbl(lRow) / 1440#))
                                End If
                                Detail.Add "Value", g.Cattle.GridValue(fgDetails, lRow, lCol, LotColumn)
                                Detail.Add "Notes", .TextMatrix(lRow, NotesCol)
                                
                                m.LotDetails.Add Detail
                            End If
                        End If
                    End If
                End If
            Next lCol
        Next lRow
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLotContentDetails.DetailsFromGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddRow
'' Description: Add a row to the grid
'' Inputs:      None
'' Returns:     New Row
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AddRow() As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lClickHereLine As Long          ' Click here line

    With fgDetails
        .Rows = .Rows + 1
        
        .MergeRow(.Rows - 1) = False
        
        .TextMatrix(.Rows - 1, GDCol(eGDCol_FeedYardID)) = Str(m.lFeedYardID)
        .TextMatrix(.Rows - 1, GDCol(eGDCol_FeedYardLotID)) = Str(m.lFeedYardLotID)
        .TextMatrix(.Rows - 1, GDCol(eGDCol_Date)) = Str(Date)
        
        lClickHereLine = ClickHereLine
        If lClickHereLine = -1& Then
            lReturn = .Rows - 1
        Else
            .RowPosition(.Rows - 1) = lClickHereLine
            lReturn = lClickHereLine
        End If
    End With
    
    AddRow = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmEditLotContentDetails.AddRow"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RowForDate
'' Description: Determine the row for the given date
'' Inputs:      Date
'' Returns:     Row ( -1 if not found )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RowForDate(ByVal strDate As String) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    lReturn = -1&
    With fgDetails
        For lIndex = .FixedRows To .Rows - 1
            If .TextMatrix(lIndex, GDCol(eGDCol_Date)) = strDate Then
                lReturn = lIndex
                Exit For
            End If
        Next lIndex
    End With
    
    RowForDate = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmEditLotContentDetails.RowForDate"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ColForLotColumnID
'' Description: Determine the column for the given lot column ID
'' Inputs:      Lot Column ID
'' Returns:     Column ( -1 if not found )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ColForLotColumnID(ByVal strLotColumnID As String) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim LotColumn As cLotColumn         ' Lot column object
    
    lReturn = -1&
    With fgDetails
        For lIndex = GDCol(eGDCol_Date) + 1 To .Cols - 2
            If TypeOf .ColData(lIndex) Is cLotColumn Then
                Set LotColumn = .ColData(lIndex)
                
                If LotColumn.ID = CLng(Val(strLotColumnID)) Then
                    lReturn = lIndex
                    Exit For
                End If
            End If
        Next lIndex
    End With
    
    ColForLotColumnID = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmEditLotContentDetails.ColForLotColumnID"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ColForKeyValueField
'' Description: Determine the column for the given key value field
'' Inputs:      Key Value Field
'' Returns:     Column ( -1 if not found )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ColForKeyValueField(ByVal strKeyValueField As String) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim LotColumn As cLotColumn         ' Lot column object
    
    lReturn = -1&
    With fgDetails
        For lIndex = GDCol(eGDCol_Date) + 1 To .Cols - 2
            If TypeOf .ColData(lIndex) Is cLotColumn Then
                Set LotColumn = .ColData(lIndex)
                
                If UCase(LotColumn.KeyValueField) = UCase(strKeyValueField) Then
                    lReturn = lIndex
                    Exit For
                End If
            End If
        Next lIndex
    End With
    
    ColForKeyValueField = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmEditLotContentDetails.ColForKeyValueField"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalcTotals
'' Description: Calculate the totals for the appropriate columns
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalcTotals()
On Error GoTo ErrSection:

    Dim lCol As Long                    ' Index into a for loop
    Dim LotColumn As cLotColumn         ' Lot column object
    Dim strTotal As String              ' Total for the column
    Dim bCalcCattleInStats As Boolean   ' Calculate the "cattle in" stats?
    
    With fgDetails
        DetailsFromGrid

        bCalcCattleInStats = False
        For lCol = GDCol(eGDCol_Date) + 1 To .Cols - 2
            If TypeOf .ColData(lCol) Is cLotColumn Then
                Set LotColumn = .ColData(lCol)
                
                strTotal = g.Cattle.CalcTotalForDetails(m.LotDetails, LotColumn)
                g.Cattle.GridValue(fgDetails, 1, lCol, LotColumn) = strTotal
                
                If (UCase(LotColumn.KeyValueField) = "AVERAGEPAYWEIGHT") Or (UCase(LotColumn.KeyValueField) = "PRICEINPERCWT") Then
                    bCalcCattleInStats = True
                End If
            End If
        Next lCol
        
        If bCalcCattleInStats = True Then
            CalculateCattleInStats 1
        End If
    
        .AutoSize 0, .Cols - 1, False, 250
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLotContentDetails.CalcTotals"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadOwners
'' Description: Load the owners collections
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadOwners()
On Error GoTo ErrSection:
    
    Dim Owners As cGdTree               ' Collection of owners
    Dim lIndex As Long                  ' Index into a for loop
    
    Set Owners = g.Cattle.Customers
    
    Set m.Owners = New cGdTree
    Set m.astrOwners = New cGdArray
    m.astrOwners.Create eGDARRAY_Strings
    
    For lIndex = 1 To Owners.Count
        AddOwner Owners(lIndex)
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLotContentDetails.LoadOwners"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetOwner
'' Description: Set the owner
'' Inputs:      Owner Number, Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetOwner(ByVal strOwnerNumber As String, ByVal lRow As Long)
On Error GoTo ErrSection:

    Dim strOwnerName As String          ' Owner name
    Dim lIndex As Long                  ' Index into a for loop
    Dim LotColumn As cLotColumn         ' Lot column object
    
    strOwnerName = m.Owners(strOwnerNumber)
    
    With fgDetails
        For lIndex = GDCol(eGDCol_Date) + 1 To .Cols - 2
            If TypeOf .ColData(lIndex) Is cLotColumn Then
                Set LotColumn = .ColData(lIndex)
                
                If UCase(LotColumn.KeyValueField) = "OWNERNAME" Then
                    .TextMatrix(lRow, lIndex) = strOwnerName
                ElseIf UCase(LotColumn.KeyValueField) = "OWNERNUMBER" Then
                    .TextMatrix(lRow, lIndex) = strOwnerNumber
                End If
            End If
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLotContentDetails.SetOwner"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddOwner
'' Description: Add the given owner to the appropriate collections if not there
'' Inputs:      Owner
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddOwner(ByVal Owner As cBrokerMessage)
On Error GoTo ErrSection:

    Dim strOwner As String              ' Owner information
    Dim lPos As Long                    ' Position in the array

    If m.Owners.Exists(Owner("Number")) = False Then
        m.Owners.Add Owner("Name"), Owner("Number")
    End If

    strOwner = Owner("Number") & vbTab & vbTab & Owner("Name")
    If m.astrOwners.BinarySearch(strOwner, lPos) = False Then
        m.astrOwners.Add strOwner, lPos
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLotContentDetails.AddOwner"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalculateFeedStats
'' Description: Calculate the feed statistics for the given row
'' Inputs:      Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalculateFeedStats(ByVal lRow As Long)
On Error GoTo ErrSection:

    Dim dPoundsFed As Double            ' Total pounds fed
    Dim dDryFeedPct As Double           ' Percent of feed that is dry
    Dim dDryPoundsFed As Double         ' Dry pounds fed
    Dim dFeedCostPerPound As Double     ' Feed cost per pound
    Dim dTotalFeedCost As Double        ' Total feed cost
    
    dPoundsFed = ValOfText(ValueForKey(lRow, "PoundsFed"))
    dDryFeedPct = ValOfText(ValueForKey(lRow, "DryFeedPct"))
    dFeedCostPerPound = ValOfText(ValueForKey(lRow, "FeedCostPerPound"))
    
    ValueForKey(lRow, "DryPoundsFed") = Str(dPoundsFed * (dDryFeedPct / 100))
    ValueForKey(lRow, "TotalFeedCost") = Str(dPoundsFed * dFeedCostPerPound)
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLotContentDetails.CalculateFeedStats"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalculateCattleInStats
'' Description: Calculate the cattle in statistics for the given row
'' Inputs:      Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalculateCattleInStats(ByVal lRow As Long)
On Error GoTo ErrSection:

    Dim dHeadIn As Double               ' Head in
    Dim dTotalPayWeight As Double       ' Total pay weight
    Dim dTotalCostOfCattle As Double    ' Total cost of cattle
    
    dHeadIn = ValOfText(ValueForKey(lRow, "HeadIn"))
    dTotalPayWeight = ValOfText(ValueForKey(lRow, "TotalPayWeight"))
    dTotalCostOfCattle = ValOfText(ValueForKey(lRow, "TotalCostOfCattle"))
    
    If dHeadIn <> 0 Then
        ValueForKey(lRow, "AveragePayWeight") = Str(dTotalPayWeight / dHeadIn)
    End If
    If dTotalPayWeight <> 0 Then
        ValueForKey(lRow, "PriceInPerCwt") = Str(dTotalCostOfCattle / (dTotalPayWeight / 100))
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLotContentDetails.CalculateCattleInStats"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyRation
'' Description: Apply the given ration
'' Inputs:      Ration
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyRation(ByVal Ration As cBrokerMessage)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Current state of the grid's redraw
    Dim Ingredients As cGdTree          ' Collection of ingredients
    Dim Ingredient As cBrokerMessage    ' Ingredient object
    Dim astrIngredientID As cGdArray    ' Array of ingredient IDs
    Dim astrPoundsFed As cGdArray       ' Array of pounds fed
    Dim astrPercentMarkup As cGdArray   ' Array of percent markup
    Dim lIndex As Long                  ' Index into a for loop
    Dim dValue As Double                ' Value
    Dim lNewRow As Long                 ' New row in the grid

    If Not Ration Is Nothing Then
        Set Ingredients = g.Cattle.Ingredients
        
        Set astrIngredientID = New cGdArray
        astrIngredientID.SplitFields Ration("IngredientID"), "|"
        Set astrPoundsFed = New cGdArray
        astrPoundsFed.SplitFields Ration("PoundsFed"), "|"
        Set astrPercentMarkup = New cGdArray
        astrPercentMarkup.SplitFields Ration("PercentMarkup"), "|"
        
        With fgDetails
            nRedraw = .Redraw
            .Redraw = flexRDNone
            
            For lIndex = 0 To astrIngredientID.Size - 1
                If Ingredients.Exists(astrIngredientID(lIndex)) Then
                    Set Ingredient = Ingredients(astrIngredientID(lIndex))
                    
                    dValue = Val(Ingredient("CostPerPound"))
                    dValue = dValue + ((Val(astrPercentMarkup(lIndex)) / 100) * dValue)
                    
                    lNewRow = AddRow
                    ValueForKey(lNewRow, "Ingredient") = Ingredient("Ingredient")
                    ValueForKey(lNewRow, "PoundsFed") = astrPoundsFed(lIndex)
                    ValueForKey(lNewRow, "DryFeedPct") = Ingredient("DryFeedPct")
                    ValueForKey(lNewRow, "FeedCostPerPound") = Str(dValue)
                    CalculateFeedStats lNewRow
                End If
            Next lIndex
            
            CalcTotals
            
            .Redraw = nRedraw
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLotContentDetails.ApplyRation"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GotoField
'' Description: Go to the tab and row for the specified key-value field
'' Inputs:      Key Value Field
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GotoField(ByVal strKeyValueField As String, ByVal bDoEdit As Boolean)
On Error GoTo ErrSection:

    Dim lCol As Long                    ' Column in the grid for the given key value field
    Dim lLastNonClickHereLine As Long   ' Last non click-here line in the grid
    
    lCol = ColForKeyValueField(strKeyValueField)
    If lCol <> -1& Then
        m.bSkipEditInRowColChange = True
        lLastNonClickHereLine = LastNonClickHereLine
        If lLastNonClickHereLine = -1& Then
            fgDetails.Row = AddRow
        Else
            fgDetails.Row = lLastNonClickHereLine
        End If
        
        fgDetails.Col = lCol
        m.bSkipEditInRowColChange = False
        
        If bDoEdit Then
DebugLog "GotoField ( " & Str(lCol) & " ) -> EditCell"
            'EditCell fgDetails
            tmrEditCell.Enabled = True
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmEditLotContentDetails.GotoField"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClickHereLine
'' Description: Row of the "Click here" line
'' Inputs:      None
'' Returns:     Row ( -1 if not found )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ClickHereLine() As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    lReturn = -1&
    With fgDetails
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
    RaiseError "frmEditLotContentDetails.ClickHereLine"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClickHereRationLine
'' Description: Row of the "Click here...ration" line
'' Inputs:      None
'' Returns:     Row ( -1 if not found )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ClickHereRationLine() As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    lReturn = -1&
    With fgDetails
        For lIndex = .FixedRows To .Rows - 1
            If RowIsClickHereRationLine(lIndex) Then
                lReturn = lIndex
                Exit For
            End If
        Next lIndex
    End With
    
    ClickHereRationLine = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmEditLotContentDetails.ClickHereRationLine"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClickHereIngredientLine
'' Description: Row of the "Click here...ingredient" line
'' Inputs:      None
'' Returns:     Row ( -1 if not found )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ClickHereIngredientLine() As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    lReturn = -1&
    With fgDetails
        For lIndex = .FixedRows To .Rows - 1
            If RowIsClickHereIngredientLine(lIndex) Then
                lReturn = lIndex
                Exit For
            End If
        Next lIndex
    End With
    
    ClickHereIngredientLine = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmEditLotContentDetails.ClickHereIngredientLine"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddClickHereLine
'' Description: Add a click here line if it doesn't already exist
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddClickHereLine()
On Error GoTo ErrSection:
    
    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid

    If ClickHereLine = -1& Then
        With fgDetails
            nRedraw = .Redraw
            .Redraw = flexRDNone
            
            .Rows = .Rows + 1
            .MergeRow(.Rows - 1) = True
            .TextMatrix(.Rows - 1, GDCol(eGDCol_FeedYardID)) = "-1"
            .Cell(flexcpText, .Rows - 1, GDCol(eGDCol_Date), .Rows - 1, NotesCol) = "Click here to add a new row"
            .Cell(flexcpForeColor, .Rows - 1, GDCol(eGDCol_Date), .Rows - 1, NotesCol) = vbBlue
            .Cell(flexcpFontUnderline, .Rows - 1, GDCol(eGDCol_Date), .Rows - 1, NotesCol) = True
            .Cell(flexcpAlignment, .Rows - 1, GDCol(eGDCol_Date), .Rows - 1, NotesCol) = flexAlignLeftCenter
            
            .Redraw = nRedraw
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLotContentDetails.AddClickHereLine"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddClickHereRationLine
'' Description: Add a click here line for ration if it doesn't already exist
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddClickHereRationLine()
On Error GoTo ErrSection:
    
    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid

    If (ClickHereRationLine = -1&) And (m.bHasIngredient = True) Then
        With fgDetails
            nRedraw = .Redraw
            .Redraw = flexRDNone
            
            .Rows = .Rows + 1
            .MergeRow(.Rows - 1) = True
            .TextMatrix(.Rows - 1, GDCol(eGDCol_FeedYardID)) = "-2"
            .Cell(flexcpText, .Rows - 1, GDCol(eGDCol_Date), .Rows - 1, NotesCol) = "Click here to insert a ration"
            .Cell(flexcpForeColor, .Rows - 1, GDCol(eGDCol_Date), .Rows - 1, NotesCol) = vbBlue
            .Cell(flexcpFontUnderline, .Rows - 1, GDCol(eGDCol_Date), .Rows - 1, NotesCol) = True
            .Cell(flexcpAlignment, .Rows - 1, GDCol(eGDCol_Date), .Rows - 1, NotesCol) = flexAlignLeftCenter
            
            .Redraw = nRedraw
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLotContentDetails.AddClickHereRationLine"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddClickHereIngredientLine
'' Description: Add a click here line for ingredient if it doesn't already exist
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddClickHereIngredientLine()
On Error GoTo ErrSection:
    
    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid

    If (ClickHereIngredientLine = -1&) And (m.bHasIngredient = True) Then
        With fgDetails
            nRedraw = .Redraw
            .Redraw = flexRDNone
            
            .Rows = .Rows + 1
            .MergeRow(.Rows - 1) = True
            .TextMatrix(.Rows - 1, GDCol(eGDCol_FeedYardID)) = "-3"
            .Cell(flexcpText, .Rows - 1, GDCol(eGDCol_Date), .Rows - 1, NotesCol) = "Click here to manage ingredients"
            .Cell(flexcpForeColor, .Rows - 1, GDCol(eGDCol_Date), .Rows - 1, NotesCol) = vbBlue
            .Cell(flexcpFontUnderline, .Rows - 1, GDCol(eGDCol_Date), .Rows - 1, NotesCol) = True
            .Cell(flexcpAlignment, .Rows - 1, GDCol(eGDCol_Date), .Rows - 1, NotesCol) = flexAlignLeftCenter
            
            .Redraw = nRedraw
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLotContentDetails.AddClickHereIngredientLine"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RowIsClickHereLine
'' Description: Determine if the given row is a click here line
'' Inputs:      Row
'' Returns:     True if Click Here line, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RowIsClickHereLine(ByVal lRow As Long) As Boolean
On Error GoTo ErrSection:

    RowIsClickHereLine = (fgDetails.TextMatrix(lRow, GDCol(eGDCol_FeedYardID)) = "-1")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmEditLotContentDetails.RowIsClickHereLine"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RowIsClickHereRationLine
'' Description: Determine if the given row is a "click here...ration" line
'' Inputs:      Row
'' Returns:     True if Click Here Ration line, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RowIsClickHereRationLine(ByVal lRow As Long) As Boolean
On Error GoTo ErrSection:

    RowIsClickHereRationLine = (fgDetails.TextMatrix(lRow, GDCol(eGDCol_FeedYardID)) = "-2")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmEditLotContentDetails.RowIsClickHereRationLine"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RowIsClickHereIngredientLine
'' Description: Determine if the given row is a "click here...ingredient" line
'' Inputs:      Row
'' Returns:     True if Click Here Ingredient line, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RowIsClickHereIngredientLine(ByVal lRow As Long) As Boolean
On Error GoTo ErrSection:

    RowIsClickHereIngredientLine = (fgDetails.TextMatrix(lRow, GDCol(eGDCol_FeedYardID)) = "-3")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmEditLotContentDetails.RowIsClickHereIngredientLine"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InsertRation
'' Description: Insert a ration
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InsertRation()
On Error GoTo ErrSection:

    ApplyRation frmCattleManage.ShowMeRations(True)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLotContentDetails.InsertRation"
    
End Sub

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
    With fgDetails
        For lIndex = .Rows - 1 To .FixedRows Step -1&
            If (RowIsClickHereLine(lIndex) = False) And (RowIsClickHereRationLine(lIndex) = False) And (RowIsClickHereIngredientLine(lIndex) = False) Then
                lReturn = lIndex
                Exit For
            End If
        Next lIndex
    End With

    LastNonClickHereLine = lReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmEditLotContentDetails.LastNonClickHereLine"
    
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
        Row = fgDetails.Row
    End If
    
    lLastValidRow = LastNonClickHereLine
    
    ValidDataRow = ((ValidGridRow(fgDetails, Row) = True) And (Row <= lLastValidRow))

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmEditLotContentDetails.ValidDataRow"
    
End Function

