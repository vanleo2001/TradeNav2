VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmEditLot 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   1215
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
      Caption         =   "frmEditLot.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmEditLot.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmEditLot.frx":004C
      RightToLeft     =   0   'False
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
         Caption         =   "frmEditLot.frx":0068
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmEditLot.frx":0096
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmEditLot.frx":00B6
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
         Caption         =   "frmEditLot.frx":00D2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmEditLot.frx":00F8
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmEditLot.frx":0118
         RightToLeft     =   0   'False
      End
   End
   Begin vsOcx6LibCtl.vsIndexTab tabLotInfo 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   5106
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   1
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   600
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FrontTabColor   =   -2147483633
      BackTabColor    =   -2147483633
      TabOutlineColor =   0
      FrontTabForeColor=   -2147483630
      Caption         =   "Tab&1"
      Align           =   0
      Appearance      =   1
      CurrTab         =   0
      FirstTab        =   0
      Style           =   1
      Position        =   1
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   -1  'True
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   0
      BorderWidth     =   0
      BoldCurrent     =   -1  'True
      DogEars         =   -1  'True
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      Begin HexUniControls.ctlUniFrameWL fraTab 
         Height          =   2520
         Index           =   0
         Left            =   45
         TabIndex        =   4
         Top             =   45
         Width           =   2805
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
         Caption         =   "frmEditLot.frx":0134
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmEditLot.frx":0154
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmEditLot.frx":0174
         RightToLeft     =   0   'False
         Begin VSFlex7LCtl.VSFlexGrid fgTab 
            Height          =   1575
            Index           =   0
            Left            =   480
            TabIndex        =   5
            Top             =   420
            Width           =   1875
            _cx             =   3307
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
      End
   End
   Begin VB.Image imgEdit 
      Height          =   240
      Left            =   3780
      Picture         =   "frmEditLot.frx":0190
      Stretch         =   -1  'True
      Top             =   2580
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgCalendar2 
      Height          =   240
      Left            =   4080
      Picture         =   "frmEditLot.frx":0513
      Top             =   2580
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgButton 
      Height          =   200
      Left            =   4140
      Picture         =   "frmEditLot.frx":0891
      Stretch         =   -1  'True
      Top             =   1860
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.Image imgCalendar 
      Height          =   240
      Left            =   3780
      Picture         =   "frmEditLot.frx":0993
      Stretch         =   -1  'True
      Top             =   2220
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgPencil 
      Height          =   225
      Left            =   4140
      Picture         =   "frmEditLot.frx":0D63
      Stretch         =   -1  'True
      Top             =   2220
      Visible         =   0   'False
      Width           =   225
   End
End
Attribute VB_Name = "frmEditLot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmEditLot.frm
'' Description: Form for allowing user to edit a feedyard lot
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 11/18/2013   DAJ         Added an ordinal value for lot columns for sorting
'' 11/19/2013   DAJ         Make sure hidden rows stay hidden after expanding; Set back colors
'' 11/26/2013   DAJ         Tweaks to Turnkey detail editing
'' 12/05/2013   DAJ         Aggregate Column Modes; Fix for groups with multiple text fields
'' 12/05/2013   DAJ         Aggregate Column Mode tweaks
'' 12/19/2013   DAJ         "Lauren List" tweaks
'' 01/14/2014   DAJ         Cattle Navigator calculations
'' 01/23/2014   DAJ         Multiple owners per lot
'' 02/07/2014   DAJ         Icons for rows that can be edited
'' 02/10/2014   DAJ         Changed icons for the rows that can be edited
'' 02/19/2014   DAJ         Check for valid row and column in HandleEdit
'' 02/25/2014   DAJ         Don't bring up details form on a right-click
'' 02/26/2014   DAJ         Double click in grid brings up lot editor on correct field
'' 02/28/2014   DAJ         Pass LotColumn collection into cTurnkeyStats
'' 03/07/2014   DAJ         Moved into NavCattle.DLL
'' 03/14/2014   DAJ         Added support for Boolean lot column type
'' 03/19/2014   DAJ         Pass default startup field to the lot content details form
'' 03/21/2014   DAJ         Added required & actual contracts and percent hedged
'' 04/08/2014   DAJ         Added Average Pay Weight and Cattle Cost per CWT
'' 04/15/2014   DAJ         Use new owner lookup form; Fixes for boolean columns
'' 05/22/2014   DAJ         Renamed cTurnkeyStats to cCattleStats; Renamed frmTurnkeyEditLot to frmEditLot;
''                          Renamed frmTurnkeyLotContentDetails to frmEditLotContentDetails
'' 05/22/2014   DAJ         Renamed frmTurnkey to frmLots; Renamed g.Turnkey to g.Cattle
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Enum eGDCols
    eGDCol_Name = 0
    eGDCol_DisplayValue = 1
    eGDCol_Format = 2
    eGDCol_ActualValue = 3
    eGDCol_KeyValueField = 4
    eGDCol_ToolTipText = 5
    eGDCol_Sort = 6
    eGDCol_OutlineLevel = 7
    
    eGDCol_NumCols
End Enum

Private Type mPrivate
    bOK As Boolean                      ' Did the user click OK?
    
    lFeedYardID As Long                 ' Feed Yard ID
    Lot As cBrokerMessage               ' Lot passed in
    LotDetails As cGdTree               ' Collection of lot details
    iButton As Integer                  ' Mouse button pressed
    strDefaultKeyValueField As String   ' Default key value field
    bAlreadyDone As Boolean             ' Have we done one time stuff?
    
    CalcStats As cCattleStats           ' Object to calculate statistics
    
    Categories As cGdTree               ' Collection of categories
    Subcategories As cGdTree            ' Collection of subcategories
    Owners As cGdTree                   ' Collection of owners
    astrOwners As cGdArray              ' Array of owners to pass to dialog
    Ordinal As cGdTree                  ' Collection of ordinal values
    
    LotColumnMap As cGdTree             ' Map of where the lot columns are
    
    lOwnerNumberRow As Long             ' Owner number row
    lOwnerNumberTab As Long             ' Owner number tab
    lOwnerNameRow As Long               ' Owner name row
    lOwnerNameTab As Long               ' Owner name tab
End Type
Private m As mPrivate

Private Property Get GDCol(ByVal nCol As eGDCols) As Long
    GDCol = nCol
End Property

Private Property Get Value(ByVal strKeyValueField As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim lTab As Long                    ' Tab number for the field
    Dim lRow As Long                    ' Row number for the field

    strReturn = ""
    If m.LotColumnMap.Exists(strKeyValueField) Then
        lTab = CLng(Val(Parse(m.LotColumnMap(strKeyValueField), ";", 1)))
        lRow = CLng(Val(Parse(m.LotColumnMap(strKeyValueField), ";", 2)))
        
        strReturn = fgTab(lTab).TextMatrix(lRow, GDCol(eGDCol_ActualValue))
    End If
    
    Value = strReturn

ErrExit:
    Exit Property

ErrSection:
    RaiseError "frmEditLot.Value.Get"
    
End Property
Private Property Let Value(ByVal strKeyValueField As String, ByVal strValue As String)
On Error GoTo ErrSection:

    Dim lTab As Long                    ' Tab number for the field
    Dim lRow As Long                    ' Row number for the field
    Dim LotColumn As cLotColumn         ' Lot column object

    If m.LotColumnMap.Exists(strKeyValueField) Then
        lTab = CLng(Val(Parse(m.LotColumnMap(strKeyValueField), ";", 1)))
        lRow = CLng(Val(Parse(m.LotColumnMap(strKeyValueField), ";", 2)))
        
        Set LotColumn = fgTab(lTab).RowData(lRow)
        SetValue lTab, lRow, strValue, LotColumn
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmEditLot.Value.Let"
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Setup and show the form
'' Inputs:      Feed Yard ID, Lot, Lot Details
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ByVal lFeedYardID As Long, Lot As cBrokerMessage, LotDetails As cGdTree, Optional ByVal strDefaultKeyValueField As String = "") As Boolean
On Error GoTo ErrSection:

    Dim Owners As cGdTree               ' Collection of owners
    Dim lIndex As Long                  ' Index into a for loop

    m.lFeedYardID = lFeedYardID
    m.strDefaultKeyValueField = strDefaultKeyValueField
    Set m.Lot = Lot
    Set m.LotDetails = LotDetails

    Set m.Categories = g.Cattle.LotColumnCategories
    Set m.Subcategories = g.Cattle.LotColumnSubCategories
    Set Owners = g.Cattle.Customers
    Set m.Owners = New cGdTree
    Set m.astrOwners = New cGdArray
    m.astrOwners.Create eGDARRAY_Strings
    Set m.Ordinal = New cGdTree
    
    For lIndex = 1 To Owners.Count
        AddOwner Owners(lIndex)
    Next lIndex

    If InitTabs > 0 Then
        SetEditorCaption Me, "Lot", Lot("Number")
        
        LoadGrids
        
        If Len(Value("RequiredContracts")) = 0 Then
            CalculateHedgingColumns
        End If
        
        ShowForm Me, eForm_Modal, g.frmMain
        
        If m.bOK Then
            Save
            Set Lot = m.Lot
            Set LotDetails = m.LotDetails
        End If
    Else
        InfBox "No lot column categories", "!", , "Error"
        m.bOK = False
    End If
    
    ShowMe = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmEditLot.ShowMe"
    
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
    RaiseError "frmEditLot.Turnkey_Customer"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Unload the form without saving the information
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
    RaiseError "frmEditLot.cmdCancel_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: Save the information and exit the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    If VerifyData Then
        m.bOK = True
        Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLot.cmdOK_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgTab_AfterCollapse
'' Description: Handle the user expanding/collapsing nodes
'' Inputs:      Index of Grid, Row, State
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgTab_AfterCollapse(Index As Integer, ByVal Row As Long, ByVal State As Integer)
On Error GoTo ErrSection:

    Dim lChildRow As Long               ' Row in the grid
    Dim LotColumn As cLotColumn         ' Lot column object

    If State = flexOutlineExpanded Then
        With fgTab(Index)
            lChildRow = .GetNodeRow(Row, flexNTFirstChild)
            Do While lChildRow <> -1&
                If TypeOf .RowData(lChildRow) Is cLotColumn Then
                    Set LotColumn = .RowData(lChildRow)
                
                    .RowHidden(lChildRow) = LotColumn.AlwaysHidden Or LotColumn.FeedyardHidden
                End If
                
                lChildRow = .GetNodeRow(lChildRow, flexNTNextSibling)
            Loop
        End With
    End If

    SetBackColors fgTab(Index)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLot.fgTab_AfterCollapse"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgTab_AfterEdit
'' Description: After the user is done editing the cell, format the value
'' Inputs:      Index of Grid, Row, Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgTab_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim strActualValue As String        ' Actual value for the displayed value
    Dim LotColumn As cLotColumn         ' Lot column object
    Dim strOwnerNumber As String        ' Owner number

    With fgTab(Index)
        If (Row >= .FixedRows) And (Row < .Rows) Then
            Set LotColumn = .RowData(Row)
            
            If UCase(LotColumn.Format) <> "DATE" Then
                strActualValue = g.Cattle.GridValue(fgTab(Index), Row, Col, LotColumn)
                SetValue Index, Row, strActualValue, LotColumn
            End If
            
            If (LotColumn.KeyValueField = "ActualContracts") Or (LotColumn.KeyValueField = "ProjectedWeightOut") Then
                CalculateHedgingColumns
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTunkeyEditLot.fgTab_AfterEdit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgTab_AfterRowColChange
'' Description: Handle the user changing cells in the grid
'' Inputs:      Index of Grid, Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgTab_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    Dim LotColumn As cLotColumn         ' Lot column object

    With fgTab(Index)
        If NewCol = 1 Then
            If (NewRow >= .FixedRows) And (NewRow < .Rows) Then
                If .RowOutlineLevel(NewRow) > 0 Then
                    Set LotColumn = .RowData(NewRow)
                    
                    If (UCase(LotColumn.KeyValueField) <> "OWNERNAME") And (UCase(LotColumn.KeyValueField) <> "OWNERNUMBER") Then
                        If (UCase(LotColumn.Format) <> "DATE") And (LotColumn.IsAggregate = 0) And (UCase(LotColumn.KeyValueField) = "OWNERNAME") Then
                            fgTab(Index).EditCell
                        End If
                    End If
                End If
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLot.fgTab_AfterRowColChange"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgTab_BeforeEdit
'' Description: Make sure the user can only edit appropriate cells
'' Inputs:      Index of Grid, Row, Column, Cancel Edit?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgTab_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim Owner As cBrokerMessage         ' Owner object
    Dim astrOwners As cGdArray          ' List of owners
    Dim lIndex As Long                  ' Index into a for loop
    Dim LotColumn As cLotColumn         ' Lot column object
    Dim bCancel As Boolean              ' Cancel the edit?
    
    bCancel = (Col <> GDCol(eGDCol_DisplayValue))
    
    If (Col = GDCol(eGDCol_DisplayValue)) Then
        With fgTab(Index)
            If TypeOf .RowData(Row) Is cLotColumn Then
                Set LotColumn = .RowData(Row)
                
                If UCase(LotColumn.Format) = "DATE" Then
                    '.ColComboList(Col) = "..."
                    bCancel = True
                ElseIf LotColumn.IsAggregate > 0 Then
                    '.ColComboList(Col) = "..."
                    bCancel = True
                ElseIf UCase(LotColumn.KeyValueField) = "OWNERNUMBER" Then
                    '.ColComboList(Col) = "..."
                    bCancel = True
                ElseIf UCase(LotColumn.KeyValueField) = "OWNERNAME" Then
                    '.ColComboList(Col) = "..."
                    bCancel = True
                ElseIf UCase(LotColumn.Format) = "BOOLEAN" Then
                    bCancel = True
                Else
                    .ColComboList(Col) = ""
                End If
            Else
                bCancel = True
            End If
        End With
    End If
    
    Cancel = bCancel

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLot.fgTab_BeforeEdit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgTab_BeforeMouseDown
'' Description: Handle the user starting to click a mouse button
'' Inputs:      Index of Grid, Button pressed, Shift/Ctrl/Alt status, Mouse
''              Location, Cancel the Click?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgTab_BeforeMouseDown(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    m.iButton = Button

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "fgTab_BeforeMouseDown"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgTab_CellButtonClick
'' Description: Handle the user clicking on the "..." button in the cell
'' Inputs:      Index of Grid, Row, Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgTab_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim pt As POINTAPI                  ' Mouse location point
    Dim dDate As Double                 ' Date to send to the edit date form
    Dim LotColumn As cLotColumn         ' Lot column object
    Dim lFeedYardLotID As Long          ' Feed Yard Lot ID
    Dim Details As cGdTree              ' Lot details
    Dim Detail As cBrokerMessage        ' Detail object
    Dim lIndex As Long                  ' Index into a for loop
    Dim bCancelled As Boolean           ' Was the dialog cancelled?
    Dim strOwner As String              ' Owner selected
    Dim dTotal As Double                ' Total of the values in the details

    With fgTab(Index)
        pt.X = .ColPos(Col) / Screen.TwipsPerPixelX
        'pt.Y = (.RowPos(Row) - ((frmEditDate.Height - .RowHeight(Row)) / 2)) / Screen.TwipsPerPixelY
        pt.Y = (.RowPos(Row) + .RowHeight(Row)) / Screen.TwipsPerPixelY
        ClientToScreen .hWnd, pt
        
        Set LotColumn = .RowData(Row)
        m.strDefaultKeyValueField = LotColumn.KeyValueField
        
        If LotColumn.AggregateIsGroup Then
            EditDetailsForCategory Index, Row, Col, False
            
        ElseIf LotColumn.AggregateIsSingle Then
            EditDetails Index, Row, Col
            
        ElseIf LotColumn.AggregateIsOwner Then
            EditDetailsForCategory Index, Row, Col, True
            
        ElseIf UCase(LotColumn.Format) = "DATE" Then
            pt.X = pt.X * Screen.TwipsPerPixelX
            pt.Y = pt.Y * Screen.TwipsPerPixelY
            dDate = Val(.TextMatrix(Row, GDCol(eGDCol_ActualValue)))
            If dDate = 0 Then dDate = Date
            
            frmEditDate.BackColor = BackColor
            dDate = frmEditDate.ShowMe(pt.X, pt.Y, dDate, Me, , , , , , , , , , , bCancelled)
            If bCancelled = False Then
                SetValue Index, Row, Str(dDate), LotColumn
            End If
        ElseIf UCase(LotColumn.KeyValueField) = "OWNERNUMBER" Then
            'strOwner = g.AppBridge.AccountLookup(m.astrOwners, .TextMatrix(Row, Col), , True)
            strOwner = frmOwnerLookup.ShowMe(m.astrOwners, .TextMatrix(Row, Col), , True)
            If Len(strOwner) > 0 Then
                SetOwner strOwner
            End If
        ElseIf UCase(LotColumn.KeyValueField) = "OWNERNAME" Then
            'strOwner = g.AppBridge.AccountLookup(m.astrOwners, , .TextMatrix(Row, Col), True)
            strOwner = frmOwnerLookup.ShowMe(m.astrOwners, , .TextMatrix(Row, Col), True)
            If Len(strOwner) > 0 Then
                SetOwner strOwner
            End If
        End If
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLot.fgTab_CellButtonClick"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgTab_Click
'' Description: Handle a user click in the grid
'' Inputs:      Index of the Grid
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgTab_Click(Index As Integer)
On Error GoTo ErrSection:

    If m.iButton = vbLeftButton Then
        HandleEdit Index, fgTab(Index).MouseRow, fgTab(Index).MouseCol
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLot.fgTab_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgTab_Compare
'' Description: Perform a comparison for the two rows for sorting purposes
'' Inputs:      Index of the Grid, Row 1, Row 2, Compare Value
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgTab_Compare(Index As Integer, ByVal Row1 As Long, ByVal Row2 As Long, Cmp As Integer)
On Error GoTo ErrSection:

    Dim strRow1 As String
    Dim strRow2 As String
    
    strRow1 = fgTab(Index).TextMatrix(Row1, GDCol(eGDCol_Sort))
    strRow2 = fgTab(Index).TextMatrix(Row2, GDCol(eGDCol_Sort))
    
    If strRow1 = strRow2 Then
        Cmp = 0
    ElseIf strRow1 < strRow2 Then
        Cmp = -1
    Else
        Cmp = 1
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLot.fgTab_Compare"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgTab_DblClick
'' Description: Handle a user double click in the grid
'' Inputs:      Index of the Grid
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgTab_DblClick(Index As Integer)
On Error GoTo ErrSection:

    If m.iButton = vbLeftButton Then
        HandleEdit Index, fgTab(Index).MouseRow, fgTab(Index).MouseCol
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLot.fgTab_DblClick"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgTab_KeyPress
'' Description: Handle the user pressing a key in the grid
'' Inputs:      Index of the Grid, Key Pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgTab_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo ErrSection:

    If (IsDialogRow(Index, fgTab(Index).Row) = True) And (fgTab(Index).Col = GDCol(eGDCol_DisplayValue)) Then
        If (KeyAscii = vbKeyReturn) Or (KeyAscii = vbKeySpace) Then
            HandleEdit Index, fgTab(Index).Row, fgTab(Index).Col
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLot.fgTab_KeyPress"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgTab_MouseMove
'' Description: Handle the user moving the mouse over the grid
'' Inputs:      Index of the Grid, Mouse Button Pressed, Shift/Ctrl/Alt Status,
''              X Location, Y Location
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgTab_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    Static strLastToolTip As String     ' Last Tool tip text
    Dim strTooltip As String            ' Tool tip text
    Dim lMouseRow As Long               ' Mouse row in the grid
    Dim lMouseCol As Long               ' Mouse column in the grid
    
    With fgTab(Index)
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        strTooltip = ""
        If ((lMouseRow >= .FixedRows) And (lMouseRow < .Rows)) Then
            If .RowOutlineLevel(lMouseRow) > 0 Then
                If lMouseCol = GDCol(eGDCol_Name) Then
                    strTooltip = .TextMatrix(lMouseRow, GDCol(eGDCol_ToolTipText))
                End If
            End If
        End If
    
        If strTooltip <> strLastToolTip Then
            .TooltipText = strTooltip
            strLastToolTip = strTooltip
        End If
    End With

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgTab_ValidateEdit
'' Description: Validate an edit of a cell
'' Inputs:      Index of the Grid, Row, Column, Cancel Edit?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgTab_ValidateEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTunkeyEditLot.fgTab_ValidateEdit"
    
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

    If m.bAlreadyDone = False Then
        m.bAlreadyDone = True
        If Len(m.strDefaultKeyValueField) > 0 Then
            GotoField m.strDefaultKeyValueField, True
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmEditLot.Form_Activate"
    
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

    Icon = g.AppBridge.Picture16(g.Cattle.IconName)
    
    g.Styler.StyleForm Me
    
    PlaceForm Me
    
    Set m.CalcStats = New cCattleStats
    Set m.LotColumnMap = New cGdTree
    m.bAlreadyDone = False
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLot.Form_Load"
    
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
    RaiseError "frmEditLot.Form_QueryUnload"
    
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

    Dim lSpace As Long                  ' Space between controls
    Dim lMinScaleWidth As Long          ' Minimum scale width
    Dim lMinScaleHeight As Long         ' Minimum scale height
    Dim lIndex As Long                  ' Index into a for loop
    
    lSpace = 60
    lMinScaleWidth = (fraButtons.Width * 3) + (lSpace * 3)
    lMinScaleHeight = (fraButtons.Height * 3) + (lSpace * 2)
    
    If LimitFormSize(Me, lMinScaleWidth, lMinScaleHeight) = False Then
        With tabLotInfo
            .Move lSpace, lSpace, ScaleWidth - fraButtons.Width - (lSpace * 3), ScaleHeight - (lSpace * 2)
        End With
        With fraButtons
            .Move ScaleWidth - fraButtons.Width - lSpace, lSpace
        End With
        
        For lIndex = 0 To tabLotInfo.NumTabs - 1
            With fraTab(lIndex)
                .Move 0, 0, tabLotInfo.ClientWidth, tabLotInfo.ClientHeight
            End With
            With fgTab(lIndex)
                .Move 0, 0, tabLotInfo.ClientWidth, tabLotInfo.ClientHeight
            End With
        Next lIndex
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Clean things up when the form is unloaded
'' Inputs:      Cancel the Unload?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    SaveFormPlacement Me
    
    For lIndex = 1 To tabLotInfo.NumTabs - 1
        Unload fgTab(lIndex)
        Unload fraTab(lIndex)
    Next lIndex
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLot.Form_Unload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tabLotInfo_Switch
'' Description: Handle the user switching tabs
'' Inputs:      Old Tab, New Tab, Cancel the Switch
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tabLotInfo_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
On Error GoTo ErrSection:

    fraTab(NewTab).Visible = True
    fraTab(OldTab).Visible = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLot.tabLotInfo_Switch"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitTabs
'' Description: Initialize the tab control
'' Inputs:      None
'' Returns:     Number of tabs/categories
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function InitTabs() As Long
On Error GoTo ErrSection:

    Dim lNumTabs As Long                ' Number of tabs
    Dim lIndex As Long                  ' Index into a for loop
    Dim astrCaption As cGdArray         ' Tab caption
    Dim turnkeyMsg As cBrokerMessage    ' Turnkey message
    Dim lTabNumber As Long              ' Tab number
    
    lNumTabs = m.Categories.Count
    
    If lNumTabs > 0 Then
        Set astrCaption = New cGdArray
        astrCaption.Create eGDARRAY_Strings, lNumTabs
        For lIndex = 0 To lNumTabs - 1
            Set turnkeyMsg = m.Categories(lIndex + 1)
    
            lTabNumber = CLng(Val(turnkeyMsg("TabNumber")))
            astrCaption(lTabNumber - 1) = turnkeyMsg("CategoryName")
        Next lIndex
        
        tabLotInfo.Caption = astrCaption.JoinFields("|")
        
        'RH fraTab(0).BorderStyle = 0
        InitGrid 0
        
        For lIndex = 1 To lNumTabs - 1
            Load fraTab(lIndex)
            SetParent fraTab(lIndex).hWnd, tabLotInfo.hWnd
            'RH fraTab(lIndex).BorderStyle = 0
            fraTab(lIndex).Visible = False
            
            Load fgTab(lIndex)
            SetParent fgTab(lIndex).hWnd, fraTab(lIndex).hWnd
            InitGrid lIndex
            fgTab(lIndex).Visible = True
        Next lIndex
    End If
    
    InitTabs = lNumTabs

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmEditLot.InitTabs"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitGrid
'' Description: Initialize the appropriate grid
'' Inputs:      Index
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitGrid(ByVal lIndex As Long)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    
    With fgTab(lIndex)
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExNone
        .ExtendLastCol = True
        .MergeCells = flexMergeFree
        .OutlineBar = flexOutlineBarSimpleLeaf
        .ScrollTrack = True
        .SelectionMode = flexSelectionFree
        .SheetBorder = RGB(128, 128, 128)
        .TabBehavior = flexTabCells
        .WordWrap = False
        
        .FixedRows = 0
        .Rows = 0
        .FixedCols = 0
        .Cols = GDCol(eGDCol_NumCols)
        
        .ColAlignment(GDCol(eGDCol_DisplayValue)) = flexAlignLeftCenter
        
        .ColHidden(GDCol(eGDCol_Format)) = True
        .ColHidden(GDCol(eGDCol_ActualValue)) = True
        .ColHidden(GDCol(eGDCol_KeyValueField)) = True
        .ColHidden(GDCol(eGDCol_ToolTipText)) = True
        .ColHidden(GDCol(eGDCol_Sort)) = True
        .ColHidden(GDCol(eGDCol_OutlineLevel)) = True
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLot.InitGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGrids
'' Description: Load the grids
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGrids()
On Error GoTo ErrSection:

    Dim alRedraw As cGdArray            ' Array of redraw settings
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim LotColumns As cGdTree           ' Collection of lot columns
    Dim LotColumn As cLotColumn         ' Lot column object
    Dim lTabNumber As Long              ' Tab number for the category
    Dim strActualValue As String        ' Actual value
    Dim turnkeyMessage As cBrokerMessage ' Turnkey message
    Dim lParentRow As Long              ' Parent row
    Dim lLastChild As Long              ' Last child for the parent
    
    Set alRedraw = New cGdArray
    alRedraw.Create eGDARRAY_Longs, tabLotInfo.NumTabs
    For lIndex = 0 To tabLotInfo.NumTabs - 1
        alRedraw(lIndex) = fgTab(lIndex).Redraw
        fgTab(lIndex).Redraw = flexRDNone
    Next lIndex
    
    m.LotColumnMap.Clear
    
    For lIndex = 1 To m.Subcategories.Count
        Set turnkeyMessage = m.Subcategories(lIndex)
        lTabNumber = TabNumberForCategoryID(CLng(Val(turnkeyMessage("LotColumnCategoryID"))))
        If lTabNumber > -1& Then
            With fgTab(lTabNumber)
                .Rows = .Rows + 1
                
                If m.Ordinal.Exists(turnkeyMessage("ID")) = False Then
                    m.Ordinal.Add turnkeyMessage("Ordinal"), turnkeyMessage("ID")
                End If
                
                .Cell(flexcpText, .Rows - 1, GDCol(eGDCol_Name), .Rows - 1, GDCol(eGDCol_DisplayValue)) = turnkeyMessage("SubCategoryName")
                
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Sort)) = SortKey(turnkeyMessage("ID"), "", 0&)
                .TextMatrix(.Rows - 1, GDCol(eGDCol_OutlineLevel)) = 0
                
                .MergeRow(.Rows - 1) = True
                .IsSubtotal(.Rows - 1) = False
                .RowOutlineLevel(.Rows - 1) = 0
            End With
        End If
    Next lIndex
    
    Set LotColumns = frmLots.LotColumns
    For lIndex = 1 To LotColumns.Count
        Set LotColumn = LotColumns(lIndex)
        
        lTabNumber = TabNumberForCategoryID(LotColumn.CategoryID)
        If lTabNumber > -1& Then
            With fgTab(lTabNumber)
                .Rows = .Rows + 1
                
                .RowData(.Rows - 1) = LotColumn
                
                strActualValue = m.Lot(LotColumn.KeyValueField)
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Name)) = LotColumn.ColumnHeader
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Format)) = LotColumn.Format
                .TextMatrix(.Rows - 1, GDCol(eGDCol_KeyValueField)) = LotColumn.KeyValueField
                .TextMatrix(.Rows - 1, GDCol(eGDCol_ToolTipText)) = LotColumn.TooltipText
                
                SetValue lTabNumber, .Rows - 1, strActualValue, LotColumn
                
                If UCase(LotColumn.Format) = "DATE" Then
                    .Cell(flexcpPicture, .Rows - 1, GDCol(eGDCol_DisplayValue)) = imgCalendar2.Picture
                ElseIf LotColumn.IsAggregate > 0 Then
                    .Cell(flexcpPicture, .Rows - 1, GDCol(eGDCol_DisplayValue)) = imgEdit.Picture
                End If
                
                If UCase(LotColumn.Format) = "BOOLEAN" Then
                    .Cell(flexcpPictureAlignment, .Rows - 1, GDCol(eGDCol_DisplayValue)) = flexAlignCenterCenter
                Else
                    .Cell(flexcpPictureAlignment, .Rows - 1, GDCol(eGDCol_DisplayValue)) = flexAlignRightBottom
                End If
                
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Sort)) = SortKey(Str(LotColumn.SubCategoryID), LotColumn.ColumnHeader, LotColumn.Ordinal)
                .TextMatrix(.Rows - 1, GDCol(eGDCol_OutlineLevel)) = 1
                
                .MergeRow(.Rows - 1) = False
                .IsSubtotal(.Rows - 1) = False
                .RowOutlineLevel(.Rows - 1) = 0
                
                .RowHidden(.Rows - 1) = LotColumn.AlwaysHidden Or LotColumn.FeedyardHidden
            End With
        End If
    Next lIndex

    For lIndex = 0 To tabLotInfo.NumTabs - 1
        With fgTab(lIndex)
            .Select .FixedRows, GDCol(eGDCol_Sort), .Rows - 1, GDCol(eGDCol_Sort)
            .Sort = flexSortCustom
            .Select 0, 0
            
            For lIndex2 = .FixedRows To .Rows - 1
                .IsSubtotal(lIndex2) = True
                .RowOutlineLevel(lIndex2) = .TextMatrix(lIndex2, GDCol(eGDCol_OutlineLevel))
                
                If TypeOf .RowData(lIndex2) Is cLotColumn Then
                    Set LotColumn = .RowData(lIndex2)
                
                    If m.LotColumnMap.Exists(LotColumn.KeyValueField) Then
                        m.LotColumnMap(LotColumn.KeyValueField) = Str(lIndex) & ";" & Str(lIndex2)
                    Else
                        m.LotColumnMap.Add Str(lIndex) & ";" & Str(lIndex2), LotColumn.KeyValueField
                    End If
                    
                    If UCase(LotColumn.KeyValueField) = "OWNERNUMBER" Then
                        m.lOwnerNumberRow = lIndex2
                        m.lOwnerNumberTab = lIndex
                    ElseIf UCase(LotColumn.KeyValueField) = "OWNERNAME" Then
                        m.lOwnerNameRow = lIndex2
                        m.lOwnerNameTab = lIndex
                    End If
                End If
            Next lIndex2
            
            SetBackColors fgTab(lIndex)
            
            .AutoSize 0, .Cols - 1, False
            .Redraw = alRedraw(lIndex)
        End With
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLot.LoadGrids"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TabNumberForCategoryID
'' Description: Get the tab number for the given category ID
'' Inputs:      Category ID
'' Returns:     Tab Number ( -1 if not found )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TabNumberForCategoryID(ByVal lCategoryID As Long) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim Category As cBrokerMessage      ' Category information
    
    lReturn = -1&
    If m.Categories.Exists(Str(lCategoryID)) Then
        Set Category = m.Categories(Str(lCategoryID))
        lReturn = CLng(Val(Category("TabNumber"))) - 1&
    End If
    
    TabNumberForCategoryID = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTurkeyEditLot.TabNumberForCategoryID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Save the information on the form into the Lot object
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Save()
On Error GoTo ErrSection:

    Dim lTab As Long                    ' Index into a for loop
    Dim lIndex As Long                  ' Index into a for loop
    
    For lTab = 0 To tabLotInfo.NumTabs - 1
        With fgTab(lTab)
            For lIndex = .FixedRows To .Rows - 1
                If .RowOutlineLevel(lIndex) = 1 Then
                    m.Lot.Add .TextMatrix(lIndex, GDCol(eGDCol_KeyValueField)), .TextMatrix(lIndex, GDCol(eGDCol_ActualValue))
                End If
            Next lIndex
        End With
    Next lTab

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLot.Save"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveDetails
'' Description: Remove the appropriate details out of the collection
'' Inputs:      Lot Column ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RemoveDetails(ByVal strLotColumnID As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim Detail As cBrokerMessage        ' Detail object
    
    For lIndex = m.LotDetails.Count To 1 Step -1
        Set Detail = m.LotDetails(lIndex)
        
        If Detail("LotColumnID") = strLotColumnID Then
            m.LotDetails.Remove lIndex
        End If
    Next lIndex

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmEditLot.RemoveDetails"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetDetails
'' Description: Get the appropriate details out of the collection
'' Inputs:      Lot Column ID
'' Returns:     Details collection
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetDetails(ByVal strLotColumnID As String) As cGdTree
On Error GoTo ErrSection:

    Dim Details As cGdTree              ' Collection of details
    Dim Detail As cBrokerMessage        ' Detail object
    Dim lIndex As Long                  ' Index into a for loop
    
    Set Details = New cGdTree
    For lIndex = 1 To m.LotDetails.Count
        Set Detail = m.LotDetails(lIndex)
        
        If Detail("LotColumnID") = strLotColumnID Then
            Details.Add Detail, Detail("ID")
        End If
    Next lIndex
    
    Set GetDetails = Details

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmEditLot.GetDetails"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SortKey
'' Description: Determine the sort key for the given information
'' Inputs:      Subcategory ID, Column Name
'' Returns:     Sort Key
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SortKey(ByVal strSubcategoryID As String, ByVal strColumnName As String, ByVal lColumnOrdinal As Long) As String
On Error GoTo ErrSection:

    Dim strScOrdinal As String          ' Subcategory ordinal value
    
    strScOrdinal = m.Ordinal(strSubcategoryID)
    
    SortKey = Format(CLng(Val(strScOrdinal)), "00000") & "_" & Format(CLng(Val(strSubcategoryID)), "00000") & "_" & Format(lColumnOrdinal, "00000") & "_" & strColumnName

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmEditLot.SortKey"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetOwner
'' Description: Set the owner in the grid
'' Inputs:      Owner number
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetOwner(ByVal strOwnerNumber As String)
On Error GoTo ErrSection:

    Dim strOwnerName As String          ' Owner name
    
    strOwnerName = m.Owners(strOwnerNumber)
    
    With fgTab(m.lOwnerNameTab)
        .TextMatrix(m.lOwnerNameRow, GDCol(eGDCol_ActualValue)) = strOwnerName
        .TextMatrix(m.lOwnerNameRow, GDCol(eGDCol_DisplayValue)) = strOwnerName
    End With

    With fgTab(m.lOwnerNumberTab)
        .TextMatrix(m.lOwnerNumberRow, GDCol(eGDCol_ActualValue)) = strOwnerNumber
        .TextMatrix(m.lOwnerNumberRow, GDCol(eGDCol_DisplayValue)) = strOwnerNumber
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLot.SetOwner"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EditDetails
'' Description: Edit lot content details for an individual column
'' Inputs:      Index of the Grid, Row, Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EditDetails(ByVal Index As Long, ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim LotColumn As cLotColumn         ' Lot column object
    Dim lFeedYardLotID As Long          ' Feed Yard Lot ID
    Dim Details As cGdTree              ' Lot details
    Dim Detail As cBrokerMessage        ' Detail object
    Dim lIndex As Long                  ' Index into a for loop
    Dim strTotal As String              ' Total for the details

    With fgTab(Index)
        lFeedYardLotID = CLng(Val(m.Lot("FeedYardLotID")))
        Set LotColumn = .RowData(Row)

        Set Details = GetDetails(Str(LotColumn.ID))
        If Details.Count = 0 Then
            If lFeedYardLotID > 0 Then
                Set Detail = New cBrokerMessage
                Detail.Add "FeedYardID", Str(m.lFeedYardID)
                Detail.Add "FeedYardLotID", Str(lFeedYardLotID)
                Detail.Add "LotColumnID", Str(LotColumn.ID)
                Select Case UCase(LotColumn.KeyValueField)
                    Case "HEADIN"
                        Detail.Add "Date", m.Lot("DateIn")
                    Case "SHIPPED"
                        Detail.Add "Date", m.Lot("DateOut")
                    Case Else
                        Detail.Add "Date", m.Lot("ProcessDt")
                End Select
                If Len(Detail("Date")) = 0 Then
                    Detail("Date") = Str(Date)
                End If
                
                Detail.Add "Value", .TextMatrix(Row, GDCol(eGDCol_ActualValue))
                Detail.Add "Notes", ""

                Details.Add Detail
            End If
        End If

        If frmEditLotContentDetails.ShowMe(m.lFeedYardID, lFeedYardLotID, LotColumn, Details, m.strDefaultKeyValueField) Then
            RemoveDetails Str(LotColumn.ID)

            For lIndex = 1 To Details.Count
                Set Detail = Details(lIndex)
                m.LotDetails.Add Detail, Detail("ID")
            Next lIndex
            
            strTotal = g.Cattle.CalcTotalForDetails(Details, LotColumn)
            SetValue Index, Row, strTotal, LotColumn
            
            CalculateStats
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLot.EditDetails"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EditDetailsForCategory
'' Description: Edit lot content details for a category
'' Inputs:      Index of the Grid, Row, Column, Owner mode?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EditDetailsForCategory(ByVal Index As Long, ByVal Row As Long, ByVal Col As Long, ByVal bOwner As Boolean)
On Error GoTo ErrSection:

    Dim lParentRow As Long              ' Parent row in the grid
    Dim lChildRow As Long               ' Child row in the grid
    Dim LotColumn As cLotColumn         ' Lot column object
    Dim LotColumns As cGdTree           ' Collection of lot columns
    Dim lFeedYardLotID As Long          ' Feed Yard Lot ID
    Dim Details As cGdTree              ' Lot details
    Dim CategoryDetails As cGdTree      ' Lot details for the category
    Dim Detail As cBrokerMessage        ' Detail object
    Dim lIndex As Long                  ' Index into a for loop
    Dim bHasDateIn As Boolean           ' Does the category have "DateIn"?
    Dim bHasDateOut As Boolean          ' Does the cateogry have "DateOut"?
    Dim strTotal As String              ' Total for the details

    With fgTab(Index)
        lFeedYardLotID = CLng(Val(m.Lot("FeedYardLotID")))
        Set CategoryDetails = New cGdTree
        Set LotColumns = New cGdTree
        
        bHasDateIn = False
        bHasDateOut = False
                
        lParentRow = .GetNodeRow(Row, flexNTParent)
        lChildRow = .GetNodeRow(lParentRow, flexNTFirstChild)
        Do While lChildRow > -1&
            Set LotColumn = .RowData(lChildRow)
            
            If LotColumn.AggregateIsGroup Or LotColumn.AggregateIsOwner Then
                Set Details = GetDetails(Str(LotColumn.ID))
                If Details.Count > 0 Then
                    For lIndex = 1 To Details.Count
                        CategoryDetails.Add Details(lIndex), Details.Key(lIndex)
                    Next lIndex
                
                ElseIf (UCase(LotColumn.Format) <> "DATE") And (Len(.TextMatrix(lChildRow, GDCol(eGDCol_ActualValue))) > 0) Then
                    Set Detail = New cBrokerMessage
                    
                    Detail.Add "FeedYardID", Str(m.lFeedYardID)
                    Detail.Add "FeedYardLotID", Str(lFeedYardLotID)
                    Detail.Add "LotColumnID", Str(LotColumn.ID)
                    
                    If bHasDateIn Then
                        Detail.Add "Date", m.Lot("DateIn")
                    ElseIf bHasDateOut Then
                        Detail.Add "Date", m.Lot("DateOut")
                    Else
                        Detail.Add "Date", m.Lot("ProcessDt")
                    End If
                    If Val(Detail("Date")) = 0 Then
                        Detail("Date") = Str(Date)
                    End If
                    
                    Detail.Add "Value", .TextMatrix(lChildRow, GDCol(eGDCol_ActualValue))
                    Detail.Add "Notes", ""
                    
                    CategoryDetails.Add Detail
                End If
                
                LotColumns.Add LotColumn, Str(LotColumn.ID)
            End If
            
            If UCase(LotColumn.KeyValueField) = "DATEIN" Then
                bHasDateIn = True
            ElseIf UCase(LotColumn.KeyValueField) = "DATEOUT" Then
                bHasDateOut = True
            End If
            
            lChildRow = .GetNodeRow(lChildRow, flexNTNextSibling)
        Loop
        
        If frmEditLotContentDetails.ShowMeCategory(m.lFeedYardID, lFeedYardLotID, LotColumns, .TextMatrix(lParentRow, 0), bOwner, CategoryDetails, m.strDefaultKeyValueField) Then
            lChildRow = .GetNodeRow(lParentRow, flexNTFirstChild)
            Do While lChildRow > -1&
                Set LotColumn = .RowData(lChildRow)
                
                If UCase(LotColumn.KeyValueField) = "DATEIN" Then
                    If CategoryDetails.Count > 0 Then
                        Set Detail = CategoryDetails(1)
                        SetValue Index, lChildRow, Detail("Date"), LotColumn
                    End If
                ElseIf UCase(LotColumn.KeyValueField) = "DATEOUT" Then
                    If CategoryDetails.Count > 0 Then
                        Set Detail = CategoryDetails(CategoryDetails.Count)
                        SetValue Index, lChildRow, Detail("Date"), LotColumn
                    End If
                ElseIf LotColumn.AggregateIsGroup Or LotColumn.AggregateIsOwner Then
                    RemoveDetails Str(LotColumn.ID)
                    
                    For lIndex = 1 To CategoryDetails.Count
                        Set Detail = CategoryDetails(lIndex)
                        
                        If Detail("LotColumnID") = Str(LotColumn.ID) Then
                            m.LotDetails.Add Detail, Detail("ID")
                        End If
                    Next lIndex
                
                    strTotal = g.Cattle.CalcTotalForDetails(CategoryDetails, LotColumn)
                    SetValue Index, lChildRow, strTotal, LotColumn
                End If
        
                lChildRow = .GetNodeRow(lChildRow, flexNTNextSibling)
            Loop
            
            CalculateStats
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLot.EditDetailsForCategory"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    VerifyData
'' Description: Verify the data before OKing the dialog
'' Inputs:      None
'' Returns:     True if all ok, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function VerifyData() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = True
    If Len(fgTab(m.lOwnerNumberTab).TextMatrix(m.lOwnerNumberRow, GDCol(eGDCol_DisplayValue))) = 0 Then
        tabLotInfo.CurrTab = m.lOwnerNumberTab
        
        InfBox "Please specify an owner for this lot", "!", , "Error"
        bReturn = False
    End If
    
    VerifyData = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmEditLot.VerifyData"
    
End Function

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
    RaiseError "frmEditLot.AddOwner"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalculateStats
'' Description: Calculate the feedyard statistics
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalculateStats()
On Error GoTo ErrSection:

    Dim LotColumns As cGdTree           ' Lot columns

    If FormIsLoaded("frmLots") Then
        Set LotColumns = frmLots.LotColumns
        m.CalcStats.Calculate m.LotDetails, LotColumns
        
        Value("CalendarDaysOnFeed") = Str(m.CalcStats.CalendarDaysOnFeed)
        Value("TotalHeadDays") = Str(m.CalcStats.TotalHeadDays)
        Value("FinalADG") = Str(m.CalcStats.FinalADG)
        Value("FeedCost") = Str(m.CalcStats.FeedCostPerCwt)
        Value("CurrentBreakEven") = Str(m.CalcStats.CurrentBreakEven)
        Value("FinalBreakEven") = Str(m.CalcStats.FinalBreakEven)
        Value("CostGain") = Str(m.CalcStats.CostOfGain)
        Value("FinalCostOfGain") = Str(m.CalcStats.FinalCostOfGain)
        Value("AveragePayWeight") = Str(m.CalcStats.AveragePayWeight)
        Value("PriceInPerCwt") = Str(m.CalcStats.CattleCostPerCwt)
        
        CalculateHedgingColumns
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLot.CalculateStats"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HandleEdit
'' Description: Handle the user wanting to edit something on the grid
'' Inputs:      Index of the Grid, Row, Column
'' Returns:     True if edit, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function HandleEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim pt As POINTAPI                  ' Mouse location point
    Dim dDate As Double                 ' Date to send to the edit date form
    Dim LotColumn As cLotColumn         ' Lot column object
    Dim lFeedYardLotID As Long          ' Feed Yard Lot ID
    Dim Details As cGdTree              ' Lot details
    Dim Detail As cBrokerMessage        ' Detail object
    Dim lIndex As Long                  ' Index into a for loop
    Dim bCancelled As Boolean           ' Was the dialog cancelled?
    Dim strOwner As String              ' Owner selected
    Dim dTotal As Double                ' Total of the values in the details

    bReturn = False
    With fgTab(Index)
        If Col <> GDCol(eGDCol_DisplayValue) Then
            Col = GDCol(eGDCol_DisplayValue)
            .Col = Col
        End If
        
        If (Row >= .FixedRows) And (Row < .Rows) And (Col >= .FixedCols) And (Col < .Cols) Then
            pt.X = .ColPos(Col) / Screen.TwipsPerPixelX
            'pt.Y = (.RowPos(Row) - ((frmEditDate.Height - .RowHeight(Row)) / 2)) / Screen.TwipsPerPixelY
            pt.Y = (.RowPos(Row) + .RowHeight(Row)) / Screen.TwipsPerPixelY
            ClientToScreen .hWnd, pt
        
            If TypeOf .RowData(Row) Is cLotColumn Then
                Set LotColumn = .RowData(Row)
                m.strDefaultKeyValueField = LotColumn.KeyValueField
                
                bReturn = True
                If LotColumn.AggregateIsGroup Then
                    EditDetailsForCategory Index, Row, Col, False
                    
                ElseIf LotColumn.AggregateIsSingle Then
                    EditDetails Index, Row, Col
                    
                ElseIf LotColumn.AggregateIsOwner Then
                    EditDetailsForCategory Index, Row, Col, True
                    
                ElseIf UCase(LotColumn.Format) = "BOOLEAN" Then
                    CheckedCell(fgTab(Index), Row, GDCol(eGDCol_DisplayValue)) = Not CheckedCell(fgTab(Index), Row, GDCol(eGDCol_DisplayValue))
                    fgTab(Index).TextMatrix(Row, GDCol(eGDCol_ActualValue)) = g.Cattle.BoolToString(CheckedCell(fgTab(Index), Row, GDCol(eGDCol_DisplayValue)))
                
                ElseIf UCase(LotColumn.Format) = "DATE" Then
                    pt.X = pt.X * Screen.TwipsPerPixelX
                    pt.Y = pt.Y * Screen.TwipsPerPixelY
                    dDate = Val(.TextMatrix(Row, GDCol(eGDCol_ActualValue)))
                    If dDate = 0 Then dDate = Date
                    
                    frmEditDate.BackColor = BackColor
                    dDate = frmEditDate.ShowMe(pt.X, pt.Y, dDate, Me, , , , , , , , , , , bCancelled)
                    If bCancelled = False Then
                        SetValue Index, Row, Str(dDate), LotColumn
                    End If
                ElseIf UCase(LotColumn.KeyValueField) = "OWNERNUMBER" Then
                    'strOwner = g.AppBridge.AccountLookup(m.astrOwners, .TextMatrix(Row, Col), , True)
                    strOwner = frmOwnerLookup.ShowMe(m.astrOwners, .TextMatrix(Row, Col), , True)
                    If Len(strOwner) > 0 Then
                        SetOwner strOwner
                    End If
                ElseIf UCase(LotColumn.KeyValueField) = "OWNERNAME" Then
                    'strOwner = g.AppBridge.AccountLookup(m.astrOwners, , .TextMatrix(Row, Col), True)
                    strOwner = frmOwnerLookup.ShowMe(m.astrOwners, , .TextMatrix(Row, Col), True)
                    If Len(strOwner) > 0 Then
                        SetOwner strOwner
                    End If
                Else
                    bReturn = False
                End If
            End If
        End If
    End With
    
    HandleEdit = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmTunkeyEditLot.HandleEdit"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsDialogRow
'' Description: Determine if the given row in the given grid is a dialog row
'' Inputs:      Index of the Grid, Row
'' Returns:     True if dialog row, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IsDialogRow(ByVal Index As Long, ByVal Row As Long) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim LotColumn As cLotColumn         ' Lot column object
    
    bReturn = False
    With fgTab(Index)
        If TypeOf .RowData(Row) Is cLotColumn Then
            Set LotColumn = .RowData(Row)
        
            If LotColumn.IsAggregate > 0 Then
                bReturn = True
                
            ElseIf UCase(LotColumn.Format) = "DATE" Then
                bReturn = True
            
            ElseIf UCase(LotColumn.KeyValueField) = "OWNERNUMBER" Then
                bReturn = True
            
            ElseIf UCase(LotColumn.KeyValueField) = "OWNERNAME" Then
                bReturn = True
            
            End If
        End If
    End With
    
    IsDialogRow = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmEditLot.IsDialogRow"
    
End Function

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

    Dim iTab As Integer                 ' Tab number for the field
    Dim lRow As Long                    ' Row number for the field

    If m.LotColumnMap.Exists(strKeyValueField) Then
        iTab = Int(Val(Parse(m.LotColumnMap(strKeyValueField), ";", 1)))
        lRow = CLng(Val(Parse(m.LotColumnMap(strKeyValueField), ";", 2)))
        
        tabLotInfo.CurrTab = iTab
        fgTab(iTab).Row = lRow
        fgTab(iTab).Col = GDCol(eGDCol_DisplayValue)
        
        If bDoEdit Then
            If HandleEdit(iTab, lRow, GDCol(eGDCol_DisplayValue)) = False Then
                fgTab(iTab).EditCell
            End If
        End If
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmEditLot.GotoField"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetValue
'' Description: Set the actual and formatted values on the given row
'' Inputs:      Index, Row, Actual Value, Lot Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetValue(ByVal Index As Long, ByVal Row As Long, ByVal strActualValue As String, ByVal LotColumn As cLotColumn)
On Error GoTo ErrSection:

    fgTab(Index).TextMatrix(Row, GDCol(eGDCol_ActualValue)) = strActualValue
    g.Cattle.GridValue(fgTab(Index), Row, GDCol(eGDCol_DisplayValue), LotColumn) = strActualValue

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLot.SetValue"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalculateHedgingColumns
'' Description: Calculate the values for the hedging columns
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalculateHedgingColumns()
On Error GoTo ErrSection:

    Dim lHeadIn As Long                 ' Total number of head in
    Dim lNumShipped As Long             ' Total number of cattle shiped
    Dim lNumDead As Long                ' Total number of dead cattle
    Dim dProjectedWeightOut As Double   ' Projected weight out
    Dim dTotalSalesWeight As Double     ' Total sales weight
    Dim dWeightOut As Double            ' Weight out
    Dim dRequiredContracts As Double    ' Required contracts
    Dim dActualContracts As Double      ' Actual contracts
    Dim dPercentHedged As Double        ' Percent hedged
    Dim dContractSizeLe As Double       ' Contract size for the LE
    Dim dContractSizeGf As Double       ' Contract size for the GF
    
    lHeadIn = CLng(Val(Value("HeadIn")))
    lNumShipped = CLng(Val(Value("NumberShip")))
    lNumDead = CLng(Val(Value("NumberDeads")))
    dProjectedWeightOut = Val(Value("ProjectedWeightOut"))
    dTotalSalesWeight = Val(Value("TotalSalesWeight"))
    dActualContracts = Val(Value("ActualContracts"))
    dContractSizeLe = GetIniFileProperty("LE", 40000#, "ContractSize", AddSlash(g.strAppPath) & "Provided\Provided.INI")
    dContractSizeGf = GetIniFileProperty("GF", 50000#, "ContractSize", AddSlash(g.strAppPath) & "Provided\Provided.INI")
    
    dWeightOut = dTotalSalesWeight + (dProjectedWeightOut * (lHeadIn - lNumShipped - lNumDead))
    dRequiredContracts = dWeightOut / dContractSizeLe
    If dRequiredContracts = 0# Then
        dPercentHedged = 0#
    Else
        dPercentHedged = (dActualContracts / dRequiredContracts) * 100
    End If
    
    Value("RequiredContracts") = Format(dRequiredContracts, "#0.00")
    Value("PercentHedged") = Format(dPercentHedged, "#0.00")

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmEditLot.CalculateHedgingColumns"
    
End Sub

