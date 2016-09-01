VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmCattleReport 
   Caption         =   "Form1"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   10320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ActiveToolBars.SSActiveToolBars tbToolbar 
      Left            =   9000
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131083
      ToolBarsCount   =   1
      ToolsCount      =   5
      Tools           =   "frmCattleReport.frx":0000
      ToolBars        =   "frmCattleReport.frx":1A46
   End
   Begin VSFlex7LCtl.VSFlexGrid fgReport 
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   5895
      _cx             =   10398
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
   Begin HexUniControls.ctlUniFrameWL fraFilter 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8715
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
      Caption         =   "frmCattleReport.frx":1B94
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmCattleReport.frx":1BC0
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmCattleReport.frx":1BE0
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniComboImageXP cboReports 
         Height          =   315
         Left            =   600
         TabIndex        =   2
         Top             =   0
         Width           =   2235
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ButtonBackColor =   -2147483633
         ButtonForeColor =   -2147483630
         ButtonStyle     =   -1
         SelectorStyle   =   -1
         SelBackColor    =   -2147483635
         SelForeColor    =   -2147483634
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
         Tip             =   "frmCattleReport.frx":1BFC
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmCattleReport.frx":1C1C
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboCustomers 
         Height          =   315
         Left            =   6480
         TabIndex        =   3
         Top             =   0
         Width           =   2235
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ButtonBackColor =   -2147483633
         ButtonForeColor =   -2147483630
         ButtonStyle     =   -1
         SelectorStyle   =   -1
         SelBackColor    =   -2147483635
         SelForeColor    =   -2147483634
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
         Tip             =   "frmCattleReport.frx":1C38
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmCattleReport.frx":1C58
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniComboImageXP cboLots 
         Height          =   315
         Left            =   3300
         TabIndex        =   4
         Top             =   0
         Width           =   2235
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ButtonBackColor =   -2147483633
         ButtonForeColor =   -2147483630
         ButtonStyle     =   -1
         SelectorStyle   =   -1
         SelBackColor    =   -2147483635
         SelForeColor    =   -2147483634
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
         Tip             =   "frmCattleReport.frx":1C74
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmCattleReport.frx":1C94
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblReport 
         Height          =   195
         Left            =   0
         Top             =   60
         Width           =   555
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
         Caption         =   "frmCattleReport.frx":1CB0
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmCattleReport.frx":1CE0
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmCattleReport.frx":1D00
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblCustomer 
         Height          =   195
         Left            =   5700
         Top             =   60
         Width           =   855
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
         Caption         =   "frmCattleReport.frx":1D1C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmCattleReport.frx":1D50
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmCattleReport.frx":1D70
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblLot 
         Height          =   195
         Left            =   2940
         Top             =   60
         Width           =   375
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
         Caption         =   "frmCattleReport.frx":1D8C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmCattleReport.frx":1DB6
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmCattleReport.frx":1DD6
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
End
Attribute VB_Name = "frmCattleReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmCattleReport.frm
'' Description: Form for allowing user to setup and view Turnkey reports
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 11/25/2013   DAJ         Show details in Turnkey reports
'' 11/26/2013   DAJ         Tweaked details; tweak to ExpandAll/CollapseAll
'' 12/03/2013   DAJ         Expand/Collapse level
'' 12/04/2013   DAJ         Fix for calculations; Tweaks
'' 12/04/2013   DAJ         Changes for ReasonIn, ReasonOut, DebitCOG, and CreditCOG
'' 03/07/2014   DAJ         Moved into NavCattle.DLL
'' 05/22/2014   DAJ         Renamed frmTurnkeyReport to frmCattleReport
'' 05/22/2014   DAJ         Renamed g.Turnkey to g.Cattle
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Enum eGDCols
    eGDCol_Type = 0
    eGDCol_Date = 1
    eGDCol_Memo = 2
    eGDCol_Debit = 3
    eGDCol_Credit = 4
    eGDCol_Balance = 5
    eGDCol_RowType = 6
    
    eGDCol_NumCols
End Enum

Private Const kExtendedCol = eGDCol_Type

Private Type mPrivate
    Lots As cGdTree                     ' Collection of feed lots
    LotDetails As cGdTree               ' Collection of feed lot details
    LotColumns As cGdTree               ' Collection of lot columns
    Trades As cGdTree                   ' Collection of trades

    lPrevColWidth As Long               ' Previous column width
    lGridFontSize As Long               ' Grid font size
    lExpandLevel As Long                ' Current expand level
End Type
Private m As mPrivate

Private Property Get GDCol(ByVal nCol As eGDCols) As Long
    GDCol = nCol
End Property

Private Property Get GridFontSize() As Long
    GridFontSize = m.lGridFontSize
End Property
Private Property Let GridFontSize(ByVal lGridFontSize As Long)
On Error GoTo ErrSection:

    If lGridFontSize <= 8 Then
        m.lGridFontSize = 8
    Else
        m.lGridFontSize = lGridFontSize
    End If
    
    With fgReport
        .Redraw = flexRDNone
        
        .FontSize = lGridFontSize
        .AutoSize 0, .Cols - 1, False, 75
        ExtendCustomColumn
        
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmCattleReport.GridFontSize.Let"
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Setup and show the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMe(ByVal Lots As cGdTree, ByVal LotDetails As cGdTree, ByVal LotColumns As cGdTree, ByVal Trades As cGdTree)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim LotColumn As cLotColumn         ' Lot column object

    Caption = g.Cattle.ProductName & " Reports"

    Set m.Lots = Lots
    Set m.LotDetails = LotDetails
    Set m.Trades = Trades
    
    Set m.LotColumns = New cGdTree
    For lIndex = 1 To LotColumns.Count
        Set LotColumn = LotColumns(lIndex)
        m.LotColumns.Add LotColumn, LotColumn.KeyValueField
    Next lIndex
    
    LoadCombos
    
    InitGridSummary
    LoadGridSummary
    
    ExpandAll m.lExpandLevel
    FilterGrid

    ShowForm Me, eForm_Modal, g.frmMain

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleReport.ShowMe"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PrintMe
'' Description: Allow the user to print the journals for the selected day
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PrintMe()
On Error GoTo ErrSection:

    ' Need more margin at the top for the header to show up...
    frmPrintPreview.ShowMe "Turnkey Report", Me, 0, 0.45, 0.25, 0.25, 0.25, False, False
            
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleReport.PrintMe"
            
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenerateReport
'' Description: Callback function for the print preview
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GenerateReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:

    With frmPrintPreview.vp
        .StartDoc
        g.AppBridge.DoPrintHeader
        
        .FontName = "Times New Roman"
        .FontSize = 14
        .FontBold = True
        .FontUnderline = False
        .TextAlign = taCenterMiddle
        
        .Text = g.Cattle.ProductName & " Reports"
        
        .FontBold = False
        .FontSize = 12
        .TextAlign = taLeftMiddle
        
        .Text = vbLf & vbLf
        
        .RenderControl = fgReport.hWnd
        
        .EndDoc
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleReport.GenerateReport"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboCustomers_Click
'' Description: Handle the user changing the Customer combo
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboCustomers_Click()
On Error GoTo ErrSection:

    If Visible Then
        FilterGrid
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleReport.cboCustomers_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cboLots_Click
'' Description: Handle the user changing the lot combo
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboLots_Click()
On Error GoTo ErrSection:

    If Visible Then
        FilterGrid
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleReport.cboLots_Click"
    
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

    Icon = g.AppBridge.Picture16(g.Cattle.IconName)
    
    g.Styler.StyleForm Me
    
    PlaceForm Me
    
    With cboReports
        .AddItem "Summary"
        .ListIndex = 0
    End With
    Enable cboReports, False

    With tbToolbar
        .Tools("ID_Print").Picture = g.AppBridge.Picture16(g.AppBridge.ToolbarIcon("ID_Print"))
        .Tools("ID_ExpandAll").Picture = g.AppBridge.Picture16(g.AppBridge.ToolbarIcon("kExpandAll"))
        .Tools("ID_CollapseAll").Picture = g.AppBridge.Picture16(g.AppBridge.ToolbarIcon("kCollapseAll"))
        .Tools("ID_TextIncrease").Picture = g.AppBridge.Picture16(g.AppBridge.ToolbarIcon("ID_TextIncrease"))
        .Tools("ID_TextDecrease").Picture = g.AppBridge.Picture16(g.AppBridge.ToolbarIcon("ID_TextDecrease"))
        
        .Tools("ID_ExpandAll").TooltipText = "Expand the tree one level"
        .Tools("ID_CollapseAll").TooltipText = "Collapse the tree one level"
    End With

    GridFontSize = GetIniFileProperty("FontSize", 8, "TurnkeyReports", g.strIniFile)
    m.lExpandLevel = GetIniFileProperty("ExpandLevel", 4&, "TurnkeyReports", g.strIniFile)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleReport.Form_Load"
    
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

    Dim lMinScaleWidth As Long          ' Minimum scale width for the form
    Dim lMinScaleHeight As Long         ' Minimum scale height for the form
    Dim lSpace As Long                  ' Space between controls
    
    lSpace = 60
    lMinScaleWidth = fraFilter.Width + (lSpace * 2)
    lMinScaleHeight = (fraFilter.Height * 10) + (lSpace * 3)
    
    If Not LimitFormSize(Me, lMinScaleWidth, lMinScaleHeight) Then
        With fraFilter
            .Move lSpace, lSpace
        End With
        
        With fgReport
            .Move lSpace, fraFilter.Height + (lSpace * 2), ScaleWidth - (lSpace * 2), ScaleHeight - fraFilter.Height - (lSpace * 3)
        End With
        
        ExtendCustomColumn
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Cleanup whent he form is unloaded
'' Inputs:      Cancel the Unload?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    SaveFormPlacement Me
    SetIniFileProperty "FontSize", GridFontSize, "TurnkeyReports", g.strIniFile
    SetIniFileProperty "ExpandLevel", m.lExpandLevel, "TurnkeyReports", g.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleReport.Form_Unload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tbToolbar_ToolClick
'' Description: Handle the user clicking on a button on the toolbar
'' Inputs:      Tool Clicked
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tbToolbar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
On Error GoTo ErrSection:

    Select Case UCase(Tool.ID)
        Case "ID_PRINT"
            PrintMe
            
        Case "ID_EXPANDALL"
            ExpandAll
        
        Case "ID_COLLAPSEALL"
            CollapseAll
            
        Case "ID_TEXTINCREASE"
            GridFontSize = GridFontSize + 1
        
        Case "ID_TEXTDECREASE"
            GridFontSize = GridFontSize - 1
        
    End Select

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmCattleReport.tbToolbar_ToolClick"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadCombos
'' Description: Load the lots and customers combos
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadCombos()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim Lot As cBrokerMessage           ' Lot object
    Dim astrCustomers As cGdArray       ' Collection of customers
    Dim strCustomer As String           ' Customer string
    Dim lPos As Long                    ' Position of the customer in the array
    
    cboLots.Clear
    cboCustomers.Clear
    
    Set astrCustomers = New cGdArray
    astrCustomers.Create eGDARRAY_Strings
    
    cboLots.AddItem "All Lots"
    cboLots.ItemData(cboLots.NewIndex) = -1&
    
    For lIndex = 1 To m.Lots.Count
        Set Lot = m.Lots(lIndex)
        
        cboLots.AddItem Lot("Number") & " (" & Lot("Name") & ")"
        cboLots.ItemData(cboLots.NewIndex) = CLng(Val(Lot("FeedYardLotID")))
        
        strCustomer = Lot("OwnerName") & " (" & Lot("OwnerNumber") & ")"
        If astrCustomers.BinarySearch(strCustomer, lPos) = False Then
            astrCustomers.Add strCustomer, lPos
        End If
    Next lIndex
    
    cboLots.ListIndex = 0&
    
    cboCustomers.AddItem "All Customers"
    cboCustomers.ItemData(cboCustomers.NewIndex) = -1&
    
    For lIndex = 0 To astrCustomers.Size - 1
        cboCustomers.AddItem astrCustomers(lIndex)
    Next lIndex
    
    cboCustomers.ListIndex = 0&

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleReport.LoadCombos"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitGridSummary
'' Description: Initialize the grid for a summary report
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitGridSummary()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid

    With fgReport
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .Editable = flexEDNone
        .ExplorerBar = flexExNone
        .ExtendLastCol = False
        .MergeCells = flexMergeSpill
        .OutlineBar = flexOutlineBarSimpleLeaf
        .ScrollTrack = True
        .SelectionMode = flexSelectionFree
        .SheetBorder = RGB(128, 128, 128)
        .WordWrap = False
        
        .Rows = 1
        .FixedRows = 1
        .Cols = GDCol(eGDCol_NumCols)
        .FixedCols = 0
        
        .TextMatrix(0, GDCol(eGDCol_Type)) = "Type"
        .TextMatrix(0, GDCol(eGDCol_Date)) = "Date"
        .TextMatrix(0, GDCol(eGDCol_Memo)) = "Memo"
        .TextMatrix(0, GDCol(eGDCol_Debit)) = "Debit"
        .TextMatrix(0, GDCol(eGDCol_Credit)) = "Credit"
        .TextMatrix(0, GDCol(eGDCol_Balance)) = "Balance"
        .TextMatrix(0, GDCol(eGDCol_RowType)) = "Row Type"
        
        .ColFormat(GDCol(eGDCol_Debit)) = "$#,##0.00"
        .ColFormat(GDCol(eGDCol_Credit)) = "$#,##0.00"
        .ColFormat(GDCol(eGDCol_Balance)) = "$#,##0.00"
        
        .ColHidden(GDCol(eGDCol_RowType)) = True
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleReport.InitGridSummary"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGridSummary
'' Description: Load the grid for a summary report
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadGridSummary()
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim lIndex As Long                  ' Index into a for loop
    Dim Lot As cBrokerMessage           ' Lot message
    Dim lTotalRow As Long               ' Row for the totals
    Dim lLotRow As Long                 ' Lot row

    With fgReport
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lIndex = 1 To m.Lots.Count
            Set Lot = m.Lots(lIndex)
            
            .Rows = .Rows + 1
            .RowData(.Rows - 1) = Lot
            .TextMatrix(.Rows - 1, GDCol(eGDCol_Type)) = Lot("Number") & " (" & Lot("Name") & ")"
            .TextMatrix(.Rows - 1, GDCol(eGDCol_RowType)) = "Lot"
            .RowOutlineLevel(.Rows - 1) = 0
            .IsSubtotal(.Rows - 1) = True
            .MergeRow(.Rows - 1) = True
            lLotRow = .Rows - 1
            
            lTotalRow = AddHeaderRow("Cattle Inventory")
            AddDataRow "TotalCostOfCattle", Lot
            If Val(Lot("NumberShip")) > 0 Then
                AddDataRow "TotalSalesAmount", Lot
            End If
            
            lTotalRow = AddHeaderRow("Finances")
            If LotColumnVisible("TotalFeedCost") Then
                AddDataRow "TotalFeedCost", Lot
            End If
            If LotColumnVisible("VetAndOtherCost") Then
                AddDataRow "VetAndOtherCost", Lot
            End If
            If LotColumnVisible("Interest") Then
                AddDataRow "Interest", Lot
            End If
            If LotColumnVisible("DebitCOG") Then
                AddDataRow "DebitCOG", Lot
            End If
            If LotColumnVisible("CreditCOG") Then
                AddDataRow "CreditCOG", Lot
            End If
            If LotColumnVisible("Debit") Then
                AddDataRow "Debit", Lot
            End If
            If LotColumnVisible("Credit") Then
                AddDataRow "Credit", Lot
            End If
            
            lTotalRow = AddHeaderRow("Hedges")
            AddTrades lTotalRow, Lot("FeedYardLotID")
        
            CalcLotTotals lLotRow
        Next lIndex
        
        .AutoSize 0, .Cols - 1, False, 75
        ExtendCustomColumn
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleReport.LoadGridSummary"
    
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

    Dim lIndex As Long                  ' Index into a for loop
    Dim lTotal As Long                  ' Total width
    Dim lDiff As Long                   ' Difference in column width

    With fgReport
        ' if column being resized is after the extended column,
        ' then change the width of the next visible column instead
        If nResizeCol >= kExtendedCol Then
            .Redraw = flexRDNone
            lDiff = .ColWidth(nResizeCol) - m.lPrevColWidth
            For lIndex = nResizeCol + 1 To .Cols - 1
                If Not .ColHidden(lIndex) Then
                    .ColWidth(lIndex) = .ColWidth(lIndex) - lDiff
                    Exit For
                End If
            Next
            m.lPrevColWidth = 0
        End If
        
        ' size the custom extended column in order to fill the client width
        .ColHidden(kExtendedCol) = True
        .Redraw = flexRDBuffered '(must do this so .ClientWidth will be correct)
        .Redraw = flexRDNone
        lTotal = 0
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
    RaiseError "frmCattleReport.ExtendCustomColumn"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddHeaderRow
'' Description: Add header row to the grid
'' Inputs:      Title
'' Returns:     Row in the Grid
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AddHeaderRow(ByVal strTitle As String, Optional ByVal lRowOutlineLevel As Long = 1&) As Long
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings from the grid
    Dim lReturn As Long                 ' Return value for the function

    With fgReport
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Rows = .Rows + 1
        lReturn = .Rows - 1
        
        .TextMatrix(.Rows - 1, GDCol(eGDCol_Type)) = strTitle
        .TextMatrix(.Rows - 1, GDCol(eGDCol_RowType)) = "Header"
        .RowOutlineLevel(.Rows - 1) = lRowOutlineLevel
        .IsSubtotal(.Rows - 1) = True
        .MergeRow(.Rows - 1) = True
        
        .Redraw = nRedraw
    End With
    
    AddHeaderRow = lReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCattleReport.AddHeaderRow"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddDataRow
'' Description: Add data row to the grid
'' Inputs:      Key/Value Field, Lot
'' Returns:     Value for the data
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AddDataRow(ByVal strKeyValueField As String, ByVal Lot As cBrokerMessage) As Double
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings from the grid
    Dim dReturn As Double               ' Return value for the function
    Dim bDebit As Boolean               ' Debit?
    Dim lRow As Long                    ' Row in the grid
    Dim LotColumn As cLotColumn         '

    With fgReport
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Rows = .Rows + 1
        lRow = .Rows - 1
        
        dReturn = Val(Lot(strKeyValueField))
        Select Case UCase(strKeyValueField)
            Case "TOTALCOSTOFCATTLE"
                bDebit = True
                .TextMatrix(lRow, GDCol(eGDCol_Type)) = "Cattle In"
                .TextMatrix(lRow, GDCol(eGDCol_Date)) = DateFormat(Val(Lot("DateIn")), MM_DD_YYYY)
                .TextMatrix(lRow, GDCol(eGDCol_Memo)) = Lot("HeadIn") & " head"
                SetCurrency lRow, GDCol(eGDCol_Debit), dReturn, False
            
            Case "TOTALSALESAMOUNT"
                bDebit = False
                .TextMatrix(lRow, GDCol(eGDCol_Type)) = "Cattle Out"
                .TextMatrix(lRow, GDCol(eGDCol_Date)) = DateFormat(Val(Lot("DateOut")), MM_DD_YYYY)
                .TextMatrix(lRow, GDCol(eGDCol_Memo)) = Lot("NumberShip") & " head"
                SetCurrency lRow, GDCol(eGDCol_Credit), dReturn, False
            
            Case "TOTALFEEDCOST"
                bDebit = True
                .TextMatrix(lRow, GDCol(eGDCol_Type)) = "Feed"
                SetCurrency lRow, GDCol(eGDCol_Debit), dReturn, False
            
            Case "VETANDOTHERCOST"
                bDebit = True
                .TextMatrix(lRow, GDCol(eGDCol_Type)) = "Veteranarian and Other"
                SetCurrency lRow, GDCol(eGDCol_Debit), dReturn, False
            
            Case "INTEREST"
                bDebit = True
                .TextMatrix(lRow, GDCol(eGDCol_Type)) = "Interest"
                SetCurrency lRow, GDCol(eGDCol_Debit), dReturn, False
                
            Case "DEBITCOG"
                bDebit = True
                .TextMatrix(lRow, GDCol(eGDCol_Type)) = "Debits COG"
                SetCurrency lRow, GDCol(eGDCol_Debit), dReturn, False
        
            Case "CREDITCOG"
                bDebit = False
                .TextMatrix(lRow, GDCol(eGDCol_Type)) = "Credits COG"
                SetCurrency lRow, GDCol(eGDCol_Credit), dReturn, False
        
            Case "DEBIT"
                bDebit = True
                .TextMatrix(lRow, GDCol(eGDCol_Type)) = "Debits"
                SetCurrency lRow, GDCol(eGDCol_Debit), dReturn, False
        
            Case "CREDIT"
                bDebit = False
                .TextMatrix(lRow, GDCol(eGDCol_Type)) = "Credits"
                SetCurrency lRow, GDCol(eGDCol_Credit), dReturn, False
        
        End Select
        
        .TextMatrix(.Rows - 1, GDCol(eGDCol_RowType)) = "Data"
        .RowOutlineLevel(lRow) = 2
        .IsSubtotal(lRow) = True
        .MergeRow(lRow) = False
        
        AddDetails strKeyValueField, Lot, bDebit
        
        .Redraw = nRedraw
    End With
    
    AddDataRow = dReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCattleReport.AddDataRow"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddDetails
'' Description: Add details to the grid for the given lot column ( if any )
'' Inputs:      Lot Column, Lot, Debit?
'' Returns:     Num Details
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AddDetails(ByVal strLotColumn As String, ByVal Lot As cBrokerMessage, ByVal bDebit As Boolean) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim nRedraw As RedrawSettings       ' Redraw settings from the grid
    Dim strLotColumnID As String        ' Lot column ID
    Dim LotColumn As cLotColumn         ' Lot column object
    Dim lIndex As Long                  ' Index into a for loop
    Dim LotDetail As cBrokerMessage     ' Lot column detail object
    Dim strType As String               ' Type
    Dim TypeDetail As cBrokerMessage    ' Detail object for the type
    Dim astrTypes As cGdArray           ' Array of types
    Dim lPos As Long                    ' Position in the array
    Dim lTypes As Long                  ' Index into a for loop
    
    lReturn = 0&
    If m.LotColumns.Exists(strLotColumn) Then
        Set LotColumn = m.LotColumns(strLotColumn)
        strLotColumnID = Str(LotColumn.ID)
        
        Select Case UCase(strLotColumn)
            Case "TOTALCOSTOFCATTLE"
                strType = "ReasonIn"
            Case "TOTALSALESAMOUNT"
                strType = "ReasonOut"
            Case "CREDITCOG"
                strType = "FinancialCategory"
            Case "DEBITCOG"
                strType = "FinancialCategory"
            Case "CREDIT"
                strType = "FinancialCategory"
            Case "DEBIT"
                strType = "FinancialCategory"
            Case "TOTALFEEDCOST"
                strType = "Ingredient"
            Case Else
                strType = ""
        End Select
        
        Set astrTypes = New cGdArray
        astrTypes.Create eGDARRAY_Strings
        
        If Len(strType) > 0 Then
            For lIndex = 1 To m.LotDetails.Count
                Set LotDetail = m.LotDetails(lIndex)
                
                If LotDetail("FeedYardID") = Lot("FeedYardID") Then
                    If LotDetail("FeedYardLotID") = Lot("FeedYardLotID") Then
                        If LotDetail("LotColumnID") = strLotColumnID Then
                            Set TypeDetail = GetDetail(strType, Lot, LotDetail("Date"))
                            If Not TypeDetail Is Nothing Then
                                If astrTypes.BinarySearch(TypeDetail("Value"), lPos) = False Then
                                    astrTypes.Add TypeDetail("Value"), lPos
                                End If
                            End If
                        End If
                    End If
                End If
            Next lIndex
        End If
        
        Set LotColumn = m.LotColumns(strLotColumn)
        strLotColumnID = Str(LotColumn.ID)
        
        With fgReport
            nRedraw = .Redraw
            .Redraw = flexRDNone
            
            If Len(strType) = 0 Then
                For lIndex = 1 To m.LotDetails.Count
                    Set LotDetail = m.LotDetails(lIndex)
                    
                    If LotDetail("FeedYardID") = Lot("FeedYardID") Then
                        If LotDetail("FeedYardLotID") = Lot("FeedYardLotID") Then
                            If LotDetail("LotColumnID") = strLotColumnID Then
                                lReturn = lReturn + 1&
                                .Rows = .Rows + 1
                                
                                .TextMatrix(.Rows - 1, GDCol(eGDCol_Date)) = DateFormat(Val(LotDetail("Date")), MM_DD_YYYY)
                                .TextMatrix(.Rows - 1, GDCol(eGDCol_Memo)) = LotDetail("Notes")
                                
                                If bDebit Then
                                    SetCurrency .Rows - 1, GDCol(eGDCol_Debit), Val(LotDetail("Value")), False
                                Else
                                    SetCurrency .Rows - 1, GDCol(eGDCol_Credit), Val(LotDetail("Value")), False
                                End If
                                
                                .TextMatrix(.Rows - 1, GDCol(eGDCol_RowType)) = "Detail"
                                .RowOutlineLevel(.Rows - 1) = 3
                                .IsSubtotal(.Rows - 1) = True
                                .MergeRow(.Rows - 1) = False
                            End If
                        End If
                    End If
                Next lIndex
            Else
                For lTypes = 0 To astrTypes.Size - 1
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, GDCol(eGDCol_Type)) = astrTypes(lTypes)
                    If bDebit Then
                        .TextMatrix(.Rows - 1, GDCol(eGDCol_RowType)) = "DebitType"
                    Else
                        .TextMatrix(.Rows - 1, GDCol(eGDCol_RowType)) = "CreditType"
                    End If
                    .RowOutlineLevel(.Rows - 1) = 3
                    .IsSubtotal(.Rows - 1) = True
                    .MergeRow(.Rows - 1) = False
                    
                    For lIndex = 1 To m.LotDetails.Count
                        Set LotDetail = m.LotDetails(lIndex)
                        
                        If LotDetail("FeedYardID") = Lot("FeedYardID") Then
                            If LotDetail("FeedYardLotID") = Lot("FeedYardLotID") Then
                                If LotDetail("LotColumnID") = strLotColumnID Then
                                    Set TypeDetail = GetDetail(strType, Lot, LotDetail("Date"))
                                    If Not TypeDetail Is Nothing Then
                                        If TypeDetail("Value") = astrTypes(lTypes) Then
                                            .Rows = .Rows + 1
                                            
                                            .TextMatrix(.Rows - 1, GDCol(eGDCol_Date)) = DateFormat(Val(LotDetail("Date")), MM_DD_YYYY)
                                            .TextMatrix(.Rows - 1, GDCol(eGDCol_Memo)) = LotDetail("Notes")
                                            .TextMatrix(.Rows - 1, GDCol(eGDCol_RowType)) = "Detail"
                                            
                                            If bDebit Then
                                                SetCurrency .Rows - 1, GDCol(eGDCol_Debit), Val(LotDetail("Value")), False
                                            Else
                                                SetCurrency .Rows - 1, GDCol(eGDCol_Credit), Val(LotDetail("Value")), False
                                            End If
                                            
                                            .RowOutlineLevel(.Rows - 1) = 4
                                            .IsSubtotal(.Rows - 1) = True
                                            .MergeRow(.Rows - 1) = False
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next lIndex
                Next lTypes
            End If
            
            .Redraw = nRedraw
        End With
    End If
    
    AddDetails = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCattleReport.AddDetails"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetCurrency
'' Description: Set a cell in the grid to a given currency value
'' Inputs:      Row, Column, Value
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetCurrency(ByVal Row As Long, ByVal Col As Long, ByVal dValue As Double, ByVal bColor As Boolean)
On Error GoTo ErrSection:

    With fgReport
        .TextMatrix(Row, Col) = dValue
        
        If (bColor = False) Or (dValue = 0) Then
            .Cell(flexcpForeColor, Row, Col) = vbBlack
        ElseIf dValue < 0 Then
            .Cell(flexcpForeColor, Row, Col) = vbRed
        Else
            .Cell(flexcpForeColor, Row, Col) = QBColor(2)
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleReport.SetCurrency"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalcLotTotals
'' Description: Calculate lot totals
'' Inputs:      Lot Row
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalcLotTotals(ByVal lLotRow As Long)
On Error GoTo ErrSection:

    Dim lHeaderRow As Long              ' Header row
    Dim lChildRow As Long               ' Child row
    Dim lTypeRow As Long                ' Type row
    Dim lDetailRow As Long              ' Detail row
    Dim dCredit As Double               ' Credit for the row
    Dim dDebit As Double                ' Debit for the row
    Dim dCreditDetail As Double         ' Total credit for the lot
    Dim dDebitDetail As Double          ' Total debit for the lot
    Dim dCreditType As Double           ' Total credit for the type
    Dim dDebitType As Double            ' Total debit for the type
    Dim dCreditCategory As Double       ' Total credit for the category
    Dim dDebitCategory As Double        ' Total debit for the category
    Dim dCreditTotal As Double          ' Total credit for the lot
    Dim dDebitTotal As Double           ' Total debit for the lot
    Dim strRowType As String            ' Row type
    
    dCreditTotal = 0#
    dDebitTotal = 0#
    
    With fgReport
        lHeaderRow = .GetNodeRow(lLotRow, flexNTFirstChild)
        
        Do While lHeaderRow <> -1&
            dCreditCategory = 0#
            dDebitCategory = 0#
            
            lChildRow = .GetNodeRow(lHeaderRow, flexNTFirstChild)
            Do While lChildRow <> -1&
                dCredit = .Cell(flexcpValue, lChildRow, GDCol(eGDCol_Credit))
                dDebit = .Cell(flexcpValue, lChildRow, GDCol(eGDCol_Debit))
                
                dCreditCategory = dCreditCategory + dCredit
                dDebitCategory = dDebitCategory + dDebit
                SetCurrency lChildRow, GDCol(eGDCol_Balance), dCreditCategory - dDebitCategory, True
                
                dCreditTotal = dCreditTotal + dCredit
                dDebitTotal = dDebitTotal + dDebit
                
                dCreditType = 0#
                dDebitType = 0#
                lTypeRow = .GetNodeRow(lChildRow, flexNTFirstChild)
                Do While lTypeRow <> -1&
                    strRowType = .TextMatrix(lTypeRow, GDCol(eGDCol_RowType))
                    
                    If strRowType = "Detail" Then
                        dCredit = .Cell(flexcpValue, lTypeRow, GDCol(eGDCol_Credit))
                        dDebit = .Cell(flexcpValue, lTypeRow, GDCol(eGDCol_Debit))
                        
                        dCreditType = dCreditType + dCredit
                        dDebitType = dDebitType + dDebit
                        SetCurrency lTypeRow, GDCol(eGDCol_Balance), dCreditType - dDebitType, True
                    Else
                        dCreditDetail = 0#
                        dDebitDetail = 0#
                        
                        lDetailRow = .GetNodeRow(lTypeRow, flexNTFirstChild)
                        Do While lDetailRow <> -1&
                            dCredit = .Cell(flexcpValue, lDetailRow, GDCol(eGDCol_Credit))
                            dDebit = .Cell(flexcpValue, lDetailRow, GDCol(eGDCol_Debit))
                            
                            dCreditDetail = dCreditDetail + dCredit
                            dDebitDetail = dDebitDetail + dDebit
                            SetCurrency lDetailRow, GDCol(eGDCol_Balance), dCreditDetail - dDebitDetail, True
                            
                            lDetailRow = .GetNodeRow(lDetailRow, flexNTNextSibling)
                        Loop
                    
                        dCreditType = dCreditType + dCreditDetail
                        dDebitType = dDebitType + dDebitDetail
                        
                        If (strRowType = "CreditType") Or (dCreditType > 0#) Then
                            SetCurrency lTypeRow, GDCol(eGDCol_Credit), dCreditType, False
                        End If
                        If (strRowType = "DebitType") Or (dDebitType > 0#) Then
                            SetCurrency lTypeRow, GDCol(eGDCol_Debit), dDebitType, False
                        End If
                        SetCurrency lTypeRow, GDCol(eGDCol_Balance), dCreditType - dDebitType, True
                    End If
                
                    lTypeRow = .GetNodeRow(lTypeRow, flexNTNextSibling)
                Loop
                
                lChildRow = .GetNodeRow(lChildRow, flexNTNextSibling)
            Loop
            
            SetCurrency lHeaderRow, GDCol(eGDCol_Credit), dCreditCategory, False
            SetCurrency lHeaderRow, GDCol(eGDCol_Debit), dDebitCategory, False
            SetCurrency lHeaderRow, GDCol(eGDCol_Balance), dCreditCategory - dDebitCategory, True
            
            lHeaderRow = .GetNodeRow(lHeaderRow, flexNTNextSibling)
        Loop
    
        SetCurrency lLotRow, GDCol(eGDCol_Credit), dCreditTotal, False
        SetCurrency lLotRow, GDCol(eGDCol_Debit), dDebitTotal, False
        SetCurrency lLotRow, GDCol(eGDCol_Balance), dCreditTotal - dDebitTotal, True
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleReport.CalcLotTotals"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddTrades
'' Description: Add the trades for the hedges
'' Inputs:      Hedge Row, Feed Yard Lot ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddTrades(ByVal lHedgeRow As Long, ByVal strFeedYardLotID As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim Trade As cBrokerMessage         ' Trade object
    Dim strSecurityType As String       ' Security type
    Dim dValue As Double                ' Value
    Dim strEntry As String              ' Entry string
    Dim strExit As String               ' Exit string

    With fgReport
        For lIndex = 1 To m.Trades.Count
            Set Trade = m.Trades(lIndex)
            
            If Trade("FeedYardLotID") = strFeedYardLotID Then
                strEntry = g.Cattle.TradeEntryToString(Trade, , , False)
                strExit = g.Cattle.TradeExitToString(Trade, , , False)
                
                .Rows = .Rows + 1
                
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Type)) = strEntry
                .TextMatrix(.Rows - 1, GDCol(eGDCol_Date)) = DateFormat(Val(Trade("EntryFillTime")), MM_DD_YYYY)
                                
                strSecurityType = g.AppBridge.SecurityType(Trade("Symbol"), True)
                If strSecurityType = "FO" Then
                    If Trade("IsBuy") = "0" Then
                        dValue = g.AppBridge.Profit(Trade("Symbol"), Val(Trade("EntryFillPrice")), CLng(Val(Trade("Quantity"))))
                        SetCurrency .Rows - 1, GDCol(eGDCol_Debit), dValue, False
                    Else
                        dValue = g.AppBridge.Profit(Trade("Symbol"), Val(Trade("EntryFillPrice")), CLng(Val(Trade("Quantity"))))
                        SetCurrency .Rows - 1, GDCol(eGDCol_Credit), dValue, False
                    End If
                End If
                                
                .RowOutlineLevel(.Rows - 1) = 2
                .IsSubtotal(.Rows - 1) = True
                .MergeRow(.Rows - 1) = True
                
                If Len(strExit) > 0 Then
                    .Rows = .Rows + 1
                    
                    .TextMatrix(.Rows - 1, GDCol(eGDCol_Type)) = strExit
                    .TextMatrix(.Rows - 1, GDCol(eGDCol_Date)) = DateFormat(Val(Trade("ExitFillTime")), MM_DD_YYYY)
                                    
                    strSecurityType = g.AppBridge.SecurityType(Trade("Symbol"), True)
                    If strSecurityType = "FO" Then
                        If Trade("IsBuy") = "0" Then
                            dValue = g.AppBridge.Profit(Trade("Symbol"), Val(Trade("ExitFillPrice")), CLng(Val(Trade("Quantity"))))
                            SetCurrency .Rows - 1, GDCol(eGDCol_Credit), dValue, False
                        Else
                            dValue = g.AppBridge.Profit(Trade("Symbol"), Val(Trade("ExitFillPrice")), CLng(Val(Trade("Quantity"))))
                            SetCurrency .Rows - 1, GDCol(eGDCol_Debit), dValue, False
                        End If
                    Else
                        dValue = Val(Trade("ClosedProfit"))
                        If dValue < 0 Then
                            SetCurrency .Rows - 1, GDCol(eGDCol_Debit), Abs(dValue), False
                        Else
                            SetCurrency .Rows - 1, GDCol(eGDCol_Credit), Abs(dValue), False
                        End If
                    End If
                                    
                    .RowOutlineLevel(.Rows - 1) = 2
                    .IsSubtotal(.Rows - 1) = True
                    .MergeRow(.Rows - 1) = True
                End If
            End If
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleReport.AddTrades"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FilterGrid
'' Description: Filter the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FilterGrid()
On Error GoTo ErrSection:

    Dim strLot As String                ' Lot chosen by the user
    Dim strCustomer As String           ' Customer chosen by the user
    Dim Lot As cBrokerMessage           ' Lot object
    Dim lRow As Long                    ' Row int he grid
    Dim nRedraw As RedrawSettings       ' Redraw settings
    Dim bHide As Boolean                ' Hide the lot?
    Dim lIndex As Long                  ' Index into a for loop
    Dim lFrom As Long                   ' From value for the for loop
    Dim lTo As Long                     ' To value for the for loop
    Dim nCollapsed As CollapsedSettings ' Collapsed settings
    
    With fgReport
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        strLot = ""
        If cboLots.ListIndex > 0 Then
            strLot = Str(cboLots.ItemData(cboLots.ListIndex))
        End If
        
        strCustomer = ""
        If cboCustomers.ListIndex > 0 Then
            strCustomer = Parse(cboCustomers.Text, "(", 1)
        End If
        
        lRow = .FixedRows
        Do While Not lRow = -1&
            Set Lot = .RowData(lRow)
            
            nCollapsed = .IsCollapsed(lRow)
            bHide = False
            If (Len(strLot) > 0) And (Lot("FeedYardLotID") <> strLot) Then
                bHide = True
            End If
            If (Len(strCustomer) > 0) And (Lot("OwnerName") <> strCustomer) Then
                bHide = True
            End If
            
            lFrom = lRow
            lRow = .GetNodeRow(lRow, flexNTNextSibling)
            If lRow = -1& Then
                lTo = .Rows - 1
            Else
                lTo = lRow - 1
            End If
            
            For lIndex = lFrom To lTo
                .RowHidden(lIndex) = bHide
            Next lIndex
            
            If bHide = False Then
                .IsCollapsed(lFrom) = nCollapsed
            End If
        Loop
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleReport.FilterGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExpandAll
'' Description: Expand all of the nodes in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ExpandAll(Optional lLevel As Long = -1)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    
    With fgReport
        If .Rows > .FixedRows Then
            nRedraw = .Redraw
            .Redraw = flexRDNone
            
            If lLevel = -1& Then
                m.lExpandLevel = SetGridLevel(fgReport, m.lExpandLevel + 1)
            Else
                m.lExpandLevel = SetGridLevel(fgReport, lLevel)
            End If
            
            FilterGrid
            
            .Redraw = nRedraw
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleReport.ExpandAll"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CollapseAll
'' Description: Collapse all of the nodes in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CollapseAll(Optional lLevel As Long = -1)
On Error GoTo ErrSection:

    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    
    With fgReport
        If .Rows > .FixedRows Then
            nRedraw = .Redraw
            .Redraw = flexRDNone
            
            If lLevel = -1& Then
                m.lExpandLevel = SetGridLevel(fgReport, m.lExpandLevel - 1)
            Else
                m.lExpandLevel = SetGridLevel(fgReport, lLevel)
            End If
            
            FilterGrid
            
            .Redraw = nRedraw
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCattleReport.CollapseAll"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LotColumnVisible
'' Description: Determine if the given lot column should be visible
'' Inputs:      Key Value Field Name
'' Returns:     True if visible, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LotColumnVisible(ByVal strKeyValueField As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim LotColumn As cLotColumn         ' Lot column object
    
    bReturn = False
    If m.LotColumns.Exists(strKeyValueField) Then
        Set LotColumn = m.LotColumns(strKeyValueField)
        bReturn = Not (LotColumn.AlwaysHidden Or LotColumn.FeedyardHidden)
    End If
    
    LotColumnVisible = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCattleReport.LotColumnVisible"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetDetail
'' Description: Get the detail with the given name and date
'' Inputs:      Key Value Field Name, Lot, Date
'' Returns:     Detail if found ( Nothing otherwise )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetDetail(ByVal strLotColumn As String, ByVal Lot As cBrokerMessage, ByVal strDate As String) As cBrokerMessage
On Error GoTo ErrSection:

    Dim ReturnDetail As cBrokerMessage  ' Return value for the function
    Dim LotDetail As cBrokerMessage     ' Detail object
    Dim lIndex As Long                  ' Index into a for loop
    Dim strLotColumnID As String        ' Lot column ID
    Dim LotColumn As cLotColumn         ' Lot column object
    
    Set ReturnDetail = Nothing
    Set LotColumn = m.LotColumns(strLotColumn)
    strLotColumnID = Str(LotColumn.ID)
    
    For lIndex = 1 To m.LotDetails.Count
        Set LotDetail = m.LotDetails(lIndex)

        If LotDetail("FeedYardID") = Lot("FeedYardID") Then
            If LotDetail("FeedYardLotID") = Lot("FeedYardLotID") Then
                If LotDetail("LotColumnID") = strLotColumnID Then
                    If LotDetail("Date") = strDate Then
                        Set ReturnDetail = LotDetail
                        Exit For
                    End If
                End If
            End If
        End If
    Next lIndex
    
    Set GetDetail = ReturnDetail

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCattleReport.GetDetail"
    
End Function

