VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmSnapshot 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   585
   ClientWidth     =   8340
   Icon            =   "frmSnapshot.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7320
      Top             =   4320
   End
   Begin vsOcx6LibCtl.vsIndexTab vsTab 
      Height          =   3000
      Left            =   570
      TabIndex        =   0
      Top             =   735
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   5292
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
      Caption         =   "Custom|Cash Flows|Balance Sheet|Income|General"
      Align           =   0
      Appearance      =   1
      CurrTab         =   -1
      FirstTab        =   0
      Style           =   3
      Position        =   0
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   -1  'True
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   0
      BorderWidth     =   0
      BoldCurrent     =   0   'False
      DogEars         =   -1  'True
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      Begin HexUniControls.ctlUniFrameWL Frame1 
         Height          =   2655
         Left            =   3585
         TabIndex        =   3
         Top             =   315
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   4683
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmSnapshot.frx":038A
         Enabled         =   -1  'True
         ForeColor       =   -2147483642
         BackColor       =   -2147483633
         Tip             =   "frmSnapshot.frx":03B6
         VistaStyle      =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmSnapshot.frx":03D6
         RightToLeft     =   0   'False
         Begin SHDocVwCtl.WebBrowser webBrowser 
            Height          =   2625
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   6960
            ExtentX         =   12277
            ExtentY         =   4630
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   1
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   ""
         End
         Begin HexUniControls.ctlUniLabelXP lblLoading 
            Height          =   360
            Left            =   0
            Top             =   0
            Visible         =   0   'False
            Width           =   2700
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
            Caption         =   "frmSnapshot.frx":03F2
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            Alignment       =   0
            VAlignment      =   0
            BackStyle       =   1
            BorderStyle     =   0
            AutoSize        =   0   'False
            Tip             =   "frmSnapshot.frx":044A
            Style           =   0
            Enabled         =   -1  'True
            Margin          =   0
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmSnapshot.frx":046A
            RightToLeft     =   0   'False
            WordWrap        =   0   'False
         End
      End
      Begin VSFlex7LCtl.VSFlexGrid fgData 
         Height          =   2655
         Left            =   30
         TabIndex        =   1
         Top             =   315
         Width           =   3540
         _cx             =   6244
         _cy             =   4683
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
   Begin ActiveToolBars.SSActiveToolBars tbToolbar 
      Left            =   60
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131083
      ToolBarsCount   =   1
      ToolsCount      =   4
      Tools           =   "frmSnapshot.frx":0486
      ToolBars        =   "frmSnapshot.frx":1EFC
   End
   Begin VSFlex7LCtl.VSFlexGrid fgData2 
      Height          =   990
      Left            =   2400
      TabIndex        =   2
      Top             =   3960
      Width           =   2820
      _cx             =   4974
      _cy             =   1746
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
   Begin VB.Menu mnuPref 
      Caption         =   "Preferences"
      Begin VB.Menu mnuEditFields 
         Caption         =   "Edit Fields to display"
      End
      Begin VB.Menu mnuCompare 
         Caption         =   "Add Comparison Symbol"
      End
      Begin VB.Menu mnuRemoveCompare 
         Caption         =   "Remove Comparison Symbol"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print Grid"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangeFont 
         Caption         =   "Change Font"
      End
   End
End
Attribute VB_Name = "frmSnapshot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmSnapshot.frm
'' Description: Shows the fundamental data for a given symbol
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 06/24/2013   DAJ         Timer Logging
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

' Column to custom extend
Private Const kExtendedCol = 0

' Columns in the grid
Private Enum eGDCols
    eGDCol_Name = 0
    eGDCol_ID
    eGDCol_Description
    eGDCol_Date
    eGDCol_Symbol
    eGDCol_SubSector
    eGDCol_Sector
    eGDCol_SP500
    eGDCol_Comparison
    eGDCol_NumCols
End Enum

' Fields into the Criteria Table
Private Enum eTblCols
    eTblCol_Show = 0
    eTblCol_Name
    eTblCol_Description
    eTblCol_CodedText
    eTblCol_NumDays
    eTblCol_ID
    eTblCol_IsWeekly
    eTblCol_IsBoolean
    eTblCol_PriceDisplay
    eTblCol_DecimalPlaces
    eTblCol_NumCols
End Enum

Private Enum eFormatMode
    eNumericFormat = 0
    eDateFormat = 1
    eBooleanFormat = 2
End Enum

Public Enum eGDSnapshotMode
    eGDSnapshotMode_Single = 0
    eGDSnapshotMode_ShowSiblings
End Enum

Private Enum eVSCurrTab
    eVSTab_Custom = 0
    eVSTab_Cash
    eVSTab_Balance
    eVSTab_Income
    eVSTab_General
End Enum

Private Type mPrivate
    WindowLink As New cWindowLink

    lSymbolID As Long                   ' Symbol ID of the current Symbol
    lPoolRec As Long                    ' Pool record for this Symbol ID
    lSectorID As Long                   ' Symbol ID of the current Sector
    lSubSectorID As Long                ' Symbol ID of the current SubSector
    lSP500ID As Long                    ' Symbol ID of the $SPX Index
    lCompareID As Long                  ' Symbol ID of the comparison symbol
    
    strSymbol As String                 ' Current Symbol
    strDesc As String                   ' Description for the current symbol
    strSecType As eSYM_SecType          ' Security Type for the current symbol
    strSector As String                 ' Sector for the current symbol
    strSectorDesc As String             ' Description for the sector
    strSubsector As String              ' SubSector for the current symbol
    strSubsectorDesc As String          ' Description for the subsector
    strComparison As String             ' Comparison symbol
    strComparisonDesc As String         ' Comparison description
    
    CriteriaTable As cGdTable           ' Table of Criteria information and values
    strDefaultString As String          ' Default fields string
    strFields As String                 ' Current fields from INI file string
    
    lPrevColWidth As Long               ' Used for Extend custom column
    astrCompare As cGdArray             ' Array of comparison symbols
    alSymbolIds As cGdArray             ' Array of symbol id's to calc values for
    
    Mode As eGDSnapshotMode             ' Mode to run the snapshot form in
    bDontProcess As Boolean             ' Whether or not to do the toolbar click
    strSaveSymbol As String             ' Symbol saved before a sort
    
    bSyncToActiveChart As Boolean
    
    bShowURL As Boolean
    bShowGeneralWeb As Boolean
    strBaseURL As String
    strLastGoodURL As String
    
    bCashPreloaded As Boolean
    bBalancePreloaded As Boolean
    bGeneralPreLoaded As Boolean
End Type
Private m As mPrivate

Public Property Get SymbolID() As Long
    SymbolID = m.lSymbolID
End Property

Public Property Let SymbolID(ByVal lSymbolID As Long)
On Error GoTo ErrSection:
    
    If lSymbolID <> m.lSymbolID Then
        m.lSymbolID = lSymbolID
        If DockState(Me) <> eHidden Then
            ShowMe lSymbolID
        End If
    End If
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmSnapshot.LetSymbolID", eGDRaiseError_Raise
End Property

Private Function TblCol(ByVal Col As eTblCols) As Long
    TblCol = Col
End Function

Private Function GDCol(ByVal Col As eGDCols) As Long
    GDCol = Col
End Function

Private Property Get TableNum(ByVal nField As eTblCols, ByVal lRecord As Long) As Double
    TableNum = m.CriteriaTable.Num(nField, lRecord)
End Property
Private Property Get TableStr(ByVal nField As eTblCols, ByVal lRecord As Long) As String
    TableStr = m.CriteriaTable.Item(nField, lRecord)
End Property
Private Property Let TableNum(ByVal nField As eTblCols, ByVal lRecord As Long, ByVal dValue As Double)
    m.CriteriaTable.Num(nField, lRecord) = dValue
End Property
Private Property Let TableStr(ByVal nField As eTblCols, ByVal lRecord As Long, ByVal strValue As String)
    m.CriteriaTable.Item(nField, lRecord) = strValue
End Property

Public Property Get SyncToActiveChart() As Boolean
    SyncToActiveChart = m.bSyncToActiveChart
End Property

Public Property Get WindowLink() As cWindowLink
    Set WindowLink = m.WindowLink
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Loads the fundamental data for the Symbol ID passed in
'' Inputs:      Symbol ID to show
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMe(Optional ByVal lSymbolID As Long = 0, Optional ByVal strOutFile$ = "", Optional ByVal strFields As String = "")
On Error GoTo ErrSection:

    Dim dValue As Double                ' Temporary value variable
    Dim lDate As Long                   ' Date of the value
    Dim lPos As Long                    ' Position of symbol in the array
    Dim lRow As Long
    Dim aLines As New cGdArray
    Dim SymInfo As vbSymbolInfo
    Dim lPoolRec As Long
    Dim lIndex As Long
    Dim aStrings As New cGdArray
    
    Dim i As Integer
        
    m.bSyncToActiveChart = False
    
    m.bCashPreloaded = False
    m.bBalancePreloaded = False
    m.bGeneralPreLoaded = False
    
    If DockState(Me) = eHidden Then
        'm.bSyncToActiveChart = True '(default)
    End If
       
    m.strBaseURL = FixURL(GetProvidedProperty("SnapshotWeb"))
       
    ' Set the symbol and symbol ID
    If lSymbolID = 0 Then
        lSymbolID = FindWindowLinkSymbolID(Me)
        If lSymbolID = 0 Then
            lSymbolID = m.lSymbolID
            If lSymbolID = 0 Then
                On Error Resume Next
                lSymbolID = ActiveForm.SymbolID
                On Error GoTo ErrSection:
                If lSymbolID = 0 Then
                    Set aStrings = frmSymbolSelector.ShowMe("", False, True)
                    If aStrings.Size > 0 Then
                        lSymbolID = GetSymbolID(aStrings(0))
                    End If
                End If
            End If
        End If
    End If
    m.lSymbolID = lSymbolID
    
    m.lPoolRec = g.SymbolPool.PoolRecForSymbolID(lSymbolID)
    m.strSymbol = g.SymbolPool.SymbolForID(lSymbolID)
    m.strSecType = g.SymbolPool.SecType(m.lPoolRec)
    m.strDesc = g.SymbolPool.Desc(m.lPoolRec)
    Caption = "Snapshot of " & m.strSymbol & ":  " & g.SymbolPool.Desc(m.lPoolRec)
    
    ' Get sector/subsector information for symbol
    m.lSectorID = 0
    m.lSubSectorID = 0
    m.lCompareID = 0&
    m.strSector = ""
    m.strSubsector = ""
    If lSymbolID > 0 Then
        dValue = GetSectorID(lSymbolID, False)
        If dValue > 0 Then
            m.lSectorID = dValue
            m.strSector = SU_GetSymbol(dValue)
            If SU_GetSymbolInf(m.lSectorID, SymInfo) Then
                m.strSectorDesc = SymInfo.Description
            End If
        End If
        dValue = GetSectorID(lSymbolID, True)
        If dValue > 0 Then
            m.lSubSectorID = dValue
            m.strSubsector = SU_GetSymbol(dValue)
            If SU_GetSymbolInf(m.lSubSectorID, SymInfo) Then
                m.strSubsectorDesc = SymInfo.Description
            End If
        End If
    End If
    m.lSP500ID = SU_GetSymID("$SPX")
    
    Set m.astrCompare = New cGdArray
    m.astrCompare.Create eGDARRAY_Strings
    If m.astrCompare.FromFile(AddSlash(App.Path) & "Compare.SYM") Then
        m.astrCompare.BinarySearch m.strSymbol, lPos
        If Parse(m.astrCompare(lPos), vbTab, 1) = m.strSymbol Then
            m.lCompareID = CLng(Val(Parse(m.astrCompare(lPos), vbTab, 2)))
            lPoolRec = g.SymbolPool.PoolRecForSymbolID(m.lCompareID)
            m.strComparison = g.SymbolPool.Symbol(lPoolRec)
            m.strComparisonDesc = g.SymbolPool.Desc(lPoolRec)
        End If
    End If
        
    If Len(strFields) = 0 Then
        If m.strSecType = eSYMType_Stock Or Left(m.strSymbol, 2) = "$-" Then
            m.strDefaultString = "CLOSE.SCN;52WKHIGH.SCN;52WKLOW.SCN;AVGVOL.SCN;BALSHT.SCN;CMSHARES.SCN;MKTCAP.SCN;PERATIO.SCN;EPS.SCN;EPS12MO.SCN;ESTEPS.SCN;DIV.SCN;DIVYIELD.SCN;ANNDIV.SCN;BETA.SCN;PCTHELD.SCN;EARNINGS.SCN;ASSETS.SCN;LIABIL.SCN;SALES.SCN;EBIT.SCN;CURRAT.SCN;PRTOBK.SCN;PRTOREV.SCN;RASSETS.SCN;RETEQ.SCN;LTDEBT.SCN;SHTRMDT.SCN;DEBTPEQ.SCN"
            m.strFields = GetIniFileProperty("FieldsS", m.strDefaultString, "Fundamentals", AddSlash(App.Path) & "ChartNavigator.INI")
        ElseIf m.strSecType = eSYMType_Future Then
            m.strDefaultString = "CLOSE.SCN;52WKHIGH.SCN;52WKLOW.SCN;COMLONG.SCN;COMSHORT.SCN;LGSPCL.SCN;LGSPCS.SCN;SMSPCL.SCN;SMSPCS.SCN;MKTSENT.SCN;LWFSENT.SCN"
            m.strFields = GetIniFileProperty("FieldsF", m.strDefaultString, "Fundamentals", AddSlash(App.Path) & "ChartNavigator.INI")
        Else
            m.strDefaultString = "CLOSE.SCN;52WKHIGH.SCN;52WKLOW.SCN;CALHIGH.SCN;CALLOW.SCN"
            m.strFields = GetIniFileProperty("Fields", m.strDefaultString, "Fundamentals", AddSlash(App.Path) & "ChartNavigator.INI")
        End If
    Else
        m.strFields = strFields
    End If
    
    Set m.alSymbolIds = New cGdArray
    m.alSymbolIds.Create eGDARRAY_Longs
        
    ' Get the mode...
    'If m.strSecType = eSYMType_Stock Or Left(m.strSymbol, 2) = "$-" Then
    '    m.Mode = GetIniFileProperty("Mode", eGDSnapshotMode_Single, "Fundamentals", g.strIniFile)
    'Else
        m.Mode = eGDSnapshotMode_Single
    'End If
    
    Screen.MousePointer = vbHourglass
    Select Case m.Mode
        Case eGDSnapshotMode_Single
            m.alSymbolIds(0) = m.lSymbolID
            m.alSymbolIds(1) = m.lSubSectorID
            m.alSymbolIds(2) = m.lSectorID
            m.alSymbolIds(3) = m.lSP500ID
            If m.lCompareID <> 0 Then m.alSymbolIds(4) = m.lCompareID
            
            fgData.Redraw = flexRDNone
            InitGridSingle fgData
            LoadCriteriaList
            CalcValues
            LoadGridSingle fgData
            fgData.Redraw = flexRDBuffered
            
            m.bDontProcess = True
            tbToolbar.ToolBars("General").Tools("ID_ShowSiblings").State = ssUnchecked
            m.bDontProcess = False
            
            fgData2.Redraw = flexRDNone
            InitGridSingle fgData2
            fgData2.Redraw = flexRDBuffered
            
        Case eGDSnapshotMode_ShowSiblings
            SU_GetGroupSiblings m.lSymbolID, m.alSymbolIds
            If m.strSecType = eSYMType_Stock Then
                m.alSymbolIds.Add m.lSubSectorID, 0
                m.alSymbolIds.Add m.lSectorID, 1
                m.alSymbolIds.Add m.lSP500ID, 2
            ElseIf Left(m.strSymbol, 3) = "$--" Then
                m.alSymbolIds.Add m.lSP500ID, 0
            ElseIf Left(m.strSymbol, 2) = "$-" Then
                m.alSymbolIds.Add m.lSectorID, 0
                m.alSymbolIds.Add m.lSP500ID, 1
            End If
        
            fgData.Redraw = flexRDNone
            InitGridComparison fgData
            LoadCriteriaList
            CalcValues
            LoadGridComparison fgData
            fgData.Redraw = flexRDBuffered
   
            m.bDontProcess = True
            tbToolbar.ToolBars("General").Tools("ID_ShowSiblings").State = ssChecked
            m.bDontProcess = False
            
            fgData2.Redraw = flexRDNone
            InitGridComparison fgData2
            fgData2.Redraw = flexRDBuffered
    
    End Select
    Screen.MousePointer = vbDefault
    
    EnableControls
    
    If m.strSecType = eSYMType_Stock Then
        Dim iLastTab As Long
        
        fgData2.Visible = False
        vsTab.Visible = True
        
        If m.bShowGeneralWeb Then
            vsTab.TabVisible(eVSTab_General) = True
            vsTab.TabCaption(eVSTab_Custom) = "Custom"
            iLastTab = eVSTab_General
        Else
            vsTab.TabVisible(eVSTab_General) = False
            vsTab.TabCaption(eVSTab_Custom) = "General"
            iLastTab = eVSTab_Income
        End If
        
        vsTab.CurrTab = 0

'JM 08-22-2011: This code essentially forces the current tab to be zero (not sure why done this way...)
'   Code replaced with one-liner above. Remove after awhile if all ok.
'        If DockState(Me) = eHidden Then
'            vsTab.CurrTab = 0
'        ElseIf vsTab.CurrTab > 0 And vsTab.CurrTab < vsTab.NumTabs Then
'            vsTab_Switch 0, vsTab.CurrTab, i
'        Else
'            vsTab.CurrTab = 0
'        End If
    Else
        vsTab.Visible = False
        fgData.SaveGrid g.strAppPath & "\TempGrid.txt", flexFileAll
        
        With fgData2
            .Redraw = flexRDNone
            .LoadGrid g.strAppPath & "\TempGrid.txt", flexFileAll
            .Visible = True
            .Redraw = flexRDBuffered
        End With
        
        KillFile g.strAppPath & "\TempGrid.txt", True
    End If
    
    ' Show the form (or write to file)
    If Len(strOutFile) > 0 Then
        With fgData
            aLines.Clear
            lPos = g.SymbolPool.PoolRecForSymbolID(m.lSymbolID)
            aLines.Add "FundSymbol=" & g.SymbolPool.Symbol(lPos) & ";" & g.SymbolPool.Desc(lPos)
            lPos = g.SymbolPool.PoolRecForSymbolID(m.lSubSectorID)
            aLines.Add "Subsector=" & g.SymbolPool.Symbol(lPos) & ";" & g.SymbolPool.Desc(lPos)
            lPos = g.SymbolPool.PoolRecForSymbolID(m.lSectorID)
            aLines.Add "Sector=" & g.SymbolPool.Symbol(lPos) & ";" & g.SymbolPool.Desc(lPos)
            For lRow = .FixedRows To .Rows - 1
                aLines.Add Trim(.TextMatrix(lRow, GDCol(eGDCol_ID))) & "=" _
                    & Trim(.TextMatrix(lRow, GDCol(eGDCol_Name))) & ";" _
                    & Trim(.TextMatrix(lRow, GDCol(eGDCol_Description))) & ";" _
                    & Trim(.TextMatrix(lRow, GDCol(eGDCol_Symbol))) & ";" _
                    & Trim(.TextMatrix(lRow, GDCol(eGDCol_SubSector))) & ";" _
                    & Trim(.TextMatrix(lRow, GDCol(eGDCol_Sector))) & ";" _
                    & Trim(.TextMatrix(lRow, GDCol(eGDCol_SP500)))
            Next
            aLines.ToFile strOutFile
        End With
    ElseIf DockState(Me) = eHidden Then
        DockState(Me) = eShowAsPrevious
        tmr.Tag = "PreloadCash"      'JM 06-22-2010: Glen wants web preloaded behind the scene
        tmr.Enabled = True
    Else
        tmr.Tag = ""
        tmr.Enabled = True
    End If
    
    Form_Resize ' TLB 3/5/2010: in order to be sized correctly first time (esp. after being hidden)
    
    frmMain.SetWindowLink Me
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.ShowMe", eGDRaiseError_Raise
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGridSingle
'' Description: Loads the grid with data from table in "Single Symbol" mode
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGridSingle(fgGrid As VSFlexGrid)
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid redraw
    Dim lIndex As Long                  ' Index into a for loop
    Dim lRow As Long                    ' Row of the grid to put the data into
    Dim eFormat As eFormatMode
    
    
    If fgGrid Is Nothing Then Exit Sub
    
    With fgGrid
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        ' Set column labels for symbol, sector and subsector
        .TextMatrix(0, GDCol(eGDCol_Symbol)) = m.strSymbol
        If m.lSectorID <> kNullData Then
            .TextMatrix(0, GDCol(eGDCol_Sector)) = m.strSector & vbCrLf & "(Sector)"
        End If
        If m.lSubSectorID <> kNullData Then
            .TextMatrix(0, GDCol(eGDCol_SubSector)) = m.strSubsector & vbCrLf & "(Subsector)"
        End If
        If m.lCompareID <> 0 Then
            .ColHidden(GDCol(eGDCol_Comparison)) = False
            .TextMatrix(0, GDCol(eGDCol_Comparison)) = g.SymbolPool.SymbolForID(m.lCompareID)
        Else
            .ColHidden(GDCol(eGDCol_Comparison)) = True
        End If
                
        ' Set row data
        .Rows = fgGrid.FixedRows
        For lIndex = 0 To m.CriteriaTable.NumRecords - 1
            lRow = TableNum(eTblCol_Show, lIndex)
            If lRow > 0 Then
                If lRow + 1 > .Rows Then .Rows = lRow + 1
                .TextMatrix(lRow, GDCol(eGDCol_ID)) = TableStr(eTblCol_ID, lIndex)
                .TextMatrix(lRow, GDCol(eGDCol_Name)) = " " & TableStr(eTblCol_Name, lIndex)
                .TextMatrix(lRow, GDCol(eGDCol_Description)) = TableStr(eTblCol_Description, lIndex)
                If TableNum(eTblCol_IsBoolean, lIndex) <> 0 Then
                    eFormat = eBooleanFormat
                ElseIf InStr(UCase(TableStr(eTblCol_Name, lIndex) & " " _
                        & TableStr(eTblCol_Description, lIndex)), "DATE") > 0 Then
                    eFormat = eDateFormat
                Else
                    eFormat = eNumericFormat
                End If
                .TextMatrix(lRow, GDCol(eGDCol_Date)) = ""
                .TextMatrix(lRow, GDCol(eGDCol_Symbol)) = FormatValue(TableNum(eTblCol_NumCols, lIndex), eFormat, TableNum(eTblCol_PriceDisplay, lIndex), TableNum(eTblCol_DecimalPlaces, lIndex), m.strSymbol)
                .TextMatrix(lRow, GDCol(eGDCol_SubSector)) = FormatValue(TableNum(eTblCol_NumCols + 1, lIndex), eFormat, TableNum(eTblCol_PriceDisplay, lIndex), TableNum(eTblCol_DecimalPlaces, lIndex), m.strSubsector)
                .TextMatrix(lRow, GDCol(eGDCol_Sector)) = FormatValue(TableNum(eTblCol_NumCols + 2, lIndex), eFormat, TableNum(eTblCol_PriceDisplay, lIndex), TableNum(eTblCol_DecimalPlaces, lIndex), m.strSector)
                .TextMatrix(lRow, GDCol(eGDCol_SP500)) = FormatValue(TableNum(eTblCol_NumCols + 3, lIndex), eFormat, TableNum(eTblCol_PriceDisplay, lIndex), TableNum(eTblCol_DecimalPlaces, lIndex), "$SPX")
                .TextMatrix(lRow, GDCol(eGDCol_Comparison)) = FormatValue(TableNum(eTblCol_NumCols + 4, lIndex), eFormat, TableNum(eTblCol_PriceDisplay, lIndex), TableNum(eTblCol_DecimalPlaces, lIndex), m.strComparison)
            End If
        Next
        
        .AutoSize 0, .Cols - 1, False, 75
        ExtendCustomColumn fgGrid
        
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.LoadGridSingle", eGDRaiseError_Raise
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGridComparison
'' Description: Loads the grid with data from table in "Comparison" mode
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGridComparison(fgGrid As VSFlexGrid)
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid redraw
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim lCol As Long                    ' Column of the grid to put the data into
    Dim eFormat As eFormatMode          ' Format mode for the data
    Dim lPoolRec As Long                ' Pool record for symbol id
    
    
    If fgGrid Is Nothing Then Exit Sub
    
    With fgGrid
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        ' Set row data
        .Cols = 1
        For lIndex = 0 To m.CriteriaTable.NumRecords - 1
            lCol = TableNum(eTblCol_Show, lIndex)
            If lCol > 0 Then
                If lCol + 1 > .Cols Then .Cols = lCol + 1
                .TextMatrix(GDCol(eGDCol_ID), lCol) = TableStr(eTblCol_ID, lIndex)
                .TextMatrix(GDCol(eGDCol_Name), lCol) = " " & TableStr(eTblCol_Name, lIndex)
                .TextMatrix(GDCol(eGDCol_Description), lCol) = TableStr(eTblCol_Description, lIndex)
                If TableNum(eTblCol_IsBoolean, lIndex) <> 0 Then
                    eFormat = eBooleanFormat
                    .ColAlignment(lCol) = flexAlignRightTop
                ElseIf InStr(UCase(TableStr(eTblCol_Name, lIndex) & " " & TableStr(eTblCol_Description, lIndex)), "DATE") > 0 Then
                    eFormat = eDateFormat
                    .ColAlignment(lCol) = flexAlignCenterTop
                Else
                    eFormat = eNumericFormat
                    .ColAlignment(lCol) = flexAlignRightTop
                End If
                .TextMatrix(GDCol(eGDCol_Date), lCol) = ""
                
                For lIndex2 = 0 To m.alSymbolIds.Size - 1
                    lPoolRec = g.SymbolPool.PoolRecForSymbolID(m.alSymbolIds(lIndex2))
                    .RowData(GDCol(eGDCol_NumCols) + lIndex2) = g.SymbolPool.Desc(lPoolRec)
                    .TextMatrix(GDCol(eGDCol_NumCols) + lIndex2, 0) = g.SymbolPool.Symbol(lPoolRec)
                    .TextMatrix(GDCol(eGDCol_NumCols) + lIndex2, lCol) = FormatValue(TableNum(eTblCol_NumCols + lIndex2, lIndex), eFormat, TableNum(eTblCol_PriceDisplay, lIndex), TableNum(eTblCol_DecimalPlaces, lIndex), g.SymbolPool.Symbol(lPoolRec))
                Next lIndex2
            End If
        Next
        
        If .Rows > .FixedRows Then
            .Select .FixedRows, 0, .Rows - 1, 0
            .Sort = flexSortStringAscending
            HighlightSymbol m.strSymbol
        End If
        
        SetBackColors fgGrid
        .AutoSize 0, .Cols - 1, False, 75
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        ExtendCustomColumn fgGrid
        
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.LoadGridComparison", eGDRaiseError_Raise
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FormatValue
'' Description: Format the value into a string depending on the field
'' Inputs:      Value, Is it a Date?
'' Returns:     String representation of the value
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FormatValue(ByVal dValue As Double, Optional ByVal eFormatMode As eFormatMode = eNumericFormat, Optional ByVal nPriceDisplay As eCriteriaPriceDisplay = eCriteria_AutoRound, Optional ByVal lDecimals As Long = 2, Optional ByVal strSymbol As String = "") As String
On Error GoTo ErrSection:

    Dim strValue As String              ' String representation of the value
    
    Select Case eFormatMode
        Case eBooleanFormat
            If dValue = 0 Then
                strValue = "False"
            ElseIf Abs(dValue) = 1 Then
                strValue = "TRUE"
            End If
    
        Case eDateFormat
            If dValue > 0 Then
                strValue = DateFormat(dValue)
            End If
            
        Case Else
            If dValue <> kNullData Then
                'If m.strSecType = eSYMType_Stock Then
                '    strValue = Format(dValue, "#,##0.##")
                'Else
                '    strValue = Format(dValue, "#,##0.#####")
                'End If
                'If Right(strValue, 1) = "." Then
                '    strValue = Left(strValue, Len(strValue) - 1)
                'End If
                Select Case nPriceDisplay
                    Case eCriteria_AutoRound
                        If dValue = 0 Then
                            strValue = "0"
                        ElseIf Abs(dValue) >= 100000 Then
                            strValue = Format(dValue, "#,##0")
                        ElseIf Abs(dValue) > 10000 Or Int(dValue) = dValue Then
                            strValue = Format(dValue, "0")
                        ElseIf Abs(dValue) < 10 Then
                            strValue = Format(dValue, "0.####")
                        Else
                            strValue = Format(dValue, "0.##")
                        End If
                    
                    Case eCriteria_RoundToDecimal
                        strValue = Trim(NumStr(dValue, 0, lDecimals))
                    
                    Case eCriteria_TradingUnits
                        strValue = PriceDisplay(dValue, strSymbol, True)
                        
                End Select
            End If
    End Select

    FormatValue = strValue

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSnapshot.FormatValue", eGDRaiseError_Raise
        
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EditFieldsToShow
'' Description: Allows the user to select which fields to show in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EditFieldsToShow(fgGrid As VSFlexGrid)
On Error GoTo ErrSection:
    
    Dim aAvail As New cGdArray          ' String array of available kinds
    Dim aInUse As New cGdArray          ' String array of currently shown kinds
    Dim aDefaults As New cGdArray       ' String array of the default kinds
    Dim lIndex As Long                  ' Index into a for loop
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim strTemp As String               ' Temporary string of selected fields
    
    If fgGrid Is Nothing Then Exit Sub
        
    ' Create the arrays
    aAvail.Create eGDARRAY_Strings
    aInUse.Create eGDARRAY_Strings
    aDefaults.Create eGDARRAY_Strings
        
    ' Load up the Available, InUse, and Default arrays
    aDefaults.SplitFields m.strDefaultString, ";"
    For lIndex = 0 To m.CriteriaTable.NumRecords - 1
        If TableNum(eTblCol_Show, lIndex) > 0 Then
            aInUse(TableNum(eTblCol_Show, lIndex) - 1) = TableStr(eTblCol_Name, lIndex)
        Else
            aAvail.Add TableStr(eTblCol_Name, lIndex)
        End If
        For lIndex2 = 0 To aDefaults.Size - 1
            If TableStr(eTblCol_ID, lIndex) = aDefaults(lIndex2) Then
                aDefaults(lIndex2) = TableStr(eTblCol_Name, lIndex)
                Exit For
            End If
        Next lIndex2
    Next lIndex
    aAvail.Sort
    
    '  Let user edit the fields list
    If frmAddRemove.ShowMe(aAvail, aInUse, eOrderMode_Ordered, aDefaults, "Arrange Criteria to Display") Then
        For lIndex = 0 To m.CriteriaTable.NumRecords - 1
            TableNum(eTblCol_Show, lIndex) = 0
        Next lIndex
        If aInUse.Size = 0 Then
            Set aInUse = aDefaults.MakeCopy
        End If
        strTemp = ""
        For lIndex = 0 To aInUse.Size - 1
            For lIndex2 = 0 To m.CriteriaTable.NumRecords - 1
                If TableStr(eTblCol_Name, lIndex2) = aInUse(lIndex) Then
                    TableNum(eTblCol_Show, lIndex2) = lIndex + 1
                    strTemp = strTemp & ";" & TableStr(eTblCol_ID, lIndex2)
                    Exit For
                End If
            Next lIndex2
        Next lIndex
        If Len(strTemp) = 0 Then strTemp = ";"
        If m.strSecType = eSYMType_Stock Or Left(m.strSymbol, 2) = "$-" Then
            SetIniFileProperty "FieldsS", Right(strTemp, Len(strTemp) - 1), "Fundamentals", g.strIniFile
        ElseIf m.strSecType = eSYMType_Future Then
            SetIniFileProperty "FieldsF", Right(strTemp, Len(strTemp) - 1), "Fundamentals", g.strIniFile
        Else
            SetIniFileProperty "Fields", Right(strTemp, Len(strTemp) - 1), "Fundamentals", g.strIniFile
        End If
        
        Screen.MousePointer = vbHourglass
        fgGrid.Redraw = flexRDNone
        Select Case m.Mode
            Case eGDSnapshotMode_Single
                InitGridSingle fgGrid
                CalcValues
                LoadGridSingle fgGrid
            
            Case eGDSnapshotMode_ShowSiblings
                InitGridComparison fgGrid
                CalcValues
                LoadGridComparison fgGrid
                
        End Select
        fgGrid.Redraw = flexRDBuffered
        Screen.MousePointer = vbDefault
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.EditFieldsToShow", eGDRaiseError_Raise
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgData_AfterSort
'' Description: After sorting, make sure that the alternate coloring is correct
'' Inputs:      Column Sorted, Order Sorted
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgData_AfterSort(ByVal Col As Long, Order As Integer)
On Error GoTo ErrSection:
    

    GridAfterSort fgData, Col, Order

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.fgData.AfterSort", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgData_AfterUserResize
'' Description: Make sure that our Custom Extend Column is extended appropriately
'' Inputs:      Row and Column resized
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgData_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:


    GridAfterUserResize fgData, Row, Col

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.fgData.AfterUserResize", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgData_BeforeSort
'' Description: Make sure to do a string sort on the symbol column if in
''              ShowSiblings mode
'' Inputs:      Column to sort, Order to sort the column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgData_BeforeSort(ByVal Col As Long, Order As Integer)
On Error GoTo ErrSection:

    GridBeforeSort fgData, Col, Order
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.fgData.BeforeEdit", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgData_BeforeUserResize
'' Description: Make sure that our Custom Extend Column is extended appropriately
'' Inputs:      Row and Column resized, Whether or not to Cancel the Resize
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgData_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    GridBeforeUserResize fgData, Row, Col, Cancel

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.fgData.BeforeUserResize", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgData_DblClick
'' Description: If the user double clicks in the grid, bring up the add/remove
''              form so that they can edit the fields list
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgData_DblClick()
On Error GoTo ErrSection:

    GridDblClick fgData

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.fgData.DblClick", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgData_KeyDown
'' Description: Send any key strokes to the fgKeyDown routine
'' Inputs:      Code of the Key Pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgData_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If fgKeyDown(KeyCode, Shift) Then Exit Sub

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.fgData.KeyDown", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgData_MouseDown
'' Description: If the user right clicks in the grid, bring up the popup
'' Inputs:      Button Pressed, Shift/Ctrl/Alt status, Mouse Location
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgData_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    GridMouseDown fgData, Button, Shift, X, Y

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.fgData.MouseDown", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgData_MouseMove
'' Description: As the user moves the mouse, show the description of the
''              current mouse row as the tool tip text
'' Inputs:      Button Pressed, Shift/Ctrl/Alt status, Mouse Location
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgData_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    GridMouseMove fgData, Button, Shift, X, Y

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.fgData.MouseMove", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

Private Sub fgData2_AfterSort(ByVal Col As Long, Order As Integer)
    GridAfterSort fgData2, Col, Order
End Sub

Private Sub fgData2_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    GridAfterUserResize fgData2, Row, Col
End Sub

Private Sub fgData2_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    GridBeforeUserResize fgData2, Row, Col, Cancel
End Sub

Private Sub fgData2_DblClick()
    GridDblClick fgData2
End Sub

Private Sub fgData2_KeyDown(KeyCode As Integer, Shift As Integer)
    If fgKeyDown(KeyCode, Shift) Then Exit Sub
End Sub

Private Sub fgData2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GridMouseDown fgData2, Button, Shift, X, Y
End Sub

Private Sub fgData2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GridMouseMove fgData2, Button, Shift, X, Y
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: When the form gets activated, reset the toolbar and the window
''              list
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:

    m.WindowLink.Init Me

    ToolbarSync Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.Form.Activate", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Deactivate
'' Description: When the form is deactivated set it as the previous form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Deactivate()
On Error GoTo ErrSection:

    SetPrevActiveForm Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.Form.Deactivate", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_KeyDown
'' Description: If the user presses a function key, pass it on to KeyPress
'' Inputs:      Code of the Key pressed, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyF1 Then
        g.Help.ShowF1Help Me
    ElseIf KeyCode >= vbKeyF2 And KeyCode <= vbKeyF12 Then
        KeyPress KeyCode, Shift
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_KeyPress
'' Description: If the user presses a key, pass it on to the global KeyPress
'' Inputs:      Ascii version of the Key pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo ErrSection:

    KeyPress KeyAscii

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.Form.KeyPress", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: When the form gets loaded, size it and center it
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Me.Icon = Picture16(ToolbarIcon("ID_Snapshot"), , True)
    CenterTheForm Me
    
    g.Styler.StyleForm Me
    
    mnuPref.Visible = False
'    m.bSyncToActiveChart = True '(default)
    
    ''m.WindowLink.SymbolColor = GetIniFileProperty(Me.Name, 0&, "SymbolLink", g.strIniFile)
    
    With tbToolbar
        .Tools("ID_ChangeSymbol").Picture = Picture16(ToolbarIcon("ID_Symbol"))
        .Tools("ID_Print").Picture = Picture16(ToolbarIcon("ID_Print"))
        .Tools("ID_EditFields").Picture = Picture16(ToolbarIcon("ID_Settings"))
        .Tools("ID_ShowSiblings").PictureDown = Picture16(ToolbarIcon("kCheckOn"))
        .Tools("ID_ShowSiblings").Picture = Picture16(ToolbarIcon("kCheckOff"))
    End With
    
    vsTab.CurrTab = 0
    vsTab.BoldCurrent = True
    vsTab.FrontTabForeColor = &H8000000D        'flag this to be set to blue for classic & light color scheme
    
    m.bShowURL = FileExist(kShowUrlFlagFile)
    m.bShowGeneralWeb = FileExist(g.strAppPath & "\Provided\SnapshotGenWeb.flg")
   
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.Form.Load", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user clicks on the X, reset the toolbar
'' Inputs:      Whether or not to cancel the unload, Mode of the unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode = 0 Then
        frmMain.tbToolbar.Tools("ID_Snapshot").State = ssUnchecked
    End If
    
    If Cancel = 0 Then m.WindowLink.Unhook
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: When the form is resized, resize the grid as well
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next
    
    Static nWidth&, nHeight&
    
'    If nWidth = Me.Width And nHeight = Me.Height Then Exit Sub      '5469
    
    With fgData2
        .Redraw = flexRDNone
        .Move ScaleLeft, ScaleTop, ScaleWidth, ScaleHeight
        ExtendCustomColumn fgData2
        .Redraw = flexRDBuffered
    End With

    With vsTab
        .Move ScaleLeft, ScaleTop, ScaleWidth, ScaleHeight
        .Refresh
    End With

    With fgData
        .Redraw = flexRDNone
        .Move vsTab.ClientLeft, vsTab.ClientTop, vsTab.ClientWidth - 30, vsTab.ClientHeight
        ExtendCustomColumn fgData
        .Redraw = flexRDBuffered
    End With

    With webBrowser
        .Move vsTab.Left, vsTab.Top, vsTab.Width - 75, ScaleHeight - 350        '5422
    End With
    
    nWidth = Me.Width
    nHeight = Me.Height
        
    AutoSizeChart

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: When the form gets unloaded, reset the window list
'' Inputs:      Whether or not to cancel the update
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    ' store window link color and unhook the window proc
    ''SetIniFileProperty Me.Name, m.WindowLink.SymbolColor, "SymbolLink", g.strIniFile
    Set m.WindowLink = Nothing

    m.lSymbolID = 0
    ToolbarSync Me, False
       
    frmMain.DockPro.RemoveForm Me.Name
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.Form.Unload", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenerateReport
'' Description: Set up the print preview form for this grid
'' Inputs:      Arguments passed in from PrintMe
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
        .TextAlign = taCenterMiddle
        .Text = Me.Caption
        .TextAlign = taLeftMiddle
        .Font.Bold = False
        
        .Paragraph = ""
        .Paragraph = ""
        
        If frmPrintPreview.GoingToFile Then
            frmPrintPreview.GridToTable fgData
        Else
            fgData.AutoSize kExtendedCol, , , 450           '5284
            .RenderControl = fgData.hWnd
        End If
        
        .EndDoc
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.GenerateReport", eGDRaiseError_Raise
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PrintMe
'' Description: Bring up the print preview for this form
'' Inputs:      None
'' Returns:     True on OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PrintMe() As Boolean
On Error GoTo ErrSection

    If vsTab.Visible And vsTab.CurrTab <> 0 Then
        webBrowser.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DODEFAULT        '5470
    Else
        PrintMe = frmPrintPreview.ShowMe("CNV Fundamental", Me, , , , 0.75, 0.75)
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSnapshot.PrintMe", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    KeyPress
'' Description: Perform an action based on the key the user pressed
'' Inputs:      Key Pressed, Shift/Ctrl/Alt Status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub KeyPress(KeyAscii As Integer, Optional Shift As Integer = -1)
On Error Resume Next

    Dim frm As Form                     ' Charting Form
    Dim bLookForChart As Boolean        ' Should we look for the chart?

    If KeyAscii = 0 Then Exit Sub

    If Shift >= 0 Then ' (came from KeyDown event)
        If KeyAscii >= vbKeyF2 And KeyAscii <= vbKeyF12 Then
            bLookForChart = True
        End If
    Else ' (came from KeyPress event)
        Select Case Asc(UCase(Chr(KeyAscii)))
            
            Case 83:        ' S
                ChangeSymbol
                KeyAscii = 0
                
            Case 65 To 90, 48 To 57, 43, 45, 61:
                bLookForChart = True
        End Select
    End If
       
    If bLookForChart Then
        Set frm = ActiveChart
        If Not frm Is Nothing Then
            frm.KeyPress KeyAscii, Shift
        End If
        KeyAscii = 0
    End If
       
    Set frm = Nothing

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuChangeFont_Click
'' Description: Change the font of the quotes grid if the user chooses to
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuChangeFont_Click()
On Error GoTo ErrSection:

    Dim fgGrid As VSFlexGrid

    If fgData.Visible Then
        Set fgGrid = fgData
    ElseIf fgData2.Visible Then
        Set fgGrid = fgData2
    End If
    
    If Not fgGrid Is Nothing Then
        If ChangeGridFont(fgGrid, True) Then
            SetIniFileProperty "Fundamental", FontToString(fgGrid.Font), "Fonts", g.strIniFile
            FormResize Me
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.mnuChangeFont.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExtendCustomColumn
'' Description: Adjust all column widths to accomodate the custom "extend column"
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ExtendCustomColumn(fgGrid As VSFlexGrid)
On Error GoTo ErrSection:

    Dim lTotal As Long                  ' New width of the extended column
    Dim lIndex As Long                  ' Index into a for loop

    With fgGrid
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
    RaiseError "frmSnapshot.ExtendCustomColumn", eGDRaiseError_Raise
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuCompare_Click
'' Description: Allow the user to add/change a comparison symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuCompare_Click()
On Error GoTo ErrSection:

    If fgData.Visible Then
        AddComparison fgData
    ElseIf fgData2.Visible Then
        AddComparison fgData2
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.mnuCompare.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuEditFields_Click
'' Description: If the user clicks on the Edit Fields menu item, show them the
''              add/remove form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuEditFields_Click()
On Error GoTo ErrSection:

    If fgData.Visible Then
        EditFieldsToShow fgData
    ElseIf fgData2.Visible Then
        EditFieldsToShow fgData2
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.mnuEditFields.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuPrint_Click
'' Description: When the user selects to print, bring up the print preview
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuPrint_Click()
On Error GoTo ErrSection:

    PrintMe

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.mnuPrint.Click", eGDRaiseError_Show
    Resume ErrExit
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalcValues
'' Description: Recalculate values for all symbols and expressions
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CalcValues()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim hArray As Long                  ' Array handle
    Dim lFromDate As Long               ' Date to load data from
    Dim lLastDateOfData As Long         ' Last date of data
    Dim rc As Long                      ' Return code from function calls
    Dim lSymbol As Long                 ' Index into a for loop
    Dim strSymbol As String             ' Symbol to get data for
    Dim lSymbolID As Long               ' Symbol ID to get data for
    Dim dPrice As Double                ' Price from the results array
    Dim dPrev As Double                 ' Previous value
    Dim strCodedText As String          ' Coded text for an English expression
    Dim lNumDays As Long                ' Number of days required to run expr
            
    Dim Bars As New cGdBars             ' Data for the main market
    Dim GC As New cGdBars               ' Data for Gold 67 contract
    Dim Weekly As New cGdBars           ' Data for weekly bars of main market
    Dim GCWeekly As New cGdBars         ' Weekly data for Gold 67 contract
    
    Dim astrParms As New cGdArray       ' Paramaters array for the engine
    Dim astrBarNames As New cGdArray    ' Array of bar names
    Dim aScanExpr As New cGdArray       ' Array of coded text expressions
    Dim aArrayOfResults As New cGdArray ' Array of results
    Dim aArrayOfBars As New cGdArray    ' Array of bars structures
    Dim aTableIndex As New cGdArray     ' Index into the Criteria Table
    
    Dim astrParmsW As New cGdArray      ' Paramaters array for the engine
    'Dim astrBarNamesW As New cGdArray   ' Array of bar names
    Dim aScanExprW As New cGdArray      ' Array of coded text expressions
    Dim aArrayOfResultsW As New cGdArray ' Array of results
    Dim aArrayOfBarsW As New cGdArray   ' Array of bars structures
    Dim aTableIndexW As New cGdArray    ' Index into the Criteria Table
    
    Dim adTemp As New cGdArray          ' Temporary array
    
    Dim SecondaryMarkets As New cGdTree ' Bars collection of secondary markets
    Dim lBars As Long                   ' Index into a for loop
    Dim strError As String              ' Error back from setup expressions

    ' Create the arrays
    aScanExpr.Create eGDARRAY_Strings
    aArrayOfResults.Create eGDARRAY_Longs
    aScanExprW.Create eGDARRAY_Strings
    aArrayOfResultsW.Create eGDARRAY_Longs
    aTableIndex.Create eGDARRAY_Longs
    aTableIndexW.Create eGDARRAY_Longs
    
    ' Initialize the number of days
    lNumDays = 0&
    
    ' Set up the expressions, values, and results arrays
    For lIndex = 0 To m.CriteriaTable.NumRecords - 1
        If TableNum(eTblCol_Show, lIndex) > 0 Then
            strCodedText = TableStr(eTblCol_CodedText, lIndex)
            If Len(Trim(strCodedText)) > 0 Then
                hArray = gdCreateArray(eGDARRAY_Doubles, 0)
                
                If TableNum(eTblCol_IsWeekly, lIndex) = 0 Then
                    aScanExpr.Add Trim(strCodedText)
                    aArrayOfResults.Add hArray
                    aTableIndex(aArrayOfResults.Size - 1) = lIndex
                    
                    ' Calculate the maximum number of days needed to run
                    If TableNum(eTblCol_NumDays, lIndex) > lNumDays Then
                        lNumDays = TableNum(eTblCol_NumDays, lIndex)
                    End If
                Else
                    aScanExprW.Add Trim(strCodedText)
                    aArrayOfResultsW.Add hArray
                    aTableIndexW(aArrayOfResultsW.Size - 1) = lIndex
                
                    ' Calculate the maximum number of days needed to run
                    If (TableNum(eTblCol_NumDays, lIndex) + 1) * 5 > lNumDays Then
                        lNumDays = (TableNum(eTblCol_NumDays, lIndex) + 1) * 5
                    End If
                End If
            End If
        End If
    Next lIndex
    
    ' Calc FromDate, adjusting for weekends and holidays
    ' (need to fudge a little to the safe side)
    If lNumDays = 1 Then lNumDays = 5
    lLastDateOfData = LastDailyDownload
    lFromDate = lLastDateOfData - Int(lNumDays * 1.46 + 0.5) - 2
    
    MarketsInExpressions aScanExpr, lFromDate, False, astrBarNames, SecondaryMarkets, "Daily"
    MarketsInExpressions aScanExprW, lFromDate, False, astrBarNames, SecondaryMarkets, "Weekly"

    If aScanExpr.Size + aScanExprW.Size > 0 Then
        ' Init the expression evaluator with list of scan expressions
        'astrBarNames(0) = "Market1"
        'astrBarNames(1) = "Weekly"
        'astrBarNames(2) = "GC"
        astrParms(0) = "DailyFundCalc"
        If Not SetupExpressions(astrParms, astrBarNames, aScanExpr, strError) Then
            InfBox "i=[] ; h=CalcValues ; An error exists in an expression."
            DebugLog "DailyFundCalc: Error in expression = " & strError
            Exit Sub
        End If

        'astrBarNamesW(0) = "Market1"
        'astrBarNamesW(1) = "Weekly"
        'astrBarNamesW(2) = "GC"
        astrParmsW(0) = "WeeklyFundCalc"
        If Not SetupExpressions(astrParmsW, astrBarNames, aScanExprW, strError) Then
            InfBox "i=[] ; h=CalcValues ; An error exists in an expression."
            DebugLog "WeeklyFundCalc: Error in expression = " & strError
            Exit Sub
        End If
        
        aArrayOfBars.Create eGDARRAY_Longs
        aArrayOfBarsW.Create eGDARRAY_Longs
        
        ' Load Gold in case we need it
        'DM_GetBars GC, "GC-067", 0, lFromDate, , , , , False
        'GCWeekly.BuildBars "Weekly", GC.BarsHandle
        
        For lSymbol = 0 To m.alSymbolIds.Size - 1
            lSymbolID = m.alSymbolIds(lSymbol)
            strSymbol = g.SymbolPool.SymbolForID(lSymbolID)
            
            Bars.Size = 0
            SetBarProperties Bars, lSymbolID
            
            If lSymbolID <> 0 Then
                If Not DM_GetBars(Bars, lSymbolID, 0, lFromDate, , , , , False) Then
                    Bars.Size = 0
                End If
            End If

            If Bars.Size > 0 Then
                Weekly.BuildBars "Weekly", Bars.BarsHandle
                
                aArrayOfBars.Num(0) = Bars.BarsHandle '(in case changed)
                aArrayOfBars.Num(1) = Weekly.BarsHandle
                'aArrayOfBars.Num(2) = GC.BarsHandle
                For lBars = 2 To astrBarNames.Size - 1
                    aArrayOfBars.Num(lBars) = SecondaryMarkets(lBars + 1).BarsHandle
                Next lBars
                
                ' Run engine to evaluate expressions for this symbol
                astrParms.Size = 1
                rc = RunExpressions(astrParms.ArrayHandle, _
                    astrBarNames.ArrayHandle, aArrayOfBars.ArrayHandle, _
                    aArrayOfResults.ArrayHandle, ByVal 0&, ByVal 0&)
                If rc = 0 Then
                    '  Set current value for each expression
                    For lIndex = 0 To aArrayOfResults.Size - 1
                        ' Get most recent value
                        hArray = aArrayOfResults.Num(lIndex)
                        If gdIsConstantValue(hArray) Then
                            dPrice = gdGetNum(hArray, 0)
                        Else
                            dPrice = gdGetNum(hArray, Bars.Size - 1)
                        End If

                        ' Store into the scan's array for this symbol
                        TableNum(eTblCol_NumCols + lSymbol, aTableIndex(lIndex)) = dPrice
                    Next lIndex
                End If
                
                aArrayOfBarsW.Num(0) = Weekly.BarsHandle '(in case changed)
                aArrayOfBarsW.Num(1) = Weekly.BarsHandle
                'aArrayOfBarsW.Num(2) = GC.BarsHandle
                For lBars = 2 To astrBarNames.Size - 1
                    aArrayOfBarsW.Num(lBars) = SecondaryMarkets(lBars + 1).BarsHandle
                Next lBars
                
                ' Run engine to evaluate expressions for this symbol
                astrParmsW.Size = 1
                rc = RunExpressions(astrParmsW.ArrayHandle, _
                    astrBarNames.ArrayHandle, aArrayOfBarsW.ArrayHandle, _
                    aArrayOfResultsW.ArrayHandle, ByVal 0&, ByVal 0&)
                If rc = 0 Then
                    '  Set current value for each expression
                    For lIndex = 0 To aArrayOfResultsW.Size - 1
                        ' Get most recent value
                        hArray = aArrayOfResultsW.Num(lIndex)
                        If gdIsConstantValue(hArray) Then
                            dPrice = gdGetNum(hArray, 0)
                        ElseIf Weekly(eBARS_DateTime, Weekly.Size - 1) > LastDailyDownload Then
                            dPrice = gdGetNum(hArray, Weekly.Size - 2)
                        Else
                            dPrice = gdGetNum(hArray, Weekly.Size - 1)
                        End If

                        ' Store into the scan's array for this symbol
                        TableNum(eTblCol_NumCols + lSymbol, aTableIndexW(lIndex)) = dPrice
                    Next lIndex
                End If
            End If
            
            'we'll yield to other threads only every 1/2 second
            Sleep -0.5
        Next lSymbol
        
        ' clear the expression evaluator when done with it
        SetupExpressions astrParms '(clear expressions)
        SetupExpressions astrParmsW '(clear expressions)
    End If
    
   ' Destroy all the temporary result arrays
    For lIndex = 0 To aArrayOfResults.Size - 1
        gdDestroyArray aArrayOfResults(lIndex)
    Next lIndex
    aArrayOfResults.Size = 0
    For lIndex = 0 To aArrayOfResultsW.Size - 1
        gdDestroyArray aArrayOfResultsW(lIndex)
    Next lIndex
    aArrayOfResultsW.Size = 0
    
    Set adTemp = Nothing
            
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.CalcValues", eGDRaiseError_Raise
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadCriteriaList
'' Description: Loads up the Criteria Table with all non-boolean criteria from
''              the symbol pool
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadCriteriaList()
On Error GoTo ErrSection:

    Dim Criteria As New cCriteria       ' Local Criteria object
    Dim lIndex As Long                  ' Index into the table
    Dim astrFields As New cGdArray      ' Array of fields to show
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim lPoolField As Long              ' Pool field for the Symbol Group
    Dim strText As String

    ' Initialize the table
    Set m.CriteriaTable = New cGdTable
    With m.CriteriaTable
        .CreateField eGDARRAY_Longs, , "Show"
        .CreateField eGDARRAY_Strings, , "Name"
        .CreateField eGDARRAY_Strings, , "Desc"
        .CreateField eGDARRAY_Strings, , "Coded Text"
        .CreateField eGDARRAY_Longs, , "Num Days"
        .CreateField eGDARRAY_Strings, , "ID"
        .CreateField eGDARRAY_TinyInts, , "Is Weekly"
        .CreateField eGDARRAY_TinyInts, , "Is Boolean"
        .CreateField eGDARRAY_Longs, , "PriceDisplay"
        .CreateField eGDARRAY_Longs, , "DecimalPlaces"
        '.CreateField eGDARRAY_Doubles, , "Value"
        '.CreateField eGDARRAY_Doubles, , "Sector Value"
        '.CreateField eGDARRAY_Doubles, , "Subsector Value"
        '.CreateField eGDARRAY_Doubles, , "SPX Value"
        '.CreateField eGDARRAY_Doubles, , "Compare"
        For lIndex = 0 To m.alSymbolIds.Size - 1
            .CreateField eGDARRAY_Doubles
        Next lIndex
    End With
    
    astrFields.SplitFields m.strFields, ";"
    
    ' Load the table with non-boolean criterias from the symbol pool
    With g.SymbolPool
        lIndex = 0&
        For Each Criteria In .Criterias
            With Criteria
                If HasModule(.Required) Then
                     ''lPoolField = g.SymbolPool.ArrayTable.FieldNum(Parse(Parse(.GroupID, ":", 2), ".", 1))
'                    lPoolField = g.SymbolPool.FieldNumForID(.GroupID)
'                    If g.SymbolPool.ArrayTable(lPoolField, m.lPoolRec) <> 0 Then
                        TableNum(eTblCol_Show, lIndex) = 0
                        For lIndex2 = 0 To astrFields.Size - 1
                            If UCase(astrFields(lIndex2)) = UCase(.ID) Then
                                TableNum(eTblCol_Show, lIndex) = lIndex2 + 1
                                Exit For
                            End If
                        Next lIndex2
                        TableStr(eTblCol_Name, lIndex) = .Name
                        TableStr(eTblCol_Description, lIndex) = .Desc
                        On Error Resume Next
                        strText = .CodedText
                        If Len(strText) = 0 Then
                            strText = GetCodedText(.EnglishText)
                        End If
                        On Error GoTo ErrSection
                        TableStr(eTblCol_CodedText, lIndex) = strText
                        If Len(strText) = 0 Then
                            InfBox "Error in criteria expression for:|" & .Name, "e", , "Error"
                        End If
                        TableNum(eTblCol_NumDays, lIndex) = .NumDays
                        TableStr(eTblCol_ID, lIndex) = .ID
                        TableNum(eTblCol_IsWeekly, lIndex) = .IsWeekly
                        TableNum(eTblCol_IsBoolean, lIndex) = .IsBoolean
                        TableNum(eTblCol_PriceDisplay, lIndex) = .PriceDisplay
                        TableNum(eTblCol_DecimalPlaces, lIndex) = .DecimalPlaces
                        lIndex = lIndex + 1&
'                    End If
                End If
            End With
        Next
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.LoadCriteriaList", eGDRaiseError_Raise
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetCodedText
'' Description: Given an English version of a function call, hand back the
''              coded text for that expression
'' Inputs:      English function call to translate
'' Returns:     Coded text version of the expression
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetCodedText(ByVal strEnglishText As String) As String
On Error GoTo ErrSection:
   
    Dim lIndex As Long                  ' Index for a for loop
    Dim strChk As String                ' Temporary string for input checking
    Dim strNotKnown As String           ' Inputs not recognized
    Dim bExtraInputs As Boolean         ' Are there unrecognized inputs?
    Dim Expr As cExpression             ' Expression object for translation
    Dim Inputs As cInputs               ' Inputs collection in the expression
 
    ' Verify the expression to get the coded text from the english text
    Set Expr = New cExpression
    With Expr
        .PortfolioNavigator = False
        .Functions = g.Functions
        .ValidateFunctionRule strEnglishText
        If strEnglishText = "" Then
            lIndex = lIndex
        End If
    End With
        
    ' Check to make sure there are no unrecognized inputs
    bExtraInputs = False
    strNotKnown = ""
    If Not Expr.Inputs Is Nothing Then
        Set Inputs = Expr.Inputs
        For lIndex = 1 To Expr.Inputs.Count
            strChk = UCase(Inputs.Item(lIndex).ParmName)
            If strChk <> "WEEKLY" And _
                    strChk <> "GC" And _
                    strChk <> "MARKET1" Then
                strNotKnown = strNotKnown & "|" & Inputs.Item(lIndex).ParmName
                bExtraInputs = True
            End If
        Next
    End If
    
    If bExtraInputs Then
        InfBox "Error: Unrecognized items in expression:|" & strNotKnown & "|", "!", , "Error"
    Else
        GetCodedText = Expr.CodedText
    End If
    
    
ErrExit:
    Set Expr = Nothing
    Set Inputs = Nothing
    Exit Function

ErrSection:
    Set Expr = Nothing
    Set Inputs = Nothing
    RaiseError "frmSnapshot.GetCodedText", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitGridSingle
'' Description: Initialize the grid for "Single Symbol" mode
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitGridSingle(fgGrid As VSFlexGrid)
On Error GoTo ErrSection:

    Dim oFont As New StdFont
    Dim strFont As String               ' Font for the grid from the ini file
    Dim lRedraw As Long                 ' Current state of the grid redraw
    
    
    If fgGrid Is Nothing Then Exit Sub

    ' Get font from INI file
    strFont = GetIniFileProperty("Fundamental", "", "Fonts", g.strIniFile)
    
    ' Initialize the grid
    SetupGrid fgGrid, eGridMode_Grid
    With fgGrid
        lRedraw = .Redraw
        .Redraw = flexRDNone
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Cols = GDCol(eGDCol_NumCols)
        .ExplorerBar = flexExSortShow
        .ExtendLastCol = False
        .FixedRows = 1
        .FixedCols = 0
        .Rows = 2
        .RowHeight(0) = .RowHeight(1) * 2
        .Rows = 1
        .FrozenCols = 1
        
        .TextMatrix(0, GDCol(eGDCol_ID)) = "Data Kind ID"
        .TextMatrix(0, GDCol(eGDCol_Name)) = "Type Of Data" ' & vbCrLf & "(dbl-click grid to edit rows to display)"
        .TextMatrix(0, GDCol(eGDCol_Description)) = "Description"
        .TextMatrix(0, GDCol(eGDCol_Date)) = "Last Updated"
        .TextMatrix(0, GDCol(eGDCol_Symbol)) = "Symbol"
        .TextMatrix(0, GDCol(eGDCol_SubSector)) = "SubSector"
        .TextMatrix(0, GDCol(eGDCol_Sector)) = "Sector"
        .TextMatrix(0, GDCol(eGDCol_SP500)) = "S&P 500" & vbCrLf & "Stocks"
        .TextMatrix(0, GDCol(eGDCol_Comparison)) = "Compare"
        
        .ColAlignment(GDCol(eGDCol_Description)) = flexAlignLeftCenter
        .ColAlignment(GDCol(eGDCol_Name)) = flexAlignLeftCenter
        .ColAlignment(GDCol(eGDCol_Date)) = flexAlignCenterCenter
        .ColAlignment(GDCol(eGDCol_Symbol)) = flexAlignRightCenter
        .ColAlignment(GDCol(eGDCol_SubSector)) = flexAlignRightCenter
        .ColAlignment(GDCol(eGDCol_Sector)) = flexAlignRightCenter
        .ColAlignment(GDCol(eGDCol_SP500)) = flexAlignRightCenter
        .ColAlignment(GDCol(eGDCol_Comparison)) = flexAlignRightCenter
        
        .ColHidden(GDCol(eGDCol_Date)) = True
        .ColHidden(GDCol(eGDCol_ID)) = True
        .ColHidden(GDCol(eGDCol_Description)) = True
        
        Select Case m.strSecType
            Case eSYMType_Index
                If Left(m.strSymbol, 3) = "$--" Then
                    .ColHidden(GDCol(eGDCol_SubSector)) = True
                    .ColHidden(GDCol(eGDCol_Sector)) = True
                ElseIf Left(m.strSymbol, 2) = "$-" Then
                    .ColHidden(GDCol(eGDCol_SubSector)) = True
                    .ColHidden(GDCol(eGDCol_Sector)) = (m.lSectorID = 0)
                Else
                    .ColHidden(GDCol(eGDCol_SubSector)) = True
                    .ColHidden(GDCol(eGDCol_Sector)) = True
                End If
                
            Case eSYMType_Stock
                .ColHidden(GDCol(eGDCol_SubSector)) = (m.lSubSectorID = 0)
                .ColHidden(GDCol(eGDCol_Sector)) = (m.lSectorID = 0)
                
            Case Else
                .ColHidden(GDCol(eGDCol_SubSector)) = True
                .ColHidden(GDCol(eGDCol_Sector)) = True
        
        End Select
        
        If Len(strFont) > 0 Then
            'JM(10-19-2009) - for some reason, just passing the grid's font object makes the font bold every 3rd time
            If FontFromString(oFont, strFont) Then
                .Font.Bold = oFont.Bold
                .FontItalic = oFont.Italic
                .Font.Size = oFont.Size
                .FontUnderline = oFont.Underline
                .Font.Name = oFont.Name
                .FontStrikethru = oFont.Strikethrough
            End If
        End If
        
        .ColFormat(GDCol(eGDCol_Date)) = DateFormat("Format", MM_DD_YYYY)
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
                
        .Editable = flexEDNone
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.InitGridSingle", eGDRaiseError_Raise
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitGridComparison
'' Description: Initialize the grid for "Comparison" mode
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitGridComparison(fgGrid As VSFlexGrid)
On Error GoTo ErrSection:

    Dim strFont As String               ' Font for the grid from the ini file
    Dim lRedraw As Long                 ' Current state of the grid redraw


    If fgGrid Is Nothing Then Exit Sub

    ' Get font from INI file
    strFont = GetIniFileProperty("Fundamental", "", "Fonts", g.strIniFile)
    
    ' Initialize the grid
    SetupGrid fgGrid, eGridMode_Grid
    With fgGrid
        lRedraw = .Redraw
        .Redraw = flexRDNone
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Rows = .FixedRows
        .Rows = GDCol(eGDCol_NumCols) + m.alSymbolIds.Size
        .Cols = 1
        .ExplorerBar = flexExSortShow
        .ExtendLastCol = False
        .FixedRows = 1
        .FixedCols = 0
        .RowHeight(0) = .RowHeight(1)
        .FrozenCols = 1
        
        .TextMatrix(GDCol(eGDCol_Name), 0) = "Symbol"
        .TextMatrix(GDCol(eGDCol_ID), 0) = "Data Kind ID"
        .TextMatrix(GDCol(eGDCol_Description), 0) = "Description"
        .TextMatrix(GDCol(eGDCol_Date), 0) = "Last Updated"
        .TextMatrix(GDCol(eGDCol_Symbol), 0) = "Symbol"
        .TextMatrix(GDCol(eGDCol_SubSector), 0) = "SubSector"
        .TextMatrix(GDCol(eGDCol_Sector), 0) = "Sector"
        .TextMatrix(GDCol(eGDCol_SP500), 0) = "S&P 500 Stocks"
        .TextMatrix(GDCol(eGDCol_Comparison), 0) = "Compare"
        
        .RowHidden(GDCol(eGDCol_ID)) = True
        .RowHidden(GDCol(eGDCol_Description)) = True
        .RowHidden(GDCol(eGDCol_Date)) = True
        .RowHidden(GDCol(eGDCol_Symbol)) = True
        .RowHidden(GDCol(eGDCol_SubSector)) = True
        .RowHidden(GDCol(eGDCol_Sector)) = True
        .RowHidden(GDCol(eGDCol_SP500)) = True
        .RowHidden(GDCol(eGDCol_Comparison)) = True
        
        If strFont <> "" Then FontFromString .Font, strFont
        
        .Editable = flexEDNone
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.InitGridComparison", eGDRaiseError_Raise
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuRemoveSymbol_Click
'' Description: Allow the user to remove a comparison symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuRemoveCompare_Click()
On Error GoTo ErrSection:

    If fgData.Visible Then
        RemoveComparison fgData
    ElseIf fgData2.Visible Then
        RemoveComparison fgData2
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.mnuRemoveCompare.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ComparisonHidden
'' Description: Determines whether there is a comparison symbol or not
'' Inputs:      None
'' Returns:     True if no comparison symbol, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ComparisonHidden() As Boolean
On Error GoTo ErrSection:

    Select Case m.Mode
        Case eGDSnapshotMode_Single
            If fgData.Visible Then
                ComparisonHidden = fgData.ColHidden(GDCol(eGDCol_Comparison))
            ElseIf fgData2.Visible Then
                ComparisonHidden = fgData2.ColHidden(GDCol(eGDCol_Comparison))
            End If
            
        Case eGDSnapshotMode_ShowSiblings
            If fgData.Visible Then
                ComparisonHidden = fgData.RowHidden(GDCol(eGDCol_Comparison))
            ElseIf fgData2.Visible Then
                ComparisonHidden = fgData2.RowHidden(GDCol(eGDCol_Comparison))
            End If

    End Select

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSnapshot.ComparisonHidden", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToolTipSingle
'' Description: Determine the tool tip for the grid for "Single Symbol" mode
'' Inputs:      Row and Column of mouse
'' Returns:     ToolTipText to display
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ToolTipTextSingle(ByVal Row As Long, ByVal Col As Long) As String
On Error GoTo ErrSection:

    Dim fgGrid As VSFlexGrid
    
    If fgData.Visible Then
        Set fgGrid = fgData
    ElseIf fgData2.Visible Then
        Set fgGrid = fgData2
    End If
    
    If fgGrid Is Nothing Then Exit Function

    With fgGrid
        If Row < .FixedRows And Row >= 0 Then
            Select Case Col
                Case GDCol(eGDCol_Name)
                    ToolTipTextSingle = SORT_BY_PREFIX & "Data Type"
                
                Case GDCol(eGDCol_Symbol)
                    If Len(m.strDesc) = 0 Then
                        ToolTipTextSingle = SORT_BY_PREFIX & "Values for " & m.strSymbol
                    Else
                        ToolTipTextSingle = SORT_BY_PREFIX & "Values for " & m.strSymbol & " (" & m.strDesc & ")"
                    End If
                
                Case GDCol(eGDCol_SubSector)
                    If Len(m.strSubsectorDesc) = 0 Then
                        ToolTipTextSingle = SORT_BY_PREFIX & "Values for " & m.strSubsector
                    Else
                        ToolTipTextSingle = SORT_BY_PREFIX & "Values for " & m.strSubsector & " (" & m.strSubsectorDesc & ")"
                    End If
                
                Case GDCol(eGDCol_Sector)
                    If Len(m.strSectorDesc) = 0 Then
                        ToolTipTextSingle = SORT_BY_PREFIX & "Values for " & m.strSector
                    Else
                        ToolTipTextSingle = SORT_BY_PREFIX & "Values for " & m.strSector & " (" & m.strSectorDesc & ")"
                    End If
                
                Case GDCol(eGDCol_SP500)
                    ToolTipTextSingle = SORT_BY_PREFIX & "Values for S&P 500"
                    
                Case GDCol(eGDCol_Comparison)
                    If Len(m.strComparisonDesc) = 0 Then
                        ToolTipTextSingle = SORT_BY_PREFIX & "Values for " & m.strComparison
                    Else
                        ToolTipTextSingle = SORT_BY_PREFIX & "Values for " & m.strComparison & " (" & m.strComparisonDesc & ")"
                    End If
            
            End Select
        
        ElseIf Col = GDCol(eGDCol_Name) Then
            If Row >= .FixedRows And Row < .Rows Then
                ToolTipTextSingle = .TextMatrix(Row, GDCol(eGDCol_Description))
            Else
                ToolTipTextSingle = ""
            End If
        Else
            ToolTipTextSingle = ""
        End If
    End With

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSnapshot.ToolTipTextSingle", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToolTipComparison
'' Description: Determine the tool tip for the grid for "Comparison" mode
'' Inputs:      Row and Column of mouse
'' Returns:     ToolTipText to display
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ToolTipTextComparison(ByVal Row As Long, ByVal Col As Long) As String
On Error GoTo ErrSection:

    Dim fgGrid As VSFlexGrid
    
    If fgData.Visible Then
        Set fgGrid = fgData
    ElseIf fgData2.Visible Then
        Set fgGrid = fgData2
    End If
    
    If fgGrid Is Nothing Then Exit Function

    With fgGrid
        If Col = 0 Then
            Select Case Row
                Case GDCol(eGDCol_Name)
                    ToolTipTextComparison = SORT_BY_PREFIX & "Symbol"
                
                Case -1
                    ToolTipTextComparison = ""
                
                Case Else
                    ToolTipTextComparison = .RowData(Row)
                            
            End Select
        
        ElseIf Row = GDCol(eGDCol_Name) Then
            If Col >= 0 And Col < .Cols Then
                ToolTipTextComparison = SORT_BY_PREFIX & .TextMatrix(GDCol(eGDCol_Description), Col)
            Else
                ToolTipTextComparison = ""
            End If
        Else
            ToolTipTextComparison = ""
        End If
    End With

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmSnapshot.ToolTipTextSingle", eGDRaiseError_Raise
    
End Function

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

    'tbToolbar.Tools("ID_ShowSiblings").Visible = (m.strSecType = eSYMType_Stock Or Left(m.strSymbol, 2) = "$-")
    tbToolbar.Tools("ID_ShowSiblings").Visible = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.EnableControls", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddComparison
'' Description: Allow the user to add a comparison symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddComparison(fgGrid As VSFlexGrid)
On Error GoTo ErrSection:

    Dim astrSymbol As cGdArray
    Dim lPos As Long
    Dim lPoolRec As Long
    
    
    If fgGrid Is Nothing Then Exit Sub
    
    Set astrSymbol = New cGdArray
    astrSymbol.Create eGDARRAY_Strings
    
    Set astrSymbol = frmSymbolSelector.ShowMe("", False, True, "Comparison Symbol", False)
    If astrSymbol.Size > 0 Then
        m.lCompareID = g.SymbolPool.SymbolIDforSymbol(astrSymbol(0))
    
        m.astrCompare.BinarySearch m.strSymbol, lPos
        If Parse(m.astrCompare(lPos), vbTab, 1) = m.strSymbol Then
            m.astrCompare(lPos) = m.strSymbol & vbTab & Str(m.lCompareID)
        Else
            m.astrCompare.Add m.strSymbol & vbTab & Str(m.lCompareID), lPos
        End If
        m.astrCompare.ToFile AddSlash(App.Path) & "Compare.SYM"
        
        lPoolRec = g.SymbolPool.PoolRecForSymbolID(m.lCompareID)
        m.strComparison = g.SymbolPool.Symbol(lPoolRec)
        m.strComparisonDesc = g.SymbolPool.Desc(lPoolRec)
        
        Select Case m.Mode
            Case eGDSnapshotMode_Single
                InitGridSingle fgGrid
                LoadCriteriaList
                CalcValues
                LoadGridSingle fgGrid
                
            Case eGDSnapshotMode_ShowSiblings
                InitGridComparison fgGrid
                LoadCriteriaList
                CalcValues
                LoadGridComparison fgGrid
                
        End Select
    End If
    
    EnableControls

ErrExit:
    Set astrSymbol = Nothing
    Exit Sub
    
ErrSection:
    Set astrSymbol = Nothing
    RaiseError "frmSnapshot.AddComparison", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RemoveComparison
'' Description: Allow the user to remove a comparison symbol
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RemoveComparison(fgGrid As VSFlexGrid)
On Error GoTo ErrSection:

    Dim lPos As Long
    
    If fgGrid Is Nothing Then Exit Sub
    
    m.astrCompare.BinarySearch m.strSymbol, lPos
    If Parse(m.astrCompare(lPos), vbTab, 1) = m.strSymbol Then
        m.astrCompare.Remove lPos
    End If
    m.astrCompare.ToFile AddSlash(App.Path) & "Compare.SYM"
    
    m.lCompareID = 0&
    
    Select Case m.Mode
        Case eGDSnapshotMode_Single
            fgData.ColHidden(GDCol(eGDCol_Comparison)) = True
            fgData2.ColHidden(GDCol(eGDCol_Comparison)) = True
            LoadGridSingle fgGrid
            
        Case eGDSnapshotMode_ShowSiblings
            fgData.RowHidden(GDCol(eGDCol_Comparison)) = True
            fgData2.RowHidden(GDCol(eGDCol_Comparison)) = True
            LoadGridComparison fgGrid
            
    End Select
    
    EnableControls
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.RemoveComparison", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    tbToolbar_ToolClick
'' Description: Handle a user choice on the toolbar
'' Inputs:      Tool the user clicked on
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub tbToolbar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
On Error GoTo ErrSection:

    If Not Me.Visible Then Exit Sub
    If m.bDontProcess Then Exit Sub

    Select Case Tool.ID
        Case "ID_ChangeSymbol"
            m.bSyncToActiveChart = False
            ChangeSymbol
    
        Case "ID_Print"
            PrintMe
        
        Case "ID_EditFields"
            If fgData.Visible Then                  '5811
                EditFieldsToShow fgData
            ElseIf fgData2.Visible Then
                EditFieldsToShow fgData2
            End If
                    
        Case "ID_ShowSiblings"
            If Tool.State = ssChecked Then
                SetIniFileProperty "Mode", eGDSnapshotMode_ShowSiblings, "Fundamentals", g.strIniFile
                ShowMe m.lSymbolID
            Else
                SetIniFileProperty "Mode", eGDSnapshotMode_Single, "Fundamentals", g.strIniFile
                ShowMe m.lSymbolID
            End If
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.tbToolbar.ToolClick", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeSymbol
'' Description: Allow the user to change the symbol they are viewing
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ChangeSymbol()
On Error GoTo ErrSection:

    Dim astrSymbols As New cGdArray     ' Symbol(s) back from the symbol selector
    Dim lSymbolID As Long               ' Symbol ID for the symbol selected
    
    astrSymbols.Create eGDARRAY_Strings

    Set astrSymbols = frmSymbolSelector.ShowMe("", False)
    If astrSymbols.Size > 0 Then
        lSymbolID = g.SymbolPool.SymbolIDforSymbol(astrSymbols(0))
    End If
    If lSymbolID = 0 Then
        Beep
    Else
        m.bSyncToActiveChart = False
        ShowMe lSymbolID
    End If

ErrExit:
    Set astrSymbols = Nothing
    Exit Sub
    
ErrSection:
    Set astrSymbols = Nothing
    RaiseError "frmSnapshot.ChangeSymbol", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HighlightSymbol
'' Description: Highlight the line with the given symbol
'' Inputs:      Symbol to highlight
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HighlightSymbol(ByVal strSymbol As String)
On Error GoTo ErrSection:

    Dim fgGrid As VSFlexGrid
    Dim lIndex As Long                  ' Index into a for loop
    
    If fgData.Visible Then
        Set fgGrid = fgData
    ElseIf fgData2.Visible Then
        Set fgGrid = fgData2
    End If
    
    If fgGrid Is Nothing Then Exit Sub
    
    If m.Mode = eGDSnapshotMode_ShowSiblings And Len(strSymbol) > 0 Then
        With fgGrid
            For lIndex = .FixedRows To .Rows - 1
                If .TextMatrix(lIndex, 0) = strSymbol Then
                    .Row = lIndex
                    .RowSel = lIndex
                    .ShowCell .Row, 0
                    Exit For
                End If
            Next lIndex
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.HighlightSymbol", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RenameCriteria
'' Description: Handle a renamed criteria file (which also changes the ID)
'' Inputs:      Old Criteria ID, New Criteria ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub RenameCriteria(ByVal strOldCriteriaID As String, ByVal strNewCriteriaID As String)
On Error GoTo ErrSection:

    Dim strFields As String             ' Fields for the form
    
    strFields = Replace(m.strFields, UCase(strOldCriteriaID), UCase(strNewCriteriaID))
    If strFields <> m.strFields Then
        If m.strSecType = eSYMType_Stock Or Left(m.strSymbol, 2) = "$-" Then
            SetIniFileProperty "FieldsS", strFields, "Fundamentals", g.strIniFile
        ElseIf m.strSecType = eSYMType_Future Then
            SetIniFileProperty "FieldsF", strFields, "Fundamentals", g.strIniFile
        Else
            SetIniFileProperty "Fields", strFields, "Fundamentals", g.strIniFile
        End If
        m.strFields = strFields
        
        ShowMe m.lSymbolID, , strFields
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.RenameCriteria"
    
End Sub

Private Sub tmr_Timer()
On Error Resume Next

    Static bInProgress As Boolean
    Dim strUrl$
    
    TimerStart "frmSnapshot.tmr"
    If bInProgress Then Exit Sub
    
    bInProgress = True

    If tmr.Tag = "PreloadCash" Then
        strUrl = m.strBaseURL & m.strSymbol & "&statement=cash"
        StatusMsg "Preloading Cash Flow"
        webBrowser.Navigate2 strUrl
        m.bCashPreloaded = True
        If vsTab.CurrTab = eVSTab_Custom And Not m.bBalancePreloaded Then
            tmr.Tag = "PreloadBalance"
        Else
            tmr.Enabled = False
        End If
    ElseIf tmr.Tag = "PreloadBalance" Then
        strUrl = m.strBaseURL & m.strSymbol & "&statement=balance"
        StatusMsg "Preloading Balance"
        webBrowser.Navigate2 strUrl
        m.bBalancePreloaded = True
        If m.bShowGeneralWeb Then
            If vsTab.CurrTab = eVSTab_Custom And Not m.bGeneralPreLoaded Then
                tmr.Tag = "PreloadGeneral"
            Else
                tmr.Enabled = False
            End If
        Else
            tmr.Enabled = False
        End If
    ElseIf tmr.Tag = "PreloadGeneral" Then
        strUrl = m.strBaseURL & m.strSymbol & "&statement=general"
        StatusMsg "Preloading General"
        webBrowser.Navigate2 strUrl
        m.bGeneralPreLoaded = True
        tmr.Enabled = False
    ElseIf vsTab.CurrTab = eVSTab_Custom Then
        If Not m.bCashPreloaded Then
            tmr.Tag = "PreloadCash"
        ElseIf Not m.bBalancePreloaded Then
            tmr.Tag = "PreloadBalance"
        ElseIf m.bShowGeneralWeb And Not m.bGeneralPreLoaded Then
            tmr.Tag = "PreloadGeneral"
        Else
            tmr.Enabled = False
        End If
    ElseIf Not webBrowser.Visible Then
        'if it's been half a second and web browser is still not showing then display loading message
        lblLoading.Visible = True
        tmr.Enabled = False
    Else
        tmr.Enabled = False
    End If
    
    bInProgress = False
    TimerEnd "frmSnapshot.tmr", tmr.Interval
    
End Sub

Private Sub vsTab_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
On Error GoTo ErrSection:

    Dim strUrl$
    
    If g.bStarting Then Exit Sub
    
    If NewTab = eVSTab_Custom Then
        If Not m.bBalancePreloaded Then
            If Len(tmr.Tag) = 0 Then
                tmr.Tag = "PreloadBalance"
                If Not tmr.Enabled Then tmr.Enabled = True
            End If
        ElseIf Not m.bCashPreloaded Then
            If Len(tmr.Tag) = 0 Then
                tmr.Tag = "PreloadCash"
                If Not tmr.Enabled Then tmr.Enabled = True
            End If
        End If
    ElseIf Len(tmr.Tag) > 0 Then
        StatusMsg "Preload in progress, cannot switch tab."
        tmr.Tag = ""
        NewTab = OldTab
    End If
    
    If NewTab <> eVSTab_Custom Then
        strUrl = m.strBaseURL
        Select Case NewTab
            Case eVSTab_Cash
                strUrl = strUrl & m.strSymbol & "&statement=cash"
                m.bCashPreloaded = True
            Case eVSTab_Balance
                strUrl = strUrl & m.strSymbol & "&statement=balance"
                m.bBalancePreloaded = True
            Case eVSTab_Income
                strUrl = strUrl & m.strSymbol & "&statement=income"
            Case eVSTab_General
                strUrl = strUrl & m.strSymbol & "&statement=general"
        End Select
        
        If Len(strUrl) > 0 Then
            If strUrl <> m.strLastGoodURL Then
                If m.bShowURL Then StatusMsg strUrl
                webBrowser.Navigate2 strUrl
                webBrowser.Visible = False
                tmr.Enabled = True
            Else
                Screen.MousePointer = vbDefault
            End If
        End If
    End If
            
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.vsTab_Switch"

End Sub

Private Sub webBrowser_BeforeNavigate2(ByVal pDisp As Object, Url As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    If Len(tmr.Tag) = 0 Then Screen.MousePointer = vbHourglass
End Sub

Private Sub webBrowser_NavigateComplete2(ByVal pDisp As Object, Url As Variant)
On Error GoTo ErrSection:

    m.strLastGoodURL = Url
    lblLoading.Visible = False
    Screen.MousePointer = vbDefault
    
    If Len(tmr.Tag) = 0 Then
        webBrowser.Visible = True
    Else
        tmr.Tag = ""
    End If
    
    StatusMsg ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.webBrowser_NavigateComplete2"

End Sub

Private Sub GridAfterSort(fgGrid As VSFlexGrid, ByVal Col As Long, Order As Integer)
On Error GoTo ErrSection:

    If fgGrid Is Nothing Then Exit Sub

    SetBackColors fgGrid
    HighlightSymbol m.strSaveSymbol

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.GridAfterSort"

End Sub

Private Sub GridAfterUserResize(fgGrid As VSFlexGrid, ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim lWidth As Long                  ' Amount to adjust the column width
    Dim lIndex As Long                  ' Index into a for loop
    
    If fgGrid Is Nothing Then Exit Sub
    
    ' if column being resized is the extended column,
    ' then make the next column bigger (instead of adjusting
    ' the extended column)
    If Col >= kExtendedCol Then
        With fgGrid
            .Redraw = flexRDNone
            lWidth = .ColWidth(Col) - m.lPrevColWidth
            For lIndex = Col + 1 To .Cols - 1
                If Not .ColHidden(lIndex) Then
                    .ColWidth(lIndex) = fgGrid.ColWidth(lIndex) - lWidth
                    Exit For
                End If
            Next
            m.lPrevColWidth = 0
            ExtendCustomColumn fgGrid
            .Redraw = flexRDBuffered
        End With
   Else
        ExtendCustomColumn fgGrid
   End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.GridAfterUserResize"

End Sub

Private Sub GridBeforeSort(fgGrid As VSFlexGrid, ByVal Col As Long, Order As Integer)
On Error GoTo ErrSection:

    If fgGrid Is Nothing Then Exit Sub

    If m.Mode = eGDSnapshotMode_ShowSiblings Then
        If Col = 0 Then
            If Order = flexSortGenericAscending Then
                Order = flexSortStringAscending
            ElseIf Order = flexSortGenericDescending Then
                Order = flexSortStringDescending
            End If
        End If
        
        m.strSaveSymbol = fgGrid.TextMatrix(fgGrid.RowSel, 0)
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.GridBeforeSort"

End Sub

Private Sub GridBeforeUserResize(fgGrid As VSFlexGrid, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    If Not fgGrid Is Nothing Then
        ' if column being resized is the extended column, save size
        If Col >= kExtendedCol Then
            m.lPrevColWidth = fgGrid.ColWidth(Col)
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.GridBeforeUserResize"

End Sub

Private Sub GridDblClick(fgGrid As VSFlexGrid)
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Current row of the mouse
    Dim lMouseCol As Long               ' Current column of the mouse
    Dim lSymbolID As Long

    
    If fgGrid Is Nothing Then Exit Sub
    
    lMouseRow = fgGrid.MouseRow
    lMouseCol = fgGrid.MouseCol

    Select Case m.Mode
        Case eGDSnapshotMode_Single
            If lMouseRow >= fgGrid.FixedRows And lMouseRow < fgGrid.Rows Then
                Select Case lMouseCol
                Case eGDCol_SubSector
                    lSymbolID = m.lSubSectorID
                Case eGDCol_Sector
                    lSymbolID = m.lSectorID
                Case eGDCol_SP500
                    lSymbolID = m.lSP500ID
                Case Else
                    lSymbolID = m.lSymbolID
                End Select
                SetActiveChartSymbol lSymbolID
            End If
        
        Case eGDSnapshotMode_ShowSiblings
            If lMouseRow >= fgGrid.FixedRows And lMouseRow < fgGrid.Rows Then
                SetActiveChartSymbol fgGrid.TextMatrix(lMouseRow, 0)
            End If
    
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.GridDblClick"

End Sub

Private Sub GridMouseDown(fgGrid As VSFlexGrid, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Current row of the mouse
    Dim lMouseCol As Long               ' Current column of the mouse

    If fgGrid Is Nothing Then Exit Sub
    
    With fgGrid
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If lMouseRow >= .FixedRows And lMouseRow < .Rows Then
            .Row = lMouseRow
            .RowSel = .Row
        End If
            
        If Button = vbRightButton Then
            If ComparisonHidden Then
                mnuCompare.Caption = "Add Comparison Symbol"
                mnuRemoveCompare.Visible = False
            Else
                mnuCompare.Caption = "Change Comparison Symbol"
                mnuRemoveCompare.Visible = True
            End If
            
            PopupMenu mnuPref
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.GridMouseDown"

End Sub

Private Sub GridMouseMove(fgGrid As VSFlexGrid, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim lMouseRow As Long
    Dim lMouseCol As Long

    If fgGrid Is Nothing Then Exit Sub

    With fgGrid
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
    
        Select Case m.Mode
            Case eGDSnapshotMode_Single
                .ToolTipText = ToolTipTextSingle(lMouseRow, lMouseCol)
            
            Case eGDSnapshotMode_ShowSiblings
                .ToolTipText = ToolTipTextComparison(lMouseRow, lMouseCol)
        
        End Select
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmSnapshot.GridMouseMove"

End Sub

