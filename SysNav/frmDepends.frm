VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmDepends 
   Caption         =   "Form1"
   ClientHeight    =   6735
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   8595
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   8595
   StartUpPosition =   3  'Windows Default
   Begin VSFlex7LCtl.VSFlexGrid fgUsedIn 
      Height          =   1995
      Left            =   120
      TabIndex        =   4
      Top             =   3300
      Visible         =   0   'False
      Width           =   3735
      _cx             =   6588
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
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   3015
      Left            =   5100
      TabIndex        =   1
      Top             =   240
      Width           =   1095
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
      Caption         =   "frmDepends.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmDepends.frx":0020
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmDepends.frx":0040
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniCheckXP chkRecursive 
         Height          =   220
         Left            =   0
         TabIndex        =   5
         Top             =   1500
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   397
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmDepends.frx":005C
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmDepends.frx":0090
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmDepends.frx":00B0
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdClose 
         Cancel          =   -1  'True
         Default         =   -1  'True
         Height          =   435
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1095
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
         Caption         =   "frmDepends.frx":00CC
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmDepends.frx":00F8
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmDepends.frx":0118
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdPrint 
         Height          =   435
         Left            =   0
         TabIndex        =   2
         Top             =   600
         Width           =   1095
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
         Caption         =   "frmDepends.frx":0134
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmDepends.frx":0160
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmDepends.frx":0180
         RightToLeft     =   0   'False
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgDepends 
      Height          =   2475
      Left            =   120
      TabIndex        =   0
      Top             =   420
      Width           =   3735
      _cx             =   6588
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
   Begin HexUniControls.ctlUniLabelXP lblUsedIn 
      Height          =   315
      Left            =   120
      Top             =   3000
      Visible         =   0   'False
      Width           =   3795
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
      Caption         =   "frmDepends.frx":019C
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmDepends.frx":01D0
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmDepends.frx":01F0
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblDependsOn 
      Height          =   315
      Left            =   120
      Top             =   120
      Width           =   3675
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
      Caption         =   "frmDepends.frx":020C
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmDepends.frx":0246
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmDepends.frx":0266
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print Dependencies"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangeFont 
         Caption         =   "&Change Font"
      End
   End
End
Attribute VB_Name = "frmDepends"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmDepends.frm
'' Description: Shows dependencies for a System, Rule, or Function
''
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Public Enum eDepends
    eDepends_System = 0
    eDepends_Rule = 1
    eDepends_Function = 2
End Enum

Private Enum eGDCols
    eGDCol_ID = 0
    eGDCol_Name
    eGDCol_Type
    eGDCol_LibraryID
    eGDCol_LibraryName
    eGDCol_NumCols
End Enum

Private Type mPrivate
    strName As String                   ' Name of the item
    lID As Long                         ' ID of the item
    Depends As eDepends                 ' Type of the item
    bRecursive As Boolean               ' Calculate recursively?
End Type
Private m As mPrivate

Private Function GDCol(ByVal Col As eGDCols) As Long
    GDCol = Col
End Function

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

    With fgDepends
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .ExplorerBar = flexExSortShow
        .ExtendLastCol = True
        .SelectionMode = flexSelectionListBox
        .Editable = flexEDNone
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .AllowUserResizing = flexResizeColumns
        .ScrollTrack = True
        
        .Rows = 1
        .FixedRows = 1
        .Cols = GDCol(eGDCol_NumCols)
        .FixedCols = 0
        
        .TextMatrix(0, GDCol(eGDCol_ID)) = "ID"
        .TextMatrix(0, GDCol(eGDCol_Name)) = "Name"
        .TextMatrix(0, GDCol(eGDCol_Type)) = "Type"
        .TextMatrix(0, GDCol(eGDCol_LibraryID)) = "Library ID"
        .TextMatrix(0, GDCol(eGDCol_LibraryName)) = "Library"
        
        .ColHidden(GDCol(eGDCol_ID)) = True
        .ColHidden(GDCol(eGDCol_LibraryID)) = True
        
        .Redraw = lRedraw
    End With

    With fgUsedIn
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .ExplorerBar = flexExSortShow
        .ExtendLastCol = True
        .SelectionMode = flexSelectionListBox
        .Editable = flexEDNone
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .AllowUserResizing = flexResizeColumns
        .ScrollTrack = True
        
        .Rows = 1
        .FixedRows = 1
        .Cols = GDCol(eGDCol_NumCols)
        .FixedCols = 0
        
        .TextMatrix(0, GDCol(eGDCol_ID)) = "ID"
        .TextMatrix(0, GDCol(eGDCol_Name)) = "Name"
        .TextMatrix(0, GDCol(eGDCol_Type)) = "Type"
        .TextMatrix(0, GDCol(eGDCol_LibraryID)) = "Library ID"
        .TextMatrix(0, GDCol(eGDCol_LibraryName)) = "Library"
        
        .ColHidden(GDCol(eGDCol_ID)) = True
        .ColHidden(GDCol(eGDCol_LibraryID)) = True
        
        .Redraw = lRedraw
    End With

ErrExit:
    fgDepends.Redraw = lRedraw
    Exit Sub
    
ErrSection:
    RaiseError "frmDepends.InitGrid", eGDRaiseError_Raise

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

    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim rs As Recordset                 ' Recordset for working with the database
    Dim astrDepends As New cGdArray     ' Array to store dependencies in
    Dim lIndex As Long                  ' Index into a for loop

    astrDepends.Create eGDARRAY_Strings

    Select Case m.Depends
        Case eDepends_System
            SetEditorCaption Me, "Strategy Dependencies", m.strName
            Me.Icon = Picture16(ToolbarIcon("ID_Strategies"), , True)
            Set rs = g.dbNav.OpenRecordset("SELECT tblRules.Name, tblRules.RuleID, tblLibrarys.LibraryID, tblLibrarys.LibraryName " & _
                    "FROM tblLibrarys INNER JOIN (tblRules INNER JOIN tblSystemRules ON tblRules.RuleID = tblSystemRules.RuleID) ON tblLibrarys.LibraryID = tblRules.LibraryID " & _
                    "WHERE (((tblSystemRules.SystemNumber)=" & Str(m.lID) & "));", dbOpenDynaset)
            If Not rs.EOF Then
                rs.MoveFirst
                Do While Not rs.EOF
                    RuleDepends astrDepends, rs!RuleID, m.bRecursive
                    
                    rs.MoveNext
                Loop
            End If
                    
        Case eDepends_Function
            SetEditorCaption Me, "Function Dependencies", m.strName
            Me.Icon = Picture16(ToolbarIcon("ID_Functions"), , True)
            FuncDepends astrDepends, m.lID, m.bRecursive
            
        Case eDepends_Rule
            SetEditorCaption Me, "Rule Dependencies", m.strName
            Me.Icon = Picture16(ToolbarIcon("ID_Rules"), , True)
            RuleDepends astrDepends, m.lID, m.bRecursive
            
    End Select
    
    astrDepends.Sort eGdSort_DeleteDuplicates Or eGdSort_IgnoreCase
        
    With fgDepends
        lRedraw = .Redraw
        .Redraw = flexRDNone
        .Rows = .FixedRows

        For lIndex = 0 To astrDepends.Size - 1
            .AddItem astrDepends(lIndex)
        Next lIndex
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With
    
    
    If m.Depends = eDepends_Function Then
        astrDepends.Size = 0
        UsedInFunc astrDepends, m.lID, m.bRecursive
        UsedInRule astrDepends, m.lID
        astrDepends.Sort eGdSort_DeleteDuplicates Or eGdSort_IgnoreCase
    
        With fgUsedIn
            .Redraw = flexRDNone
            .Rows = .FixedRows
            
            For lIndex = 0 To astrDepends.Size - 1
                .AddItem astrDepends(lIndex)
            Next lIndex
        
            .AutoSize 0, .Cols - 1, False, 75
            .Redraw = flexRDBuffered
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDepends.LoadGrid", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Intialize and show the form
'' Inputs:      Item Type, Item ID, Item Name
'' Returns:     True on success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ByVal Depends As eDepends, ByVal lID As Long, ByVal strName As String) As Boolean
On Error GoTo ErrSection:

    Static bAlreadyDone As Boolean      ' Have we already done this?
    
    If Not bAlreadyDone Then
        bAlreadyDone = True
        ' just first time -- set default
        m.bRecursive = True
    End If
    chkRecursive = Abs(m.bRecursive)

    m.Depends = Depends
    m.lID = lID
    m.strName = strName

    fgDepends.Redraw = flexRDNone
    InitGrid
    LoadGrid
    fgDepends.Redraw = flexRDBuffered
    
    lblUsedIn.Visible = (Depends = eDepends_Function)
    fgUsedIn.Visible = (Depends = eDepends_Function)

    ShowForm Me, True, , , ALT_GRID_ROW_COLOR
    ShowMe = True
    
ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmDepends.ShowMe", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkRecursive_Click
'' Description: Allow the user to choose whether to show recursive dependencies
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkRecursive_Click()
On Error GoTo ErrSection:

    If Me.Visible Then
        m.bRecursive = chkRecursive
        LoadGrid
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDepends.chkRecursive.Click", eGDRaiseError_Show
    Resume ErrExit
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

    Me.Hide
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDepends.cmdClose.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdPrint
'' Description: Bring up the Print Preview for the dependencies
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdPrint_Click()
On Error GoTo ErrSection:

    PrintMe
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDepends.cmdPrint.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgDepends_MouseDown
'' Description: Show the pop-up menu when the user right clicks
'' Inputs:      Button Pressed, Shift/Ctrl/Alt status, Location of Mouse
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgDepends_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:
    
    Dim lMouseRow As Long               ' Row the mouse is currently over
    Dim lMouseCol As Long               ' Column the mouse is currently over
    
    With fgDepends
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If Button = vbRightButton Then
            If lMouseRow >= .FixedRows And lMouseRow < .Rows Then
                .RowSel = lMouseRow
                If .IsSelected(lMouseRow) Then .Row = lMouseRow
            End If
            
            mnuPrint.Enabled = (.Rows > .FixedRows)
            
            PopupMenu mnuPopUp
        End If
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmDepends.fgDepends.MouseDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgDepends_MouseMove
'' Description: Show appropriate tooltips as the user moves the mouse
'' Inputs:      Button Pressed, Shift/Ctrl/Alt status, Location of Mouse
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgDepends_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    GridTooltip fgDepends
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_KeyDown
'' Description: Show the help if the user presses the F1 key
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
    RaiseError "frmDepends.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize the controls and the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strFont As String               ' Font from the ini file

    Me.Icon = Picture16("kBlank")
    CenterTheForm Me
    
    g.Styler.StyleForm Me
    
    mnuPopUp.Visible = False
    
    strFont = GetIniFileProperty("Depends", "", "Fonts", g.strIniFile)
    If strFont <> "" Then FontFromString fgDepends.Font, strFont
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDepends.Form.Load", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Resize and move the controls as the form resizes
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    Dim lHeight As Long                 ' Height the form should be

    With fraButtons
        .Move ScaleWidth - .Width - lblDependsOn.Left, lblDependsOn.Top
    End With
    
    With fgDepends
        If m.Depends = eDepends_Function Then
            lHeight = (ScaleHeight - (lblDependsOn.Height * 2) - (lblDependsOn.Top * 3)) / 2
        Else
            lHeight = ScaleHeight - lblDependsOn.Height - (lblDependsOn.Top * 2)
        End If
        
        .Move lblDependsOn.Left, lblDependsOn.Height + lblDependsOn.Top, _
             ScaleWidth - fraButtons.Width - (lblDependsOn.Left * 3), lHeight
    End With
    
    With lblUsedIn
        .Move .Left, fgDepends.Top + fgDepends.Height + lblDependsOn.Top
    End With
    
    With fgUsedIn
        .Move .Left, lblUsedIn.Top + lblUsedIn.Height, fgDepends.Width, lHeight
    End With

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenerateReport
'' Description: Callback function for the print preview form
'' Inputs:      Arguments
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GenerateReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:

    With frmPrintPreview.vp
        .StartDoc
        DoPrintHeader
               
        .Font.Name = "Times New Roman"
        .Font.Size = 14
        .Font.Bold = True
        .Text = Me.Caption & vbCrLf
        
        .Font.Size = 12
        .Text = "Depends On..." & vbLf
        If frmPrintPreview.GoingToFile Then
            frmPrintPreview.GridToTable fgDepends
        Else
            .RenderControl = fgDepends.hWnd
        End If
        
        If fgUsedIn.Visible Then
            .Text = vbLf & "Used In..." & vbLf
            If frmPrintPreview.GoingToFile Then
                frmPrintPreview.GridToTable fgUsedIn
            Else
                .RenderControl = fgUsedIn.hWnd
            End If
        End If
        
        .EndDoc
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDepends.GenerateReport", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Save the font when user unloads the form
'' Inputs:      Whether to Cancel the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    SetIniFileProperty "Depends", FontToString(fgDepends.Font), "Fonts", g.strIniFile

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmDepends.Form.Unload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuChangeFont_Click
'' Description: Allow the user to change the font on the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub mnuChangeFont_Click()
On Error GoTo ErrSection:

    ChangeGridFont fgDepends

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmDepends.mnuChangeFont.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    mnuPrint_Click
'' Description: Allow the user to print the grids
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
    RaiseError "frmDepends.mnuPrint.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PrintMe
'' Description: Print the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PrintMe()
On Error GoTo ErrSection:

    frmPrintPreview.ShowMe "Depends", Me, 0
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmDepends.PrintMe", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FuncDepends
'' Description: Figure out the dependant functions for the Function ID passed in
'' Inputs:      Array of Dependencies, ID to calc, Do Recursive check?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FuncDepends(astrDepends As cGdArray, ByVal lID As Long, ByVal bRecursive As Boolean)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database

    Set rs = g.dbNav.OpenRecordset("SELECT tblFunctions_1.FunctionName, tblFunctions_1.FunctionID, tblLibrarys.LibraryID, tblLibrarys.LibraryName " & _
            "FROM ((tblFunctions INNER JOIN tblFunctionRefs ON tblFunctions.FunctionID = tblFunctionRefs.FunctionID) INNER JOIN tblFunctions AS tblFunctions_1 ON tblFunctionRefs.FunctionIDRef = tblFunctions_1.FunctionID) INNER JOIN tblLibrarys ON tblFunctions_1.LibraryID = tblLibrarys.LibraryID " & _
            "WHERE (((tblFunctions.FunctionID)=" & Str(lID) & "));", dbOpenDynaset)
    If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
            astrDepends.Add Str(rs!FunctionID) & vbTab & rs!FunctionName & vbTab & "Function" & vbTab & Str(rs!LibraryID) & vbTab & rs!LibraryName
            If bRecursive Then
                FuncDepends astrDepends, rs!FunctionID, bRecursive
            End If
            rs.MoveNext
        Loop
    End If
            
ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    RaiseError "frmDepends.FuncDepends", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RuleDepends
'' Description: Figure out the dependant functions for the Rule ID passed in
'' Inputs:      Array of Dependencies, ID to calc, Do Recursive check?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RuleDepends(astrDepends As cGdArray, ByVal lID As Long, ByVal bRecursive As Boolean)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database

    Set rs = g.dbNav.OpenRecordset("SELECT tblFunctions.FunctionName, tblFunctions.FunctionID, tblLibrarys.LibraryID, tblLibrarys.LibraryName " & _
            "FROM tblRules INNER JOIN ((tblLibrarys INNER JOIN tblFunctions ON tblLibrarys.LibraryID = tblFunctions.LibraryID) INNER JOIN tblFunctionRules ON tblFunctions.FunctionID = tblFunctionRules.FunctionIDRef) ON tblRules.RuleID = tblFunctionRules.RuleID " & _
            "WHERE (((tblRules.RuleID)=" & Str(lID) & "));", dbOpenDynaset)

    If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
            astrDepends.Add Str(rs!FunctionID) & vbTab & rs!FunctionName & vbTab & "Function" & vbTab & Str(rs!LibraryID) & vbTab & rs!LibraryName
            FuncDepends astrDepends, rs!FunctionID, bRecursive
            
            rs.MoveNext
        Loop
    End If

ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    RaiseError "frmDepends.RuleDepends", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UsedInFunc
'' Description: Determine what functions this function id is used in
'' Inputs:      Array of Functions Used In, Function ID, Check Recursive?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UsedInFunc(astrDepends As cGdArray, ByVal lID As Long, ByVal bRecursive As Boolean)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database

    Set rs = g.dbNav.OpenRecordset("SELECT tblFunctions.FunctionName, tblFunctions.FunctionID, tblLibrarys.LibraryID, tblLibrarys.LibraryName " & _
            "FROM ((tblFunctions INNER JOIN tblFunctionRefs ON tblFunctions.FunctionID = tblFunctionRefs.FunctionID) INNER JOIN tblFunctions AS tblFunctions_1 ON tblFunctionRefs.FunctionIDRef = tblFunctions_1.FunctionID) INNER JOIN tblLibrarys ON tblFunctions_1.LibraryID = tblLibrarys.LibraryID " & _
            "WHERE (((tblFunctionRefs.FunctionIDRef)=" & Str(lID) & "));", dbOpenDynaset)

    If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
            astrDepends.Add Str(rs!FunctionID) & vbTab & rs!FunctionName & vbTab & "Function" & vbTab & Str(rs!LibraryID) & vbTab & rs!LibraryName
            If bRecursive Then
                UsedInFunc astrDepends, rs!FunctionID, bRecursive
            End If
            rs.MoveNext
        Loop
    End If
            
ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    RaiseError "frmDepends.UsedInFunc", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UsedInRule
'' Description: Determine what rules this function id is used in
'' Inputs:      Array of Rules Used In, Function ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UsedInRule(astrDepends As cGdArray, ByVal lID As Long)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database

    Set rs = g.dbNav.OpenRecordset("SELECT tblRules.Name, tblRules.RuleID, tblLibrarys.LibraryID, tblLibrarys.LibraryName " & _
            "FROM ((tblRules INNER JOIN tblFunctionRules ON tblRules.RuleID = tblFunctionRules.RuleID)) INNER JOIN tblLibrarys ON tblRules.LibraryID = tblLibrarys.LibraryID " & _
            "WHERE (((tblFunctionRules.FunctionIDRef)=" & Str(lID) & "));", dbOpenDynaset)

    If Not rs.EOF Then
        rs.MoveFirst
        Do While Not rs.EOF
            astrDepends.Add Str(rs!RuleID) & vbTab & rs!Name & vbTab & "Rule" & vbTab & Str(rs!LibraryID) & vbTab & rs!LibraryName
            rs.MoveNext
        Loop
    End If
            
ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    RaiseError "frmDepends.UsedInRule", eGDRaiseError_Raise
    
End Sub

