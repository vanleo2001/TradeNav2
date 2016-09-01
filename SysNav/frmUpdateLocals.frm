VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmUpdateLocals 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   1095
      Left            =   4440
      TabIndex        =   2
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
      Caption         =   "frmUpdateLocals.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmUpdateLocals.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmUpdateLocals.frx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Height          =   495
         Left            =   0
         TabIndex        =   1
         Top             =   600
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
         Caption         =   "frmUpdateLocals.frx":0068
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmUpdateLocals.frx":0096
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmUpdateLocals.frx":00B6
         RightToLeft     =   0   'False
      End
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
         Caption         =   "frmUpdateLocals.frx":00D2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmUpdateLocals.frx":00F8
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmUpdateLocals.frx":0118
         RightToLeft     =   0   'False
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgRules 
      Height          =   2115
      Left            =   180
      TabIndex        =   0
      Top             =   840
      Width           =   4095
      _cx             =   7223
      _cy             =   3731
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
   Begin HexUniControls.ctlUniLabelXP lblDescription 
      Height          =   615
      Left            =   180
      Top             =   120
      Width           =   4095
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
      Caption         =   "frmUpdateLocals.frx":0134
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmUpdateLocals.frx":02B0
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmUpdateLocals.frx":02D0
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmUpdateLocals"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum eGDCols
    eGDCol_Update = 0
    eGDCol_SystemName
    eGDCol_RuleID
    eGDCol_SystemID
    eGDCol_NumCols
End Enum

Private Type mPrivate
    Rule As cRule
    bOK As Boolean
End Type
Private m As mPrivate

Private Function GDCol(ByVal Col As eGDCols) As Long
    GDCol = Col
End Function

Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    m.bOK = False
    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmUpdateLocals.cmdCancel.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    m.bOK = True
    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmUdpateLocals.cmdOK.Click", eGDRaiseError_Show
    Resume ErrExit

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
    RaiseError "frmUpdateLocals.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    Me.Icon = Picture16(ToolbarIcon("ID_Rules"), , True)
    Caption = "Update Local Rules"
    CenterTheForm Me
    
    g.Styler.StyleForm Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmUpdateLocals.Form.Load", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub InitGrid()
On Error GoTo ErrSection:

    Dim lRedraw As Long

    With fgRules
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = True
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExSortShow
        .ExtendLastCol = True
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        
        .Cols = GDCol(eGDCol_NumCols)
        .FixedCols = 0
        .Rows = 1
        .FixedRows = 1
        
        .TextMatrix(0, GDCol(eGDCol_Update)) = "Update"
        .TextMatrix(0, GDCol(eGDCol_SystemName)) = "Strategy"
        
        .ColHidden(GDCol(eGDCol_RuleID)) = True
        .ColHidden(GDCol(eGDCol_SystemID)) = True
        
        .ColDataType(GDCol(eGDCol_Update)) = flexDTBoolean
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmUpdateLocals.InitGrid", eGDRaiseError_Raise

End Sub

Private Sub LoadGrid()
On Error GoTo ErrSection

    Dim rs As Recordset
    Dim lRedraw As Long
    
    With fgRules
        lRedraw = .Redraw
        .Redraw = flexRDNone
    
        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] " & _
                "WHERE [SystemNumber]<>0 AND [Name]='" & m.Rule.Name & "';", dbOpenDynaset)
        Do While Not rs.EOF
            .Rows = .Rows + 1
            CheckedCell(fgRules, .Rows - 1, GDCol(eGDCol_Update)) = True
            .TextMatrix(.Rows - 1, GDCol(eGDCol_SystemName)) = SystemNameForID(rs!SystemNumber)
            .TextMatrix(.Rows - 1, GDCol(eGDCol_RuleID)) = rs!RuleID
            .TextMatrix(.Rows - 1, GDCol(eGDCol_SystemID)) = rs!SystemNumber
            
            rs.MoveNext
        Loop
        
        .Redraw = lRedraw
    End With

ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    RaiseError "frmUpdateLocals.LoadGrid", eGDRaiseError_Raise
    
End Sub

Public Function ShowMe(ByVal Rule As cRule) As Boolean
On Error GoTo ErrSection:

    Set m.Rule = Rule
    
    InitGrid
    LoadGrid
    
    If fgRules.Rows > fgRules.FixedRows Then
        ShowForm Me, True, , , ALT_GRID_ROW_COLOR
        
        If m.bOK Then
            UpdateLocals
        End If
    End If
    
    ShowMe = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmUpdateLocals.ShowMe", eGDRaiseError_Raise
    
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        m.bOK = False
        Cancel = True
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmUpdateLocals.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub UpdateLocals()
On Error GoTo ErrSection:

    Dim lIndex As Long
    Dim lSystemID As Long
    Dim lRuleID As Long
    Dim lForm As Long
    Dim bFound As Boolean
    Dim NewRule As New cRule
    Dim frm As frmSystemManager
    
    With fgRules
        For lIndex = .FixedRows To .Rows - 1
            If CheckedCell(fgRules, lIndex, GDCol(eGDCol_Update)) Then
                lSystemID = CLng(.TextMatrix(lIndex, GDCol(eGDCol_SystemID)))
                lRuleID = CLng(.TextMatrix(lIndex, GDCol(eGDCol_RuleID)))
                
                bFound = False
                For lForm = 0 To Forms.Count - 1
                    If UCase(Forms(lForm).Name) = "FRMSYSTEMMANAGER" Then
                        If Forms(lForm).ID = lSystemID Then
                            Set frm = Forms(lForm)
                            bFound = True
                            Exit For
                        End If
                    End If
                Next lForm
                
                NewRule.LoadWithSystemInfo lRuleID
                NewRule.CopyRuleInfo m.Rule
                
                If bFound Then
                    frm.AddRule NewRule
                Else
                    NewRule.SaveWithSystemInfo
                    RefreshRule NewRule
                End If
            End If
        Next lIndex
    End With

ErrExit:
    Set NewRule = Nothing
    Exit Sub
    
ErrSection:
    Set NewRule = Nothing
    RaiseError "frmUpdateLocals.UpdateLocals", eGDRaiseError_Raise
    
End Sub

