VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{C0F09D2D-1125-11D7-8DA9-0004757A4B66}#1.9#0"; "NavTradeSenseOCXV3.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmHighlightBarReporter 
   Caption         =   "Highlight Bar Reporter"
   ClientHeight    =   4620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   ScaleHeight     =   4620
   ScaleWidth      =   4320
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraEntryExit 
      Height          =   1365
      Left            =   233
      TabIndex        =   4
      Top             =   3128
      Width           =   3855
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
      Caption         =   "frmHighlightBarReporter.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmHighlightBarReporter.frx":0038
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmHighlightBarReporter.frx":0058
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtMaxBars 
         Height          =   285
         Left            =   2640
         TabIndex        =   3
         Top             =   945
         Width           =   735
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmHighlightBarReporter.frx":0074
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
         Alignment       =   1
         ScrollBars      =   0
         PasswordChar    =   ""
         TrapTab         =   0   'False
         EnableContextMenu=   -1  'True
         RaiseChangeEvent=   -1  'True
         Tip             =   "frmHighlightBarReporter.frx":0098
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmHighlightBarReporter.frx":00B8
      End
      Begin HexUniControls.ctlUniRadioXP optShort 
         Height          =   255
         Left            =   2760
         TabIndex        =   6
         Top             =   540
         Width           =   735
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
         Caption         =   "frmHighlightBarReporter.frx":00D4
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmHighlightBarReporter.frx":00FE
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmHighlightBarReporter.frx":011E
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniRadioXP optLong 
         Height          =   255
         Left            =   2760
         TabIndex        =   5
         Top             =   240
         Width           =   735
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
         Caption         =   "frmHighlightBarReporter.frx":013A
         Enabled         =   -1  'True
         Align           =   0
         RadioBackColor  =   -2147483643
         RadioForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   -1  'True
         Tip             =   "frmHighlightBarReporter.frx":0162
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmHighlightBarReporter.frx":0182
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label3 
         Height          =   255
         Left            =   240
         Top             =   960
         Width           =   2415
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
         Caption         =   "frmHighlightBarReporter.frx":019E
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmHighlightBarReporter.frx":01F8
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmHighlightBarReporter.frx":0218
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label2 
         Height          =   255
         Left            =   240
         Top             =   375
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
         Caption         =   "frmHighlightBarReporter.frx":0234
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmHighlightBarReporter.frx":0298
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmHighlightBarReporter.frx":02B8
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3173
      TabIndex        =   2
      Top             =   1448
      Width           =   915
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
      Caption         =   "frmHighlightBarReporter.frx":02D4
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmHighlightBarReporter.frx":0302
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmHighlightBarReporter.frx":0322
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdRun 
      Default         =   -1  'True
      Height          =   375
      Left            =   3173
      TabIndex        =   1
      Top             =   968
      Width           =   915
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
      Caption         =   "frmHighlightBarReporter.frx":033E
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmHighlightBarReporter.frx":0366
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmHighlightBarReporter.frx":0386
      RightToLeft     =   0   'False
   End
   Begin VSFlex7LCtl.VSFlexGrid fgCondition 
      Height          =   2535
      Left            =   233
      TabIndex        =   0
      Top             =   413
      Width           =   2715
      _cx             =   4789
      _cy             =   4471
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
      ScrollBars      =   2
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
   Begin NavTradeSenseV3.Editor Editor1 
      Height          =   315
      Left            =   3240
      TabIndex        =   7
      Top             =   2400
      Visible         =   0   'False
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   556
   End
   Begin HexUniControls.ctlUniLabelXP Label1 
      Height          =   255
      Left            =   233
      Top             =   128
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
      Caption         =   "frmHighlightBarReporter.frx":03A2
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmHighlightBarReporter.frx":0400
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmHighlightBarReporter.frx":0420
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmHighlightBarReporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmHighligtBarReporter.frm
'' Description: Form to show the highlight bar report
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 06/16/2011   MJM         Created
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    Chart As cChart
    iFirstValidRow As Long
End Type
Private m As mPrivate

Public Sub ShowMe()

    Dim i&

    If ActiveChart Is Nothing Then Exit Sub
    Set m.Chart = ActiveChart.Chart
    If m.Chart Is Nothing Then Exit Sub
    If m.Chart.Tree Is Nothing Then Exit Sub

    LoadGrid
    
    If fgCondition.Rows <= 0 Then
        InfBox "There are no highlight bars on chart.", "I", , "Highlight Bar Reporter"
    Else
        ShowForm Me, False, frmMain, , ALT_GRID_ROW_COLOR
        
        'select row with first visible indicator
        With fgCondition
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpFontBold, i, 0) = True Then
                    .Row = i
                    m.iFirstValidRow = i
                    Exit For
                End If
            Next
        End With
    End If

ErrExit:
    If Not Me.Visible Then Unload Me
    Exit Sub

ErrSection:
    RaiseError "frmHighlightBarREport.ShowMe"

End Sub

Private Sub cmdCancel_Click()
On Error Resume Next
    
    Unload Me

End Sub

Private Sub RunReport(Optional Ind As cIndicator = Nothing)
On Error GoTo ErrSection
    
    Dim iBars&, idx&, strEnglish$
    
    Dim oFunction As cFunction
    Dim hbReporter As cHighlightBarReporter
    
    'validate text for max bars
    iBars = Int(ValOfText(txtMaxBars.Text))
    If iBars <= 0 Then
        InfBox "Max number of bars must be a positive integer.", "I"
        Exit Sub
    End If
    
    If Ind Is Nothing Then
        With fgCondition
            If .Row >= .FixedRows And .Row < .Rows And .Col >= .FixedCols And .Col < .Cols Then
                idx = Val(.TextMatrix(.Row, 1))
                If TypeOf m.Chart.Tree(idx) Is cIndicator Then Set Ind = m.Chart.Tree(idx)
            End If
        End With
    End If
    
    If Not Ind Is Nothing Then strEnglish = Ind.Expression

    If Len(strEnglish) > 0 Then
        Me.Hide
        DoEvents
        Set hbReporter = New cHighlightBarReporter
        hbReporter.RunFromBars strEnglish, optLong.Value, m.Chart.Bars, iBars
    Else
        InfBox "Please select a highlight bar from the grid.", "I", , "Highlight Bar Reporter"
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmHighlightBarREport.RunReport"

End Sub

Private Sub cmdRun_Click()
On Error GoTo ErrSection:

    RunReport

ErrExit:
    If Not Me.Visible Then Unload Me
    Exit Sub

ErrSection:
    RaiseError "frmHighlightBarREport.cmdRun_Click"

End Sub

Private Sub fgCondition_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim iRow&, idx&, idxParent&
    Dim Ind As cIndicator
    Dim Pane As cPane
    
    With fgCondition
        If .MouseRow >= .FixedRows And .MouseRow < .Rows Then
            iRow = .MouseRow
            If .Cell(flexcpFontBold, iRow, 0) = False Then
                If InfBox("Highlight bar must be visible on chart. " & _
                          "Would you like to turn this highlight bar on?", "?", "Yes|No", _
                          "Highlight Bar Reporter") = "Y" Then
                
                        idx = Val(.TextMatrix(iRow, 1))
                        
                        If TypeOf m.Chart.Tree(idx) Is cIndicator Then
                            Set Ind = m.Chart.Tree(idx)
                            idxParent = m.Chart.Tree.RelativeIndex(idx, eTREE_Parent)
                            
                            If Not m.Chart.Tree(idxParent) Is Nothing Then
                                'parent of a highlight bar is always an indicator, make sure it is turned on
                                m.Chart.Tree(idxParent).Display = True
                                
                                'now get pane object
                                idxParent = m.Chart.Tree.RelativeIndex(idx, eTREE_Root)
                                If TypeOf m.Chart.Tree(idxParent) Is cPane Then Set Pane = m.Chart.Tree(idxParent)
                                
                                If Not Pane Is Nothing Then
                                    Pane.Display = True
                                    Ind.Display = True
                                    m.Chart.geResetPanes
                                    m.Chart.GenerateChart eRedo3_Settings
                                    .Row = iRow
                                    RunReport Ind
                                End If
                            End If
                        End If
                
                ElseIf m.iFirstValidRow >= .FixedRows And m.iFirstValidRow < .Rows Then
                    .Row = m.iFirstValidRow
                End If
            End If
        Else
            If m.iFirstValidRow >= .FixedRows And m.iFirstValidRow < .Rows Then
                .Row = m.iFirstValidRow
            End If
        End If
    End With

ErrExit:
    If Not Me.Visible Then Unload Me
    Exit Sub

ErrSection:
    RaiseError "frmHighlightBarREport.fgCondition_MouseDown"

End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:
    
    Me.Icon = Picture16(ToolbarIcon("kHBReporter"), , True)
    m.iFirstValidRow = -1
    CenterTheForm Me
    
    g.Styler.StyleForm Me

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmHighlightBarREport.Form_Load"

End Sub

Private Sub LoadGrid()
On Error GoTo ErrSection:

    Dim i&, iParent&, iRow&
    Dim Ind As cIndicator
    Dim Pane As cPane
    
    Dim bVisible As Boolean
    
    SetupGrid fgCondition, eGridMode_List
    
    With fgCondition
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .FixedRows = 0
        .FixedCols = 0
        .Rows = 0
        .Cols = 2           'col1=indicator name, col2=index into tree
        .ColHidden(1) = True
    End With
    
    With m.Chart
        For i = 1 To .Tree.Count
            If TypeOf .Tree(i) Is cIndicator Then
                Set Ind = .Tree(i)
                
                If Ind.DataType = eINDIC_BooleanArray Then
                    bVisible = False        'reset
                    iParent = .Tree.RelativeIndex(i, eTREE_Parent)
                    If Not .Tree(iParent) Is Nothing Then bVisible = .Tree(iParent).Display
                    
                    'parent of a boolean indicator is also an indicator
                    'if parent indicator is visible then check that pane is visible else don't bother
                    If bVisible Then
                        iParent = .Tree.RelativeIndex(i, eTREE_Root)
                        If TypeOf .Tree(iParent) Is cPane Then Set Pane = .Tree(iParent)
                        If Not Pane Is Nothing Then bVisible = Pane.Display
                    End If
                        
                    fgCondition.Rows = fgCondition.Rows + 1
                    iRow = fgCondition.Rows - 1
                    
                    fgCondition.TextMatrix(iRow, 0) = Ind.Name
                    fgCondition.TextMatrix(iRow, 1) = i
                    
                    If bVisible And Ind.Display Then
                        fgCondition.Cell(flexcpForeColor, iRow, 0) = 0      'set to zero to use default grid's forecolor
                        fgCondition.Cell(flexcpFontBold, iRow, 0) = True
                    Else
                        fgCondition.Cell(flexcpForeColor, iRow, 0) = vbGrayText
                        fgCondition.Cell(flexcpFontBold, iRow, 0) = False
                    End If
                End If
            End If
        Next
    End With
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmHighlightBarREport.LoadGrid"

End Sub

