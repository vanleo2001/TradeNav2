VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7L.ocx"
Object = "{DCC46394-4B19-11D3-BD95-D426EF2C7949}#1.0#0"; "VSStr7.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmFunctionList 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "List of Available Functions"
   ClientHeight    =   2700
   ClientLeft      =   60
   ClientTop       =   225
   ClientWidth     =   3990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraCategory 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   50
      Width           =   3735
      Begin VB.CommandButton cmdNewFunction 
         Caption         =   "&New Function"
         Height          =   315
         Left            =   2280
         TabIndex        =   5
         Top             =   0
         Width           =   1215
      End
      Begin VB.ComboBox cbCategory 
         Height          =   315
         Left            =   675
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label lblCategory 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Category:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   0
         TabIndex        =   4
         Top             =   60
         Width           =   675
      End
   End
   Begin RichTextLib.RichTextBox rtbDescription 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2280
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   873
      _Version        =   393217
      BackColor       =   -2147483633
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmFunctionList.frx":0000
   End
   Begin VSFlex7LCtl.VSFlexGrid vsFunctionList 
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   3810
      _cx             =   6720
      _cy             =   635
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2340
      Top             =   765
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   11
      ImageHeight     =   12
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFunctionList.frx":0086
            Key             =   "kEqual"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFunctionList.frx":011F
            Key             =   "kFunction"
         EndProperty
      EndProperty
   End
   Begin VSSTR7LibCtl.VSFlexString vsStr 
      Left            =   585
      Top             =   780
      _ConvInfo       =   1
      Text            =   ""
      Pattern         =   ""
      CaseSensitive   =   -1  'True
   End
End
Attribute VB_Name = "frmFunctionList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    bFunctionView As Boolean
    bShowNewFunction As Boolean
End Type
Private m As mPrivate

Public Property Get ShowNewFunction() As Boolean
    ShowNewFunction = m.bShowNewFunction
End Property
Public Property Let ShowNewFunction(ByVal bShowNewFunction As Boolean)
    m.bShowNewFunction = bShowNewFunction
    cmdNewFunction.Visible = m.bShowNewFunction
End Property

Private Sub cbCategory_Click()
On Error GoTo ErrEnd
    
    gEditingArea.GetFunctionList.LoadFilteredFunctionList
    gEditingArea.GetFunctionList.ReLoadFunctions
    
    gEditingArea.RtfTextBoxSetFocus
    

Exit Sub
ErrEnd:
    'lets not worry about this - most likely a chicken before the egg ordeal

End Sub

Private Sub cmdNewFunction_Click()
On Error GoTo ErrSection:

    gEditingArea.NewFunction cbCategory.ItemData(cbCategory.ListIndex)

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "TSOCX.frmFunctionList.cmdNewFunction.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:
    
    If UnloadMode = 0 Then
'        Me.Hide
        gEditingArea.ProcessKeyDown 27                  ' escape
        Cancel = True
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.frmFunctionList.Form.QueryUnload", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit

End Sub

Private Sub Form_Resize()
On Error Resume Next
    
    Dim dFunctionListHeight As Double
    Dim dCategoryBottom As Double
    Const C_MINWIDTH = 3870
    Const C_MINHEIGHT = 1700
    
    dCategoryBottom = fraCategory.Height + fraCategory.Top
    
    If LimitFormSize(Me, C_MINWIDTH, C_MINHEIGHT) Then Exit Sub
 
    ' This will setup the controls for viewing appropriate list...
    If m.bFunctionView Then
        ' FunctionList goes full width, to bottom of combo, top of description...
        dFunctionListHeight = ScaleHeight - fraCategory.Height - rtbDescription.Height
        
        ' Position the description box and the function list...
        rtbDescription.Move 0, ScaleHeight - rtbDescription.Height, ScaleWidth
        With vsFunctionList
            .Redraw = flexRDNone
            .Move 0, dCategoryBottom, ScaleWidth, dFunctionListHeight '.CellHeight * 6
            .Redraw = flexRDDirect
        End With
    Else
        ' FunctionList goes full width, to bottom of combo, top of description...
        dFunctionListHeight = ScaleHeight - rtbDescription.Height '- cbCategory.Height
        
        ' Position the description box and the function list...
        rtbDescription.Move 0, (ScaleHeight - rtbDescription.Height), ScaleWidth
        With vsFunctionList
            .Redraw = flexRDNone
            .Move 0, 0, ScaleWidth, dFunctionListHeight '.CellHeight * 6
            .Redraw = flexRDDirect
        End With
    End If

End Sub

Private Sub vsFunctionList_Click()
On Error GoTo ErrSection:
    
    gEditingArea.GetFunctionList.Click

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.frmFunctionList.vsFunctionList.Click", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit:

End Sub

Private Sub vsFunctionList_DblClick()
On Error GoTo ErrSection:
    
    gEditingArea.ProcessKeyDown 9                       'tab ascii code

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "TSOCX.frmFunctionList.vsFunctionList.DblClick", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit:

End Sub

Private Sub vsFunctionList_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    Select Case KeyCode
        Case 9, 13: gEditingArea.ProcessKeyDown 9       'Tab or Enterkey (select item)
        Case 33: gEditingArea.GetFunctionList.PageUp True
        Case 34: gEditingArea.GetFunctionList.PageDown True
        Case 38: gEditingArea.GetFunctionList.SearchUp True
        Case 40: gEditingArea.GetFunctionList.SearchDown True
        Case 27: gEditingArea.ProcessKeyDown KeyCode    'Esc key
    End Select

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "TSOCX.frmFunctionList.vsFunctionList.KeyDown", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit:

End Sub

Public Sub FunctionView(bFlag As Boolean)
    
    'both label and combo are set to flag
    'lblCategory.Visible = bFlag
    'cbCategory.Visible = bFlag
    m.bFunctionView = bFlag
    fraCategory.Visible = bFlag
    
    'all positioning is in resize -
    Form_Resize

End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:
    
    Const C_DEFAULTHEIGHT = 2715
    
    'lets set the height of ourself
'    Me.Height = (vsFunctionList.CellHeight * 7) + cbCategory.Height + rtbDescription.Height
    Me.Height = C_DEFAULTHEIGHT
    'lets position the widgets
    'category combo goes top left
    'cbCategory.Top = 0
    'category label goes to left of that
    'lblCategory.Top = 50
    fraCategory.Top = 50
    
    ' Show/Hide the new function button as appropriate...
    cmdNewFunction.Visible = m.bShowNewFunction
    
    ' Default to function view...
    FunctionView True

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "TSOCX.frmFunctionList.Form.Load", eGDRaiseError_Show, g.strAppPath
    Resume ErrExit:

End Sub

Private Sub vsFunctionList_SelChange()
On Error Resume Next
    
    gEditingArea.GetFunctionList.ShowDescription

End Sub
