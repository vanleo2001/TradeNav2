VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmTrees 
   Caption         =   "Tree"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11925
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   11925
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniButtonImageXP Corner 
      Height          =   285
      Left            =   10920
      TabIndex        =   6
      Top             =   7485
      Visible         =   0   'False
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
      Caption         =   "frmBuildTrees.frx":0000
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmBuildTrees.frx":002C
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmBuildTrees.frx":004C
      RightToLeft     =   0   'False
   End
   Begin vsOcx6LibCtl.vsElastic vsElastic1 
      Height          =   7725
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   13626
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
      FloodColor      =   192
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      Appearance      =   1
      AutoSizeChildren=   8
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   0   'False
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   3
      GridCols        =   3
      _GridInfo       =   $"frmBuildTrees.frx":0068
      Begin vsOcx6LibCtl.vsIndexTab vsIndexTab1 
         Height          =   1890
         Left            =   90
         TabIndex        =   2
         Top             =   5745
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   3334
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
         Caption         =   "&CodedText|&Editing|&Preview"
         Align           =   0
         Appearance      =   1
         CurrTab         =   2
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
         Begin HexUniControls.ctlUniTextBoxXP txtPreview 
            Height          =   1515
            Left            =   45
            TabIndex        =   5
            Top             =   330
            Width           =   11565
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmBuildTrees.frx":00CE
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
            MultiLine       =   -1  'True
            Alignment       =   0
            ScrollBars      =   0
            PasswordChar    =   ""
            TrapTab         =   0   'False
            EnableContextMenu=   -1  'True
            RaiseChangeEvent=   -1  'True
            Tip             =   "frmBuildTrees.frx":00EE
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmBuildTrees.frx":010E
         End
         Begin HexUniControls.ctlUniTextBoxXP txtEditText 
            Height          =   1515
            Left            =   -12210
            TabIndex        =   4
            Top             =   330
            Width           =   11565
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmBuildTrees.frx":012A
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
            MultiLine       =   -1  'True
            Alignment       =   0
            ScrollBars      =   0
            PasswordChar    =   ""
            TrapTab         =   0   'False
            EnableContextMenu=   -1  'True
            RaiseChangeEvent=   -1  'True
            Tip             =   "frmBuildTrees.frx":014A
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmBuildTrees.frx":016A
         End
         Begin HexUniControls.ctlUniTextBoxXP txtCodedtext 
            Height          =   1515
            Left            =   -12510
            TabIndex        =   3
            Top             =   330
            Width           =   11565
            _ExtentX        =   0
            _ExtentY        =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Enabled         =   -1  'True
            Locked          =   0   'False
            Text            =   "frmBuildTrees.frx":0186
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
            MultiLine       =   -1  'True
            Alignment       =   0
            ScrollBars      =   0
            PasswordChar    =   ""
            TrapTab         =   0   'False
            EnableContextMenu=   -1  'True
            RaiseChangeEvent=   -1  'True
            Tip             =   "frmBuildTrees.frx":01A6
            HideSelection   =   -1  'True
            RightToLeft     =   0   'False
            ManualStart     =   0   'False
            RoundedBorders  =   0   'False
            MousePointer    =   0
            MouseIcon       =   "frmBuildTrees.frx":01C6
         End
      End
      Begin VSFlex7LCtl.VSFlexGrid vsTree 
         Height          =   5595
         Left            =   90
         TabIndex        =   1
         Top             =   90
         Width           =   11655
         _cx             =   20558
         _cy             =   9869
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
         WordWrap        =   -1  'True
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
Attribute VB_Name = "frmTrees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    Trees As cDbTrees
End Type
Private m As mPrivate

Private Const C_RULEID = 1
Private Const C_PhraseType = 2
Private Const C_NodeFrom = 3
Private Const C_NodeTo = 4
Private Const C_Parent = 5
Private Const C_ReturnType = 6
Private Const C_Phrase = 7
Private Const C_Entered = 8
Private Const C_REQ = 9
Private Const C_Error = 10
Private Const C_Class = 11
Private Const C_ClassPtr = 12
Private Const C_Late = 13
Private Const C_PoolAddr = 14
Private Const C_SIZE = 15
Private Const C_Rebuild = 16
Private Const C_FID = 17
Private Const C_FName = 18
Private Const C_INPUTID = 19
Private Const Cols = 20

Property Let Trees(pData As cDbTrees)
    Set m.Trees = pData
End Property
Property Let CodedText(pData As String)
    txtCodedText.Text = pData
End Property
Property Let EditText(pData As String)
    txtEditText.Text = pData
End Property
Property Let Preview(pData As String)
    txtPreview.Text = pData
End Property

Public Sub LoadTrees()
On Error GoTo ErrSection:
    
    Dim X               As Integer

    If m.Trees Is Nothing Then Exit Sub

    'Show results (contents of Tree)
    With vsTree
        .Redraw = flexRDNone
        .Clear
        .AllowSelection = True
        .AllowUserResizing = flexResizeBoth
        .TabBehavior = flexTabCells
        .ExtendLastCol = True
        .ExplorerBar = flexExSortShowAndMove
        .ScrollTrack = True
        .Ellipsis = flexEllipsisEnd
        .Cols = Cols
        .Rows = 1
        .Rows = m.Trees.last + 1
        .FixedCols = 6
        .FixedRows = 1
        .WordWrap = False
        
        .BackColorFixed = vbWhite
        .Cell(flexcpFontBold, 0, .Cols - 1) = True
        .TextMatrix(0, 0) = "Nbr"
        .TextMatrix(0, C_RULEID) = "RuleSeq"
        .TextMatrix(0, C_PhraseType) = "Type"
        .TextMatrix(0, C_NodeFrom) = "NodeF"
        .TextMatrix(0, C_NodeTo) = "NodeT"
        .TextMatrix(0, C_Parent) = "Parent"
        .TextMatrix(0, C_ReturnType) = "RetType"
        .TextMatrix(0, C_Error) = "Err"
        .TextMatrix(0, C_Phrase) = "Phrase"
        .TextMatrix(0, C_Class) = "Class"
        .TextMatrix(0, C_ClassPtr) = "ClassPtr"
        .TextMatrix(0, C_Late) = "Late"
        .TextMatrix(0, C_PoolAddr) = "PoolAddr"
        .TextMatrix(0, C_SIZE) = "Size"
        .TextMatrix(0, C_Rebuild) = "RebuildArray"
        .TextMatrix(0, C_FID) = "FuncID"
        .TextMatrix(0, C_FName) = "FName"
        .TextMatrix(0, C_INPUTID) = "InpID"
        .TextMatrix(0, C_REQ) = "Req"
        .TextMatrix(0, C_Entered) = "Ent"
        .ColAlignment(C_Phrase) = flexAlignLeftTop
        .ColAlignment(Cols - 1) = flexAlignLeftTop
        
        For X = 1 To m.Trees.last
            .TextMatrix(X, 0) = X
            .TextMatrix(X, C_NodeFrom) = m.Trees.NodeFrom(X)
            .TextMatrix(X, C_NodeTo) = m.Trees.NodeTo(X)
            .TextMatrix(X, C_Parent) = m.Trees.Parent(X)
            .TextMatrix(X, C_RULEID) = m.Trees.RuleID(X)
            .TextMatrix(X, C_PhraseType) = m.Trees.GetPhraseTypeDesc(m.Trees.PhraseType(X))
            Select Case m.Trees.ReturnType(X)
                Case kSN_RetNumeric: .TextMatrix(X, C_ReturnType) = "SNbr"
                Case kSN_RetTextSeries: .TextMatrix(X, C_ReturnType) = "SText"
                Case kSN_RetTrueFalse: .TextMatrix(X, C_ReturnType) = "SBool"
                Case kSN_RetNumericConstant: .TextMatrix(X, C_ReturnType) = "CNbr"
                Case kSN_RetText: .TextMatrix(X, C_ReturnType) = "CText"
                Case kSN_RetTrueFalseConstant: .TextMatrix(X, C_ReturnType) = "CBool"
                Case kSN_RetTrades: .TextMatrix(X, C_ReturnType) = "Trades"
                Case kSN_RetBars: .TextMatrix(X, C_ReturnType) = "Bars"
                'Case gRetVariables: .TextMatrix(X, C_ReturnType) = "Vars"
            End Select
            '.TextMatrix(X, C_ReturnType) = m.Trees.ReturnType(X)
            .TextMatrix(X, C_Error) = m.Trees.PhraseError(X)
            .TextMatrix(X, C_REQ) = m.Trees.Req(X)
            .TextMatrix(X, C_Entered) = m.Trees.Entered(X)
            .TextMatrix(X, C_Phrase) = m.Trees.Phrase(X)
            .TextMatrix(X, C_Late) = m.Trees.Late(X)
            .TextMatrix(X, C_Class) = m.Trees.ClassName(X)
            .TextMatrix(X, C_ClassPtr) = m.Trees.ClassPtr(X)
            .TextMatrix(X, C_PoolAddr) = m.Trees.PoolAddr(X)
            .TextMatrix(X, C_SIZE) = m.Trees.Size(X)
            .TextMatrix(X, C_Rebuild) = m.Trees.ReBuildArray(X)
            .TextMatrix(X, C_FID) = m.Trees.FID(X)
            If m.Trees.FID(X) > 0 Then
                .TextMatrix(X, C_FName) = g.Functions.Item(Str(m.Trees.FID(X))).FunctionName
            End If
            .TextMatrix(X, C_INPUTID) = m.Trees.InputID(X)
        Next X
        
        .AutoSize 0, .Cols - 1
        .ColWidth(C_Phrase) = 4000
        .ColWidth(C_Class) = 1000
        .Redraw = flexRDDirect
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTrees.LoadTrees", eGDRaiseError_Raise
    
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
    RaiseError "frmTrees.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strText       As String
    
    CenterTheForm Me
    
    g.Styler.StyleForm Me
    
    strText = GetIniFileProperty("Trees", "", "Placement", g.strIniFile)
    SetFormPlacement Me, strText, "LT"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTrees.Form.Load", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    Me.Hide
    DoEvents

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTrees.Form.QueryUnload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    SetIniFileProperty "Trees", GetFormPlacement(Me), "Placement", g.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmTrees.Form.Unload", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

