VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmConditionBuilder 
   Caption         =   "Condition Builder"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9885
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   9885
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniTextBoxXP txtSelect 
      Height          =   195
      Left            =   540
      TabIndex        =   7
      Top             =   2220
      Visible         =   0   'False
      Width           =   500
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483631
      Enabled         =   0   'False
      Locked          =   0   'False
      Text            =   "frmConditionBuilder.frx":0000
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
      Tip             =   "frmConditionBuilder.frx":002C
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmConditionBuilder.frx":004C
   End
   Begin HexUniControls.ctlUniFrameWL fraNew 
      Height          =   1935
      Left            =   60
      TabIndex        =   1
      Top             =   3420
      Width           =   9015
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
      Caption         =   "frmConditionBuilder.frx":0068
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmConditionBuilder.frx":00B2
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmConditionBuilder.frx":00D2
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtExpression 
         Height          =   1575
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   8775
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   -1  'True
         Text            =   "frmConditionBuilder.frx":00EE
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
         ScrollBars      =   2
         PasswordChar    =   ""
         TrapTab         =   0   'False
         EnableContextMenu=   -1  'True
         RaiseChangeEvent=   -1  'True
         Tip             =   "frmConditionBuilder.frx":010E
         HideSelection   =   0   'False
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmConditionBuilder.frx":012E
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   2565
      Left            =   7620
      TabIndex        =   0
      Top             =   120
      Width           =   1455
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
      Caption         =   "frmConditionBuilder.frx":014A
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmConditionBuilder.frx":016A
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmConditionBuilder.frx":018A
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdCopy 
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   60
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
         Caption         =   "frmConditionBuilder.frx":01A6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmConditionBuilder.frx":01D6
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmConditionBuilder.frx":01F6
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Height          =   375
         Left            =   120
         TabIndex        =   8
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
         Caption         =   "frmConditionBuilder.frx":0212
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmConditionBuilder.frx":0240
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmConditionBuilder.frx":0260
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdAdd 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1140
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
         Caption         =   "frmConditionBuilder.frx":027C
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmConditionBuilder.frx":02AA
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmConditionBuilder.frx":02CA
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdDelete 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1620
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
         Caption         =   "frmConditionBuilder.frx":02E6
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmConditionBuilder.frx":031A
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmConditionBuilder.frx":033A
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdReset 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   2100
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
         Caption         =   "frmConditionBuilder.frx":0356
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmConditionBuilder.frx":038A
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmConditionBuilder.frx":03AA
         RightToLeft     =   0   'False
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid fgExpression 
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   1995
      _cx             =   3519
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
   Begin HexUniControls.ctlUniLabelXP lblDesc 
      Height          =   255
      Left            =   120
      Top             =   60
      Width           =   7095
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
      Caption         =   "frmConditionBuilder.frx":03C6
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmConditionBuilder.frx":04B2
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmConditionBuilder.frx":04D2
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmConditionBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmConditionBuilder.frm
'' Description: Allow the user to create a Trade Sense condition
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 05/01/2013   DAJ         Changed code for loading strategies
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kCols = 6
Private Const kSumColsWidth = 7330  'sum of grid's column widths + 100

Private Const kOHLC = "#9999;Open|#9998;High|#9997;Low|#9996;Close|#9995;MidBar"
Private Const kTrueFalse = "#9994;True|#9993;False"
Private Const kOperators = "= (equals)|<> (not equal)|> (greater than)|>= (greater than or equal)|< (less than)|<= (less than or equal|Is Up (higher than previous bar)|Is Down (lower than previous bar)|Crosses Above|Crosses Below"

Private Enum eGridCols
    eCol_LogicalOp = 0
    eCol_ExprA
    eCol_BarsAgoA
    eCol_Operator
    eCol_ExprB
    eCol_BarsAgoB
End Enum

Private Enum eTblFields     'fields for table holding indicator information
    eFld_IndIdxA = 0        'indicator index of expression A
    eFld_IndIdxB            'indicator index of expression B
    eFld_IsBoolA            'flag for whether indicator in expression A is a boolean
    eFld_CanHaveBarsAgoA
    eFld_CanHaveBarsAgoB
    eFld_IndExprA
    eFld_IndExprB
    eFld_BarsAgoA
    eFld_BarsAgoB
    eFld_Operator
    eFld_LogicalOp
    eFld_OpenParen
    eFld_CloseParen
End Enum

Private Enum eIdxType
    eIdx_Invalid = 9991
    eIdx_IndicatorId
    eIdx_False
    eIdx_True
    eIdx_MidBar
    eIdx_Close
    eIdx_Low
    eIdx_High
    eIdx_Open
End Enum

Public Enum eExprType
    eType_Undefined = 0
    eType_Function
    eType_Scoring
    eType_Criteria
    eType_Clipboard
    eType_Rule
    eType_HighlightBars
    eType_Alert
End Enum

Private Type mPrivate
    SysMgrForm As frmSystemManager
    RuleForm As frmRule
    CriteriaForm As frmCriteria
    ChartEditorForm As frmChartCfg
    
    Chart As cChart
    oAlert As cAlert
    aOperators As New cGdArray
    tblAllItems As New cGdTable
    tblIndInfo As New cGdTable
    tblSystemInfo As New cGdTable   'fields:0=system name,1=system ID,2=lib ID
    nSelSystemIdx As Long           'index into system info table passed back from save-as form
    eExpressionType As eExprType
    strAllItems As String
    strNonBoolItems As String
    strExpression As String
    strAlertName As String
    bPartial As Boolean
    bHideCopyTo As Boolean          '0=show copy to options, 1=don't show copy to options
    bExprBSelected As Boolean       'flag to indicate user selected expression B from drop box
    bInitialShow As Boolean
    astrMacros As cGdArray          ' Array of macro information
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Setup and show the form
'' Inputs:      Chart, Hide Copy?, Expression Type, System Manager, Rule Form,
''              Alert
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMe(Chart As cChart, _
    Optional ByVal bHideCopyTo As Boolean = False, _
    Optional ByVal eType As eExprType = eType_Undefined, _
    Optional frm As Form = Nothing, _
    Optional Alert As cAlert = Nothing)
On Error GoTo ErrSection:
    
    Set m.Chart = Chart
    
    If Not Chart Is Nothing Then
        If Chart.SeasonalCycleTypeEnum <> eCycleType_Undefined Then
            InfBox kSeasonalUnavail, "I", "Ok", "Condition Builder"
            GoTo ErrExit
        End If
    End If
    
    m.bInitialShow = True
    m.bHideCopyTo = bHideCopyTo
    m.eExpressionType = eType
    cmdCopy.Enabled = False
    
    Set m.SysMgrForm = Nothing
    Set m.RuleForm = Nothing
    Set m.CriteriaForm = Nothing
    Set m.ChartEditorForm = Nothing
    
    If Not frm Is Nothing Then
        If TypeOf frm Is frmSystemManager Then
            Set m.SysMgrForm = frm
        ElseIf TypeOf frm Is frmRule Then
            Set m.RuleForm = frm
        ElseIf TypeOf frm Is frmChartCfg Then
            Set m.ChartEditorForm = frm
        ElseIf TypeOf frm Is frmCriteria Then
            Set m.CriteriaForm = frm
        End If
    End If
    
    If eType > eType_Undefined Then
        If Not m.RuleForm Is Nothing Then
            m.nSelSystemIdx = 0
        End If
    End If
            
    CenterFormOnChart Me, m.Chart               '6499
    If eType = eType_Alert Then
        Set m.oAlert = Alert
        If m.oAlert Is Nothing Then m.strAlertName = ""
        cmdCopy.Caption = "&Save"
        cmdCancel.Caption = "&Cancel"
        LoadIndExpr
        If m.Chart Is Nothing Then
            Unload Me
            Exit Sub
        Else
            LoadCondBuilderString
            ShowForm Me, eForm_Modal
        End If
    'ElseIf m.SysMgrForm Is Nothing And m.RuleForm Is Nothing Then
    ElseIf frm Is Nothing Then
        If m.eExpressionType = eType_Rule Or m.eExpressionType = eType_Criteria Then
            'user answered "yes" to toolbox's condition builder prompt
            cmdCopy.Caption = "&Done"
            cmdCancel.Caption = "&Cancel"
        Else
            cmdCopy.Caption = "Save &As"
            cmdCancel.Caption = "&Cancel"
        End If
        ShowForm Me, False, frmMain
    Else
        cmdCopy.Caption = "&Done"
        cmdCancel.Caption = "&Cancel"
        If Not m.SysMgrForm Is Nothing Then
            ShowForm Me, True, m.SysMgrForm
        ElseIf Not m.RuleForm Is Nothing Then
            ShowForm Me, True, m.RuleForm
        ElseIf Not m.CriteriaForm Is Nothing Then
            ShowForm Me, True, m.CriteriaForm
        ElseIf Not m.ChartEditorForm Is Nothing Then
            ShowForm Me, False, m.ChartEditorForm
        Else
            Unload Me
        End If
    End If
                    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConditionBuilder.ShowMe"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdAdd_Click
'' Description: Allow user to add a row to the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdAdd_Click()
On Error GoTo ErrSection:

    Dim bCancel As Boolean              ' Cancel grid edit?
    
    AddRow
    fgExpression_BeforeEdit fgExpression.Rows - 1, eCol_ExprA, bCancel
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConditionBuilder.cmdAdd_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Unload form without saving
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    ' Let the system manager know that the user has cancelled this operation...
    If Not m.SysMgrForm Is Nothing Then
        m.SysMgrForm.CondBuilderExpr = "UserCancel"
    End If
    
    Unload Me
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConditionBuilder.cmdCancel_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCopy_Click
'' Description: Allow user to copy the expression to various places
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCopy_Click()
On Error GoTo ErrSection:
        
    Dim strSymbol$
    Dim Expr As cExpression
    Dim oFunction As cFunction
    Dim bHasAssignment As Boolean
    
    If Not m.Chart Is Nothing Then strSymbol = m.Chart.Symbol
    
    If m.eExpressionType = eType_Alert Then
        Me.Hide
        m.bPartial = False
        If m.oAlert Is Nothing Then
            If Len(m.strAlertName) = 0 Then
                m.strAlertName = AskBox("h=Name for new Alert ; i=? ; g=string ; d=" & strSymbol & " Chart Alert:")
            End If
            CreateNew
        ElseIf Not m.oAlert.Indicator Is Nothing Then
            If BuildExpression Then
                Set Expr = New cExpression          'aardvark 3650 fix
                Set oFunction = New cFunction
                With Expr
                    .PortfolioNavigator = False
                    .Functions = g.Functions
                    .ValidateFunctionRule m.strExpression
                    oFunction.CodedText = .CodedText
                    oFunction.DataTypeID = .FunctionReturnType
                    oFunction.ReturnTypeID = .FunctionReturnType
                    oFunction.LateCalculating = .LateCondition
                End With
                m.oAlert.Indicator.CodedText = oFunction.CodedText
                m.oAlert.Indicator.Expression = m.strExpression
                m.oAlert.Indicator.InitAlert m.Chart, True
                SaveCondBuilderString m.oAlert.Indicator
            End If
            Set Expr = Nothing
            Set oFunction = Nothing
        End If
        Unload Me
    ElseIf Not m.SysMgrForm Is Nothing Then
        ' if came from Strategy Editor, pass back instead of bringing up rule form
        m.bPartial = False
        BuildExpression
        m.SysMgrForm.CondBuilderExpr = m.strExpression
        Unload Me
    ElseIf Not m.RuleForm Is Nothing Then
        m.bPartial = False
        BuildExpression
        m.RuleForm.CondBuilderExpr = m.strExpression
        Unload Me
    ElseIf Not m.CriteriaForm Is Nothing Then
        m.bPartial = False
        BuildExpression
        m.CriteriaForm.CondBuilderExpr = m.strExpression
        Unload Me
    ElseIf cmdCopy.Caption = "Save &As" Then
        bHasAssignment = (InStr(m.strExpression, ":=") <> 0)
        frmConditionSave.ShowMe m.eExpressionType, m.tblSystemInfo, m.nSelSystemIdx, m.strAlertName, strSymbol, bHasAssignment
        If m.eExpressionType <> eType_Undefined Then
            Me.Hide
            CreateNew
            If m.eExpressionType = eType_Alert Then
                Unload Me
            End If
        End If
    ElseIf m.eExpressionType <> eType_Undefined Or Not m.ChartEditorForm Is Nothing Then
        CreateNew
        Unload Me
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConditionBuilder.cmdCopy_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdDelete_Click
'' Description: Allow user to delete a row from the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdDelete_Click()
On Error GoTo ErrSection:

    txtSelect.Visible = False
    DeleteRow
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConditionBuilder.cmdDelete_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdReset_Click
'' Description: Allow user to reset the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdReset_Click()
On Error GoTo ErrSection:

    txtSelect.Visible = False
    m.strExpression = ""
    txtExpression.Text = ""
    fgExpression.Rows = fgExpression.FixedRows
    InitGrid
    cmdCopy.Enabled = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConditionBuilder.cmdReset_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgExpression_AfterEdit
'' Description: Perform operations after the user is done editing
'' Inputs:      Row, Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgExpression_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim strText$, strExprB$
    Dim bBuild As Boolean
        
    m.bPartial = True
    
    With fgExpression
        'drop down next combo column
        Select Case Col
            Case eCol_LogicalOp
                strText = .TextMatrix(Row, eCol_LogicalOp)
                .Col = eCol_ExprA
                If .Row <> .Rows - 1 Then bBuild = True
            Case eCol_ExprA
                strText = .TextMatrix(Row, eCol_ExprA)
                If Len(.TextMatrix(Row, eCol_Operator)) = 0 Then
                    .Col = eCol_Operator
                Else
                    .Col = eCol_ExprB
                End If
            Case eCol_Operator
                .TextMatrix(Row, eCol_Operator) = m.aOperators(.ComboIndex)
                strText = .TextMatrix(Row, eCol_Operator)
                If strText = "Is Up" Or strText = "Is Down" Then
                    bBuild = True
                    If Row = .Rows - 1 Then AddRow
                    strText = ""
                Else
                    .Col = eCol_ExprB
                End If
            Case eCol_ExprB
                'note: if user type in a value, the ComboCloseUp event does not trigger
                strText = .TextMatrix(Row, eCol_ExprB)
                If m.bExprBSelected Then
                    .Col = eCol_LogicalOp
                    strText = ""
                    bBuild = True
                    If .Row = .Rows - 1 Then AddRow
                ElseIf Len(strText) > 0 Then
                    If Left(strText, 1) = "+" Or Left(strText, 1) = "-" Then
                        strExprB = Right(strText, Len(strText) - 1)
                    Else
                        strExprB = strText
                    End If
                    strText = ""
                    If IsNumeric(strExprB) Then
                        .Col = eCol_LogicalOp
                        If .Row = .Rows - 1 Then AddRow
                        bBuild = True
                    Else
                        InfBox "Please enter a number or select an expression from the list.", "i"
                        .Col = eCol_ExprB
                        strText = " "   'set to trigger drop down
                    End If
                End If
            Case eCol_BarsAgoA
                bBuild = True
            Case eCol_BarsAgoB
                bBuild = True
        End Select
        
        If bBuild Then
            BuildExpression
            If m.eExpressionType = eType_Scoring Then
                txtExpression = ""
            Else
                txtExpression = m.strExpression
            End If
        End If
        
        If Len(strText) > 0 Then
            fgExpression.EditCell
            SendKeys "{F4}"
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConditionBuilder.fgExpression_AfterEdit"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgExpression_BeforeEdit
'' Description: Perform operations when the user starts to edit the grid
'' Inputs:      Row, Column, Cancel the Edit?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgExpression_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim strExpr$, strOp$
           
    txtSelect.Visible = False
    cmdCopy.Enabled = Len(m.strExpression)
    With fgExpression
        If .Row = 0 Then
            Cancel = True
        ElseIf .Col = eCol_LogicalOp Then
            If .Row > .FixedRows Then
                If .MergeRow(.Row) = True Then
                    Cancel = True
                Else
                    .ComboList = "And|Or"
                End If
            Else
                Cancel = True
            End If
        ElseIf .Col = eCol_BarsAgoA Then
            .ComboList = ""
            If m.tblIndInfo(eFld_CanHaveBarsAgoA, .Row) = 0 Then Cancel = True
        ElseIf .Col = eCol_BarsAgoB Then
            .ComboList = ""
            strExpr = .TextMatrix(.Row, eCol_ExprB)
            strOp = .TextMatrix(.Row, eCol_Operator)
            If strExpr = "True" Or strExpr = "True" Or _
                strOp = "Is Up" Or strOp = "Is Down" Or _
                (m.tblIndInfo(eFld_CanHaveBarsAgoB, .Row) = 0) Then
                Cancel = True
            End If
        ElseIf .Col = eCol_ExprA Then
            .ComboList = m.strAllItems
'            If Len(.TextMatrix(.Row, .Col)) = 0 Then
'                txtClickToAdd.Move .CellLeft + 250, .CellTop + 170
'                txtClickToAdd.Visible = True
'            End If
        ElseIf .Col = eCol_ExprB Then
            m.bExprBSelected = False
            strOp = .TextMatrix(.Row, eCol_Operator)
            If strOp = "Is Up" Or strOp = "Is Down" Then
                Cancel = True
            ElseIf m.tblIndInfo(eFld_IsBoolA, .Row) = 1 Then
                .ComboList = kTrueFalse
            Else
                .ComboList = "|" & m.strNonBoolItems
            End If
        ElseIf .Col = eCol_Operator Then
            If m.tblIndInfo(eFld_IsBoolA, .Row) = 1 Then
                Cancel = True
            Else
                .ComboList = kOperators     '"=|>|>=|<|<="
            End If
        End If
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConditionBuilder.fgExpression_BeforeEdit"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgExpression_Click
'' Description: Perform operations when user clicks in the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgExpression_Click()
On Error GoTo ErrSection:

    Dim lMouseRow As Long
    Dim strText$

    With fgExpression
        lMouseRow = .MouseRow
        If lMouseRow >= .FixedRows Then
            If .MergeRow(lMouseRow) = True Then
                .Cell(flexcpText, lMouseRow, 0, lMouseRow, .Cols - 1) = ""
                .MergeRow(lMouseRow) = False
                .Col = eCol_LogicalOp
                .EditCell
                SendKeys "{F4}"
            ElseIf .Rows > .FixedRows + 1 Then
                strText = .TextMatrix(.Rows - 1, eCol_LogicalOp)
                If strText <> "And" And strText <> "Or" Then
                    .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, .Cols - 1) = "Click here to add another row ..."
                    .MergeCells = flexMergeRestrictRows
                    .MergeRow(.Rows - 1) = True
                End If
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConditionBuilder.fgExpression_Click"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgExpression_ComboCloseUp
'' Description: Perform operations when user closes combo box in grid
'' Inputs:      Row, Column, Finish Edit?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgExpression_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
On Error GoTo ErrSection:

    Dim Ind As cIndicator
    Dim strVal$

    txtSelect.Visible = False
    With fgExpression
        If .Col = eCol_ExprA And Len(.ComboData(.ComboIndex)) > 0 Then
            m.tblIndInfo(eFld_IndIdxA, .Row) = .ComboData(.ComboIndex)
            If Validate(.Row, 0, Ind) Then
                strVal = .TextMatrix(.Row, eCol_ExprB)
                If m.tblIndInfo(eFld_IsBoolA, .Row) = 1 Then
                    .TextMatrix(.Row, eCol_Operator) = "="
                    If strVal <> "True" And strVal <> "False" Then .TextMatrix(.Row, eCol_ExprB) = ""
                ElseIf strVal = "True" Or strVal = "False" Then
                    .TextMatrix(.Row, eCol_ExprB) = ""
                End If
            Else
                MsgBox "Validation failed. Please try again."
                .TextMatrix(.Row, .Col) = ""
            End If
        ElseIf .Col = eCol_ExprB And Len(.ComboData(.ComboIndex)) > 0 Then
            m.tblIndInfo(eFld_IndIdxB, .Row) = .ComboData(.ComboIndex)
            If Validate(.Row, 1, Ind) Then
                m.bExprBSelected = True
            Else
                MsgBox "Validation failed. Please try again."
                .TextMatrix(.Row, .Col) = ""
            End If
        End If
    End With
    
    FinishEdit = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConditionBuilder.fgExpression_ComboCloseUp"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgExpression_ComboDropDown
'' Description: Perform operations when user drops down combo in the grid
'' Inputs:      Row, Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgExpression_ComboDropDown(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    With fgExpression
        If Len(.EditText) = 0 Then
            If .Col = eCol_ExprB Then
                .EditText = "Select (or type a value)..."
                SendKeys "+{END}"
            Else
                txtSelect.Move .CellLeft + .Left + .Width - .ClientWidth, .CellTop + .Top + .Height - .ClientHeight
                txtSelect.Visible = True
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConditionBuilder.fgExpression_ComboDropDown"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Activate
'' Description: Perform some operations when the form gets the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Activate()
On Error GoTo ErrSection:

    If m.bInitialShow Then
        With fgExpression
            .Row = .FixedRows
            .Col = eCol_ExprA
            .EditCell
        End With
        SendKeys "{F4}"
        m.bInitialShow = False
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConditionBuilder.Form_Activate"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_KeyDown
'' Description: If the user presses F1 on the form, show the help
'' Inputs:      Code of Key Pressed, Shift/Ctrl/Alt Status
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
    RaiseError "frmConditionBuilder.Form_KeyDown"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Setup and intialize the form when it is loaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Me.Icon = Picture16(ToolbarIcon("ID_ConditionBuilder"), , True)
    
    g.Styler.StyleForm Me
    
    m.tblAllItems.CreateField eGDARRAY_Strings, 0
    m.tblAllItems.CreateField eGDARRAY_Strings, 1
    m.tblAllItems.CreateField eGDARRAY_Strings, 2
    
    InitOperatorsArray
    InitChartItemsList
    InitSystemsList
    InitGrid
    
    Set m.astrMacros = New cGdArray
    m.astrMacros.Create eGDARRAY_Strings
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConditionBuilder.Form_Load"

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

    If LimitFormSize(Me, lblDesc.Left + lblDesc.Width + fraButtons.Width, fraButtons.Top + fraButtons.Height + fraNew.Height) Then Exit Sub

    fraButtons.Left = Me.ScaleWidth - fraButtons.Width
    
    With fraNew
        .Move .Left, Me.ScaleHeight - .Height - .Left, Me.ScaleWidth - .Left * 2
    End With
    
    With txtExpression
        .Move .Left, .Top, fraNew.Width - 250, fraNew.Height - 350
    End With
    
    With fgExpression
        .Move .Left, .Top, fraButtons.Left - .Left, fraNew.Top - .Top - .Left
    End With

    'Frame1.Move fraNew.Left, 0, kSumColsWidth, Me.ScaleHeight - fraNew.Height
    'fgExpression.Move txtExpression.Left, 100, kSumColsWidth, Me.ScaleHeight - fraNew.Height

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Perform operations when user form is unloaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    Set m.Chart = Nothing
    Set m.tblIndInfo = Nothing
    Set m.tblSystemInfo = Nothing
    Set m.tblAllItems = Nothing
    Set m.aOperators = Nothing
    
    If Not m.ChartEditorForm Is Nothing Then m.ChartEditorForm.tmrChartCfg.Enabled = True

    Set m.RuleForm = Nothing
    Set m.SysMgrForm = Nothing
    Set m.ChartEditorForm = Nothing
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConditionBuilder.Form_Unload"

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

    With fgExpression
        .Redraw = flexRDNone
        SetupGrid Me.fgExpression, eGridMode_Grid
        .Sort = flexSortNone
        .ExplorerBar = flexExSortShow
        .SelectionMode = flexSelectionFree
        .FixedCols = 0
        .Editable = flexEDKbdMouse
        .FixedRows = 1
        .Rows = 2
        .Cols = kCols
        'column headers
        .TextMatrix(0, eCol_LogicalOp) = "And/Or"
        .TextMatrix(0, eCol_ExprA) = "Indicator" ' "Expression(A)"
        .TextMatrix(0, eCol_BarsAgoA) = "Bars Ago"
        .TextMatrix(0, eCol_Operator) = "Operator"
        .TextMatrix(0, eCol_ExprB) = "Indicator or Value" '"Expression(B)"
        .TextMatrix(0, eCol_BarsAgoB) = "Bars Ago"
        'column widths
        .ColWidth(eCol_LogicalOp) = 750
        .ColWidth(eCol_ExprA) = 2250
        .ColWidth(eCol_ExprB) = 2250
        .ColWidth(eCol_BarsAgoA) = 990
        .ColWidth(eCol_BarsAgoB) = 990
        .ColWidth(eCol_Operator) = 1000
        .Cell(flexcpBackColor, .FixedRows, 0) = Me.BackColor
        .Redraw = flexRDBuffered
    End With
    
    'initialize first 2 records of indicator info table
    m.tblIndInfo.Clear
    m.tblIndInfo.AddRecord "-1,-1,0,0,0,A,B,=,0,0, , , ", -1, ","
    m.tblIndInfo.AddRecord "-1,-1,0,0,0,A,B,=,0,0, , , ", -1, ","

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConditionBuilder.InitGrid"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitChartItemsList
'' Description: Initialize the chart items list
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitChartItemsList()
On Error GoTo ErrSection:

    Dim i&, strLabel$, strEnglishText$
    Dim Ind As cIndicator
    Dim Pane As cPane
            
    If m.Chart Is Nothing Then Exit Sub
    
    m.tblAllItems.NumRecords = 0
    m.strAllItems = kOHLC
    m.strNonBoolItems = kOHLC
    
    m.tblAllItems.AddRecord ""
    m.tblAllItems(0, 0) = "|#9999*1;"
    m.tblAllItems(1, 0) = "And"
    m.tblAllItems(2, 0) = "Open"
    
    m.tblAllItems.AddRecord ""
    m.tblAllItems(0, 1) = "|#9998;"
    m.tblAllItems(1, 1) = "And"
    m.tblAllItems(2, 1) = "High"
    
    m.tblAllItems.AddRecord ""
    m.tblAllItems(0, 2) = "|#9997;"
    m.tblAllItems(1, 2) = "And"
    m.tblAllItems(2, 2) = "Low"
    
    m.tblAllItems.AddRecord ""
    m.tblAllItems(0, 3) = "|#9996;"
    m.tblAllItems(1, 3) = "And"
    m.tblAllItems(2, 3) = "Close"
    
    m.tblAllItems.AddRecord ""
    m.tblAllItems(0, 4) = "|#9995;"
    m.tblAllItems(1, 4) = "And"
    m.tblAllItems(2, 4) = "MidBar"
        

    For i = 1 To m.Chart.Tree.Count
        If m.Chart.Tree.NodeLevel(i) > 0 Then
            Set Ind = m.Chart.Tree(i)
            If Not Ind Is Nothing Then
                Set Pane = m.Chart.Tree(Ind.geIndpaneId)
                If Not Pane Is Nothing Then
                    If Pane.Display And Ind.Display And _
                       (Ind.DataType = eINDIC_Array Or Ind.DataType = eINDIC_BooleanArray) Then
                        If Ind.DataType <> eINDIC_BooleanArray Then
                            m.strNonBoolItems = m.strNonBoolItems & "|#" & Str(Ind.geIndId) & ";" & Ind.ChartLabel
                        End If
                        m.strAllItems = m.strAllItems & "|#" & Str(Ind.geIndId) & ";" & Ind.ChartLabel
                        m.tblAllItems.AddRecord ""
                        m.tblAllItems(0, m.tblAllItems.NumRecords - 1) = "|#" & Str(Ind.geIndId) & ";"
                        m.tblAllItems(1, m.tblAllItems.NumRecords - 1) = "And"
                        m.tblAllItems(2, m.tblAllItems.NumRecords - 1) = Ind.ChartLabel
                    End If
                End If
            End If
        End If
    Next

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConditionBuilder.InitChartItemsList"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ValidateIdx
'' Description: Validate the index
'' Inputs:      Index, Indicator Found
'' Returns:     Index Type
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidateIdx(ByVal nIdx&, IndFound As cIndicator) As eIdxType
On Error GoTo ErrSection:

    Dim eRC As eIdxType
    Dim Ind As cIndicator
    
    eRC = eIdx_Invalid
    
    Select Case nIdx
        Case eIdx_Open
            eRC = eIdx_Open
        Case eIdx_High
            eRC = eIdx_High
        Case eIdx_Low
            eRC = eIdx_Low
        Case eIdx_Close
            eRC = eIdx_Close
        Case eIdx_MidBar
            eRC = eIdx_MidBar
        Case eIdx_True
            eRC = eIdx_True
        Case eIdx_False
            eRC = eIdx_False
    End Select
    
    If eRC <> eIdx_Invalid Then
        ValidateIdx = eRC
        Exit Function
    End If
    
    If m.Chart.Tree.NodeLevel(nIdx) >= 1 Then
        Set Ind = m.Chart.Tree(nIdx)
        If Not Ind Is Nothing Then
            eRC = eIdx_IndicatorId
            Set IndFound = Ind
        End If
    End If
    
    ValidateIdx = eRC

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmConditionBuilder.ValidateIdx"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Validate
'' Description: Validate the expression for the record
'' Inputs:      Record Number, Expression, Indicator Found
'' Returns:     True if success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function Validate(ByVal nRecord&, ByVal nExpression&, IndFound As cIndicator) As Boolean
On Error GoTo ErrSection:

    Dim nIdx&, nFunctionID&
    Dim eIndexType As eIdxType
    Dim IndFunction As cFunction
        
    If nRecord < fgExpression.FixedRows Or nRecord >= m.tblIndInfo.NumRecords Then
        Validate = False       'something VERY wrong here!
        Exit Function
    End If
        
    If nExpression = 0 Then
        nIdx = Val(m.tblIndInfo(eFld_IndIdxA, nRecord))
    Else
        nIdx = Val(m.tblIndInfo(eFld_IndIdxB, nRecord))
    End If
    eIndexType = ValidateIdx(nIdx, IndFound)
    
    If eIndexType = eIdx_Invalid Then
        If nExpression = 0 Then
            m.tblIndInfo(eFld_IndIdxA, nRecord) = -1
            m.tblIndInfo(eFld_IsBoolA, nRecord) = 0
            m.tblIndInfo(eFld_BarsAgoA, nRecord) = 0
            m.tblIndInfo(eFld_IndExprA, nRecord) = "A"
            m.tblIndInfo(eFld_CanHaveBarsAgoA, nRecord) = 0
        Else
            m.tblIndInfo(eFld_IndIdxB, nRecord) = -1
            m.tblIndInfo(eFld_BarsAgoB, nRecord) = 0
            m.tblIndInfo(eFld_IndExprB, nRecord) = "B"
            m.tblIndInfo(eFld_CanHaveBarsAgoB, nRecord) = 0
        End If
        If Not m.bPartial Then
            Validate = False
            Exit Function
        End If
    End If
    
    Dim Expr As cExpression
    Dim strTest$
    
    If nExpression = 0 Then
        m.tblIndInfo(eFld_IsBoolA, nRecord) = 0
        m.tblIndInfo(eFld_CanHaveBarsAgoA, nRecord) = 0
        If Not IndFound Is Nothing Then
            If IndFound.DataType = eINDIC_BooleanArray Then
                m.tblIndInfo(eFld_IsBoolA, nRecord) = 1
            End If
        End If
    Else
        m.tblIndInfo(eFld_CanHaveBarsAgoB, nRecord) = 0
    End If
    
    If IndFound Is Nothing Then
        If eIndexType = eIdx_Open Or eIndexType = eIdx_High Or eIndexType = eIdx_Low Or eIndexType = eIdx_Close _
            Or eIndexType = eIdx_MidBar Then
            
            If nExpression = 0 Then
                m.tblIndInfo(eFld_CanHaveBarsAgoA, nRecord) = 1
            Else
                m.tblIndInfo(eFld_CanHaveBarsAgoB, nRecord) = 1
            End If
            
        End If
    Else
        strTest = "(" & IndFound.Expression & ")." & 1
        Set Expr = New cExpression
        With Expr
            .PortfolioNavigator = False
            .Functions = g.Functions
            .ValidateFunctionRule strTest
        End With
        If nExpression = 0 Then
            m.tblIndInfo(eFld_CanHaveBarsAgoA, nRecord) = 1
        Else
            m.tblIndInfo(eFld_CanHaveBarsAgoB, nRecord) = 1
        End If
    End If
        
    Set IndFunction = Nothing
    Validate = True
    
ErrExit:
    Set Expr = Nothing
    Exit Function

ErrSection:
    'TradeSense error occurred...
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        If nExpression = 0 Then
            m.tblIndInfo(eFld_CanHaveBarsAgoA, nRecord) = 0
            fgExpression.TextMatrix(nRecord, eCol_BarsAgoA) = ""
        Else
            m.tblIndInfo(eFld_CanHaveBarsAgoB, nRecord) = 0
            fgExpression.TextMatrix(nRecord, eCol_BarsAgoB) = ""
        End If
        Set Expr = Nothing
        Validate = True
    Else
        Set Expr = Nothing
        RaiseError "frmConditionBuilder.Validate"
    End If

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddRow
'' Description: Add a row to the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddRow()
On Error GoTo ErrSection:

    Dim bExit As Boolean
    Dim strOp$
    
    If fgExpression.Rows <> m.tblIndInfo.NumRecords Then Exit Sub
        
    With fgExpression
        If .Rows > .FixedRows Then
            strOp = .TextMatrix(.Rows - 1, eCol_Operator)
            If .TextMatrix(.Rows - 1, eCol_ExprA) = "" Or strOp = "" Then
                bExit = True
            ElseIf strOp <> "Is Up" And strOp <> "Is Down" And .TextMatrix(.Rows - 1, eCol_ExprB) = "" Then
                bExit = True
            End If
        End If
    End With
    
    If bExit Then
        MsgBox ("The last row is incomplete, please complete the last row before adding a new one.")
        Exit Sub
    End If
    
    m.tblIndInfo.AddRecord "-1,-1,0,0,0,A,B,=,0,0", -1, ","
    
    With fgExpression
        .Rows = .Rows + 1
        .Row = .Rows - 1
        .Col = eCol_LogicalOp
        
        .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, .Cols - 1) = "Click here to add another row ..."
        .MergeCells = flexMergeRestrictRows
        .MergeRow(.Rows - 1) = True
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConditionBuilder.AddRow"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeleteRow
'' Description: Delete a row from the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DeleteRow()
On Error GoTo ErrSection:
    
    Dim nRowToDelete&
    
    With fgExpression
        If .Row < 0 Then
            MsgBox "Please select the row you want to delete."
        ElseIf .Row = 0 Then
            'header row (do nothing)
        Else
            nRowToDelete = .Row
            .RemoveItem nRowToDelete
            m.tblIndInfo.RemoveRecords nRowToDelete
            If .Row = .FixedRows Then .TextMatrix(.Row, eCol_LogicalOp) = ""
            BuildExpression
            txtExpression.Text = m.strExpression
       End If
       If .Rows = .FixedRows Then cmdCopy.Enabled = False
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConditionBuilder.DeleteRow"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ValidateBarsAgo
'' Description: Validate the number of Bars Ago
'' Inputs:      Row, Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ValidateBarsAgo(ByVal nRow, ByVal eCol As eGridCols) As Boolean
On Error GoTo ErrSection:
    
    Dim strVal$, bRC As Boolean
    
    bRC = True
    strVal = fgExpression.TextMatrix(nRow, eCol)
    If strVal <> "" Then
        If Not IsNumeric(strVal) Then bRC = False
    End If
    
    ValidateBarsAgo = bRC
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmConditionBuilder.ValidateBarsAgo"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetConstExpr
'' Description: Set constant expression
'' Inputs:      Record, Expression
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetConstExpr(ByVal nRecord&, ByVal nExpression&)
On Error GoTo ErrSection:

    Dim strExpr$, strBarsAgo$
    Dim strOp$, strLogicalOp$
    Dim nExprCol&, nBarCol&
    Dim nExprField&, nBarField&
    Dim nBarsAgo&
    
    If nExpression = 0 Then
        nExprCol = eCol_ExprA
        nBarCol = eCol_BarsAgoA
        nExprField = eFld_IndExprA
        nBarField = eFld_BarsAgoA
    Else
        nExprCol = eCol_ExprB
        nBarCol = eCol_BarsAgoB
        nExprField = eFld_IndExprB
        nBarField = eFld_BarsAgoB
    End If
        
    strLogicalOp = ""
    With fgExpression
        strExpr = .TextMatrix(nRecord, nExprCol)
        strLogicalOp = .TextMatrix(nRecord, eCol_LogicalOp)
        strOp = .TextMatrix(nRecord, eCol_Operator)
        If nExpression = 0 Then
            'user may have deleted the top row so double check and remove logical operator if necessary
            If Len(strLogicalOp) > 0 And nRecord = .FixedRows Then
                strLogicalOp = ""
            End If
        End If
        If strExpr = "" And (InStr(strOp, "Is Up") Or InStr(strOp, "Is Down")) Then
            strExpr = .TextMatrix(nRecord, eCol_ExprA)
            nBarsAgo = 1
            m.tblIndInfo(eFld_IndIdxB, nRecord) = m.tblIndInfo(eFld_IndIdxA, nRecord)
            If InStr(strOp, "Is Up") Then
                strOp = ">"
            Else
                strOp = "<"
            End If
        Else
            strBarsAgo = .TextMatrix(nRecord, nBarCol)
            If Len(strBarsAgo) > 0 Then nBarsAgo = Val(strBarsAgo)
        End If
        If strExpr = "MidBar" Then strExpr = "AvgHL"
        m.tblIndInfo(nExprField, nRecord) = strExpr
        m.tblIndInfo(nBarField, nRecord) = nBarsAgo
        m.tblIndInfo(eFld_Operator, nRecord) = strOp
        If Len(strLogicalOp) > 0 Then m.tblIndInfo(eFld_LogicalOp, nRecord) = strLogicalOp
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConditionBuilder.SetConstExpr"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CopyExprToRecord
'' Description: Copy expression to record
'' Inputs:      Record, Expression, Indicator
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CopyExprToRecord(ByVal nRecord&, ByVal nExpression&, Ind As cIndicator)
On Error GoTo ErrSection:

    Dim strExpr$, strSub$
    Dim strOp$, strLogicalOp$
    Dim nBarsAgo&
    
    strLogicalOp = ""
    'set expression and bar ago values
    If Ind Is Nothing Then
        SetConstExpr nRecord, nExpression
    Else
        If nExpression = 0 Then
            strExpr = fgExpression.TextMatrix(nRecord, eCol_ExprA)
            strLogicalOp = fgExpression.TextMatrix(nRecord, eCol_LogicalOp)
            If Len(strLogicalOp) > 0 And nRecord = fgExpression.FixedRows Then
                strLogicalOp = ""
            End If
        Else
            strExpr = fgExpression.TextMatrix(nRecord, eCol_ExprB)
        End If
        strSub = Right(strExpr, 2)
        strExpr = Ind.Expression
        If Left(strSub, 1) = "." Then   'extract bars ago from indicator's expression
            strExpr = Mid(strExpr, 1, Len(strExpr) - 2)
            nBarsAgo = Val(Right(strSub, 1))
        End If
        strOp = fgExpression.TextMatrix(nRecord, eCol_Operator)
        If InStr(strOp, "Is Up") > 0 Then
            strOp = ">"
            If nExpression <> 0 Then nBarsAgo = 1
        ElseIf InStr(strOp, "Is Down") > 0 Then
            strOp = "<"
            If nExpression <> 0 Then nBarsAgo = 1
        End If
        'set/override bars ago if specified by user
        With fgExpression
            If nExpression = 0 Then
                If .TextMatrix(nRecord, eCol_BarsAgoA) <> "" Then nBarsAgo = Val(.TextMatrix(nRecord, eCol_BarsAgoA))
            Else
                If .TextMatrix(nRecord, eCol_BarsAgoB) <> "" Then nBarsAgo = Val(.TextMatrix(nRecord, eCol_BarsAgoB))
            End If
        End With
        If nExpression = 0 Then
            m.tblIndInfo(eFld_IndExprA, nRecord) = strExpr
            m.tblIndInfo(eFld_BarsAgoA, nRecord) = nBarsAgo
            If Ind.DataType = eINDIC_BooleanArray Then
                m.tblIndInfo(eFld_IsBoolA, nRecord) = 1
            Else
                m.tblIndInfo(eFld_IsBoolA, nRecord) = 0
            End If
        Else
            m.tblIndInfo(eFld_IndExprB, nRecord) = strExpr
            m.tblIndInfo(eFld_BarsAgoB, nRecord) = nBarsAgo
        End If
        m.tblIndInfo(eFld_Operator, nRecord) = strOp
        If Len(strLogicalOp) > 0 Then m.tblIndInfo(eFld_LogicalOp, nRecord) = strLogicalOp
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConditionBuilder.CopyExprToRecord"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CopyToTable
'' Description: Copy the grid to the table
'' Inputs:      None
'' Returns:     True if success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CopyToTable() As Boolean
On Error GoTo ErrSection:

    Dim i&, nEnd&
    Dim bRC As Boolean
    Dim IndA As cIndicator, IndB As cIndicator
    
    If fgExpression.Rows <> m.tblIndInfo.NumRecords Then
        CopyToTable = False         'something VERY wrong here
        Exit Function
    End If
    
    bRC = True
    With fgExpression
        If .MergeRow(.Rows - 1) = True Then
            nEnd = .Rows - 2
        Else
            nEnd = .Rows - 1
        End If
        For i = .FixedRows To nEnd
            Set IndA = Nothing
            Set IndB = Nothing
            bRC = ValidateBarsAgo(i, eCol_BarsAgoA)
            If bRC Then bRC = ValidateBarsAgo(i, eCol_BarsAgoB)
            If bRC Then bRC = Validate(i, 0, IndA)
            If bRC And Not IsNumeric(.TextMatrix(i, eCol_ExprB)) Then bRC = Validate(i, 1, IndB)
            If bRC Then
                CopyExprToRecord i, 0, IndA
                CopyExprToRecord i, 1, IndB
            Else
                Exit For
            End If
        Next
    End With
        
    SetORParentheses
    CopyToTable = bRC
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmConditionBuilder.CopyToTable"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BuildExpression
'' Description: Build the expression
'' Inputs:      None
'' Returns:     True if success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BuildExpression() As Boolean
On Error GoTo ErrSection:

    Dim i&, nEnd&
    Dim strExprA$, strExprB$
    Dim strOp$, strLogicalOp$
    Dim strOpenParen$, strCloseParen$
    Dim strBarsAgoA$, strBarsAgoB$
    Dim strCondition$
    Dim bCrAboveBelow As Boolean
    Dim strExpression As String
    Dim strMacro As String
    Dim strMacros As String
    Dim strDuplicates As String
        
    If Not CopyToTable() Then
        BuildExpression = False
        Exit Function
    End If
          
    With fgExpression
        If .MergeRow(.Rows - 1) = True Then
            nEnd = .Rows - 2
        Else
            nEnd = .Rows - 1
        End If
    End With
    
    m.strExpression = ""
    For i = fgExpression.FixedRows To nEnd
        strExprA = Trim(m.tblIndInfo(eFld_IndExprA, i))
        strExprB = Trim(m.tblIndInfo(eFld_IndExprB, i))
        strBarsAgoA = Trim(m.tblIndInfo(eFld_BarsAgoA, i))
        strBarsAgoB = Trim(m.tblIndInfo(eFld_BarsAgoB, i))
        strLogicalOp = Trim(m.tblIndInfo(eFld_LogicalOp, i))
        strOpenParen = Trim(m.tblIndInfo(eFld_OpenParen, i))
        strCloseParen = Trim(m.tblIndInfo(eFld_CloseParen, i))
        strOp = Trim(m.tblIndInfo(eFld_Operator, i))
        
        If InStr(strOp, "Crosses Above") Or InStr(strOp, "Crosses Below") Then
            bCrAboveBelow = True
        Else
            bCrAboveBelow = False
        End If
        
        If strBarsAgoA <> "0" Then AppendBarsAgo strExprA, strBarsAgoA
        
        If strBarsAgoB <> "0" Then
            If strExprB <> "True" And strExprB <> "False" And Not IsNumeric(strExprB) Then
                AppendBarsAgo strExprB, strBarsAgoB
            End If
        End If
        
        If m.bPartial Or (Len(strExprA) > 0 And Len(strExprB) > 0 And Len(strOp) > 0) Then
            If m.eExpressionType = eType_Scoring Then
                BuildScoring strExprA, strExprB, strBarsAgoA, strBarsAgoB, strLogicalOp, strOp, _
                             strOpenParen, strCloseParen, bCrAboveBelow, i, nEnd
            Else
                BuildNonScoring strExprA, strExprB, strBarsAgoA, strBarsAgoB, strLogicalOp, strOp, _
                                strOpenParen, strCloseParen, bCrAboveBelow
            End If
        End If
    Next
    
    'add extra parentheses when first logical op is OR
    With fgExpression
        If .Rows > .FixedRows + 1 Then
            strLogicalOp = .TextMatrix(.FixedRows + 1, eCol_LogicalOp)
            If strLogicalOp = "Or" Then
                i = InStr(m.strExpression, " And ")
                If i > 0 Then
                    m.strExpression = "(" & Left(m.strExpression, i - 1) & ") " & Right(m.strExpression, Len(m.strExpression) - i)
                End If
            End If
        End If
    End With
    
    ' Add macros back into the expression...
    strExpression = ""
    For i = 0 To m.astrMacros.Size - 1
        strExpression = strExpression & m.astrMacros(i) & vbCrLf
        
        strMacro = Parse(m.astrMacros(i), ":=", 1)
        If InStr(strMacros, "," & strMacro & ",") = 0 Then
            strMacros = strMacros & "," & strMacro & ","
        ElseIf InStr(strDuplicates, "," & strMacro & ",") = 0 Then
            strDuplicates = strDuplicates & "," & strMacro & ","
        End If
    Next i
    If m.astrMacros.Size > 0 Then
        m.strExpression = strExpression & vbCrLf & m.strExpression
    End If
    
    If Len(strDuplicates) > 0 Then
        strDuplicates = Mid(strDuplicates, 2, Len(strDuplicates) - 2)
        InfBox "The following variables have multiple definitions which may cause unexpected results:||" & strDuplicates, "!", , "Duplicate Variables"
    End If
        
    BuildExpression = True

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmConditionBuilder.BuildExpression"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitSystemsList
'' Description: Initialize the systems list
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitSystemsList()
On Error GoTo ErrSection:

    Dim rsSystems As Recordset
    
    Set rsSystems = mSysNav.LoadStrategiesRecordset

    m.tblSystemInfo.Clear
    m.tblSystemInfo.AddRecord "Favorite,0,0", -1, ","
    
    If Not (rsSystems.BOF And rsSystems.EOF) Then
        rsSystems.MoveFirst
        
        Do While Not rsSystems.EOF
            If mSysNav.IncludeStrategiesFromRecordset(rsSystems) Then
                m.tblSystemInfo.AddRecord rsSystems!SystemName & "," & Str(rsSystems!SystemNumber) & "," & Str(rsSystems![tblSystems.LibraryID]), -1, ","
            End If
            
            rsSystems.MoveNext
        Loop
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConditionBuilder.InitSystemsList"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CreateNew
'' Description: Create a new object from the expression
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CreateNew()
On Error GoTo ErrSection:

    Dim nSysId&, nLibId&, i&
    Dim strExprA$, strExprB$, strOp$
    Dim strLastExprMsg$
    Dim Expr As cExpression
    Dim oFunction As cFunction
    Dim idx&
    Dim Ind As cIndicator

    m.bPartial = False
    If Not BuildExpression Then
        MsgBox "One or more expressions may be incomplete. Cannot copy"
        Exit Sub
    Else
        With fgExpression
            strExprA = .TextMatrix(.Rows - 1, eCol_ExprA)
            strExprB = .TextMatrix(.Rows - 1, eCol_ExprB)
            strOp = .TextMatrix(.Rows - 1, eCol_Operator)
            If Len(strExprA) > 0 And Len(strExprB) > 0 And Len(strOp) > 0 Then
            ElseIf Len(strExprA) > 0 Or Len(strExprB) > 0 Or Len(strOp) > 0 Then
                strLastExprMsg = "The last expression is incomplete and will not be copied."
                If MsgBox(strLastExprMsg, vbOKCancel) = vbCancel Then
                    Exit Sub
                End If
            End If
        End With
    End If
    
    If m.eExpressionType = eType_Scoring Then
        txtExpression.Text = ""
    Else
        txtExpression.Text = m.strExpression
    End If
    
    Select Case m.eExpressionType
        Case eType_Undefined
            If Not m.ChartEditorForm Is Nothing Then
                frmFunctionMgrCT.ShowMe 0, m.strExpression, , , , m.ChartEditorForm
            End If
        Case eType_Function
            frmFunctionMgrCT.ShowMe 0, m.strExpression
        Case eType_Scoring
            frmFunctionMgrCT.ShowMe 0, m.strExpression, , , True
            'reset so tradesense will show if user makes changes to expressions
            m.eExpressionType = eType_Function
            BuildExpression
            txtExpression.Text = m.strExpression
        Case eType_Criteria
            frmCriteria.ShowMe "", "", , , m.strExpression
        Case eType_Clipboard
            Clipboard.Clear
            Clipboard.SetText m.strExpression
        Case eType_Rule
            If m.nSelSystemIdx >= 0 Then
                nSysId = Val(m.tblSystemInfo(1, m.nSelSystemIdx))
                nLibId = Val(m.tblSystemInfo(2, m.nSelSystemIdx))
                If nSysId = 0 Then
                    frmRule.ShowMe "", 0, 0, 0, , , , , , m.strExpression
                Else
                    frmSystemManager.ShowMe nSysId, , False, m.strExpression
                    Me.ZOrder 1
                End If
            End If
        Case eType_HighlightBars
            If m.ChartEditorForm Is Nothing Then
                frmChartCfg.ShowMe m.Chart, 0, True
                frmChartCfg.AddToChart eAdd_HighlightBars, 3, True, m.strExpression
            Else
                frmChartCfg.NewFunctionAdded m.strExpression, "", 3
            End If
        Case eType_Alert
            Set Expr = New cExpression
            Set oFunction = New cFunction
            With Expr
                .PortfolioNavigator = False
                .Functions = g.Functions
                .ValidateFunctionRule m.strExpression
                oFunction.CodedText = .CodedText
                oFunction.DataTypeID = .FunctionReturnType
                oFunction.ReturnTypeID = .FunctionReturnType
                oFunction.LateCalculating = .LateCondition
            End With
                        
            idx = m.Chart.Tree.Index("PRICE")
            Set Ind = New cIndicator
            With Ind
                .DataType = eINDIC_BooleanArray
                .Display = False
                .DisplayType = eINDIC_HighlightBars
                .Color = vbBlue
                .Name = m.strAlertName
                .CodedText = oFunction.CodedText
                .Expression = m.strExpression
                .InitAlert m.Chart
            End With
            idx = m.Chart.Tree.Add(Ind, "", idx, eTREE_LastChild)
            strOp = m.Chart.Tree.Key(idx)
            Ind.geIndId = Parse(strOp, " ", 2)
            If frmAlerts.ShowMe(Ind.AlertObject, eGDAlertType_Chart) Then
                SaveCondBuilderString Ind
                Ind.UpdateAlert 1
            Else
                Ind.UpdateAlert 0
            End If
    End Select
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConditionBuilder.CreateNew"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitOperatorsArray
'' Description: Initialize the operators array
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitOperatorsArray()
On Error GoTo ErrSection:

    m.aOperators.Clear
    m.aOperators.Add "="
    m.aOperators.Add "<>"
    m.aOperators.Add ">"
    m.aOperators.Add ">="
    m.aOperators.Add "<"
    m.aOperators.Add "<="
    m.aOperators.Add "Is Up"
    m.aOperators.Add "Is Down"
    m.aOperators.Add "Crosses Above"
    m.aOperators.Add "Crosses Below"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConditionBuilder.InitOperatorsArray"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetLogicalOp
'' Description: Set the logical operator
'' Inputs:      Operator
'' Returns:     New List
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SetLogicalOp(ByVal strOp$) As String
On Error GoTo ErrSection:

    Dim i&, strList$
    
    For i = 0 To m.tblAllItems.NumRecords - 1
        m.tblAllItems(1, i) = strOp
        strList = strList & m.tblAllItems(0, i) & strOp & vbTab & m.tblAllItems(2, i)
    Next
    
    SetLogicalOp = strList

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmConditionBuilder.SetLogicalOp"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AppendBarsAgo
'' Description: Append the number of bars ago to the expression
'' Inputs:      Expression, Bars Ago
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AppendBarsAgo(strExpr$, strBarsAgo$)
On Error GoTo ErrSection:

    strExpr = Trim(strExpr)
    strBarsAgo = Trim(strBarsAgo)
    If InStr(strExpr, " ") Then
        strExpr = "(" & strExpr & ")." & strBarsAgo
    Else
        strExpr = strExpr & "." & strBarsAgo
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConditionBuilder.AppendBarsAgo"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetORParentheses
'' Description: Set the parentheses around the OR
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetORParentheses()
On Error GoTo ErrSection:

    Dim i&, k&, nLeft&
    Dim strLogicalOp$, strLogicalOpPrev$
    
    
    nLeft = 0
    k = fgExpression.FixedRows
    'clear out any previous parentheses
    For i = 0 To m.tblIndInfo.NumRecords - 1
        m.tblIndInfo(eFld_OpenParen, i) = ""
        m.tblIndInfo(eFld_CloseParen, i) = ""
    Next
    
    For i = k + 1 To m.tblIndInfo.NumRecords - 1
        strLogicalOpPrev = m.tblIndInfo(eFld_LogicalOp, i - 1)
        strLogicalOp = m.tblIndInfo(eFld_LogicalOp, i)
        If strLogicalOp = "Or" Then
            If strLogicalOpPrev = "And" Then
                m.tblIndInfo(eFld_OpenParen, i - 1) = "("
                nLeft = nLeft + 1
            End If
        ElseIf strLogicalOp = "And" Then
            If strLogicalOpPrev = "Or" And nLeft > 0 Then
                m.tblIndInfo(eFld_CloseParen, i) = ")"
                nLeft = nLeft - 1
            End If
        End If
    Next
    
    If nLeft > 1 Then
        'something very wrong ... just remove all parenthesis
        For i = 0 To m.tblIndInfo.NumRecords - 1
            m.tblIndInfo(eFld_OpenParen, i) = ""
            m.tblIndInfo(eFld_CloseParen, i) = ""
        Next
    ElseIf nLeft = 1 Then
        i = m.tblIndInfo.NumRecords - 1
        If m.tblIndInfo(eFld_IndIdxA, i) > 0 And m.tblIndInfo(eFld_IndIdxB, i) > 0 Then
            m.tblIndInfo(eFld_CloseParen, i) = ")"
        ElseIf m.tblIndInfo(eFld_IndIdxA, i - 1) > 0 And m.tblIndInfo(eFld_IndIdxB, i - 1) > 0 Then
            m.tblIndInfo(eFld_CloseParen, i - 1) = ")"
        Else
            'something very wrong ... just remove all parenthesis
            For i = 0 To m.tblIndInfo.NumRecords - 1
                m.tblIndInfo(eFld_OpenParen, i) = ""
                m.tblIndInfo(eFld_CloseParen, i) = ""
            Next
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConditionBuilder.SetORParentheses"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BuildNonScoring
'' Description: Build a non-scoring expression
'' Inputs:      Expression A, Expression B, Bars Ago A, Bars Ago B, Logical
''              Operator, Operator, Parenthesis O, Parenthesis C, Above/Below
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub BuildNonScoring(strA$, strB$, strBarsA$, strBarsB$, strLogical$, strOp$, _
    strParenO$, strParenC$, ByVal bAboveBelow As Boolean)
On Error GoTo ErrSection:

    Dim strCondition As String          ' Condition to add to the expression
    Dim astrExpression As cGdArray      ' Expression split out into lines
    Dim lPos As Long                    ' Position of something in an array
    Dim lIndex As Long                  ' Index into a for loop
    
    Set astrExpression = New cGdArray
    astrExpression.Create eGDARRAY_Strings
    
    astrExpression.SplitFields strA, vbLf
    For lIndex = 0 To astrExpression.Size - 1
        If InStr(astrExpression(lIndex), ":=") <> 0 Then
            If m.astrMacros.BinarySearch(astrExpression(lIndex), lPos, eGdSort_IgnoreCase) = False Then
                m.astrMacros.Add astrExpression(lIndex), lPos
            End If
        Else
            strA = astrExpression(lIndex)
        End If
    Next lIndex
    
    astrExpression.SplitFields strB, vbLf
    For lIndex = 0 To astrExpression.Size - 1
        If InStr(astrExpression(lIndex), ":=") <> 0 Then
            If m.astrMacros.BinarySearch(astrExpression(lIndex), lPos, eGdSort_IgnoreCase) = False Then
                m.astrMacros.Add astrExpression(lIndex), lPos
            End If
        Else
            strB = astrExpression(lIndex)
        End If
    Next lIndex
    
    If bAboveBelow Then
        strCondition = strOp & "(" & strA & "," & strB & ")"
    Else
        strCondition = strA & " " & strOp & " " & strB
    End If
    
    If Len(m.strExpression) < 1 Then
        m.strExpression = strCondition
    Else
        If strLogical = "And" Then
            m.strExpression = m.strExpression & strParenC & " " & strLogical & " " & strParenO & strCondition
        ElseIf Len(strParenO) = 0 And Len(strParenC) = 0 Then
            m.strExpression = m.strExpression & " " & strLogical & " " & strCondition
        ElseIf Len(strParenO) = 0 And strParenC = ")" Then
            m.strExpression = m.strExpression & " " & strLogical & " " & strCondition & strParenC
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConditionBuilder.BuildNonScoring"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BuildScoring
'' Description: Build a scoring function from the expressions
'' Inputs:      Expression A, Expression B, Bars Ago A, Bars Ago B, Logical
''              Operator, Operator, Parenthesis O, Parenthesis C, Above/Below,
''              Record, End
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub BuildScoring(strA$, strB$, strBarsA$, strBarsB$, strLogical$, strOp$, _
    strParenO$, strParenC$, ByVal bAboveBelow As Boolean, ByVal nRecord, ByVal nEnd)
On Error GoTo ErrSection:
    
    'scoring format: IFF(1 and (condittion string), 1.5, 0)
    'sample condition string for crosses above: (Crosses Above (Volume , Volume.3) )
    
    Dim strCondition$, i&
    Dim astrExpression As cGdArray      ' Expression split out into lines
    Dim lPos As Long                    ' Position of something in an array
    Dim lIndex As Long                  ' Index into a for loop
    
    Set astrExpression = New cGdArray
    astrExpression.Create eGDARRAY_Strings
    
    astrExpression.SplitFields strA, vbLf
    For lIndex = 0 To astrExpression.Size - 1
        If InStr(astrExpression(lIndex), ":=") <> 0 Then
            If m.astrMacros.BinarySearch(astrExpression(lIndex), lPos, eGdSort_IgnoreCase) = False Then
                m.astrMacros.Add astrExpression(lIndex), lPos
            End If
        Else
            strA = astrExpression(lIndex)
        End If
    Next lIndex
    
    astrExpression.SplitFields strB, vbLf
    For lIndex = 0 To astrExpression.Size - 1
        If InStr(astrExpression(lIndex), ":=") <> 0 Then
            If m.astrMacros.BinarySearch(astrExpression(lIndex), lPos, eGdSort_IgnoreCase) = False Then
                m.astrMacros.Add astrExpression(lIndex), lPos
            End If
        Else
            strB = astrExpression(lIndex)
        End If
    Next lIndex
        
    If strLogical = "Or" Then
        If bAboveBelow Then
            strCondition = " " & strOp & "(" & strA & "," & strB & ")"
        Else
            strCondition = " " & strA & " " & strOp & " " & strB
        End If
    Else
        If bAboveBelow Then
            strCondition = "(" & strOp & "(" & strA & "," & strB & ")"
        Else
            strCondition = "(" & strA & " " & strOp & " " & strB
        End If
    End If
    
    If Len(m.strExpression) < 1 Then
        'm.strExpression = "IFF(1 and " & strCondition & ", 1.0, 0)"    -save awhile then remove 11-30-2004
        m.strExpression = "IFF(1 and " & strCondition
    ElseIf strLogical = "And" Then
        m.strExpression = m.strExpression & "), 1.0, 0) + IFF(1 and " & strCondition
    ElseIf strLogical = "Or" Then
        m.strExpression = m.strExpression & " Or " & strCondition
    End If
    If nRecord = nEnd Then
        m.strExpression = m.strExpression & "), 1.0, 0)"
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConditionBuilder.BuildScoring"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsOpString
'' Description: Is the given string an operator?
'' Inputs:      Potential Operator
'' Returns:     Operator or Blank
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IsOpString(ByVal strOp$) As String
On Error GoTo ErrSection:

    Dim strOperator$
    
    If strOp = "=" Or _
       strOp = "<>" Or _
       strOp = ">" Or _
       strOp = ">=" Or _
       strOp = "<" Or _
       strOp = "<=" Or _
       strOp = "Is Up" Or _
       strOp = "Is Down" Or _
       strOp = "Crosses Above" Or _
       strOp = "Crosses Below" Then
       
       strOperator = strOp
    End If
    
    IsOpString = strOperator

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmConditionBuilder.IsOpString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FakeComboCloseUp
'' Description: Fake a combo close up in the grid
'' Inputs:      Items, Expression A, Expression B, Record
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FakeComboCloseUp(aItems As cGdArray, ByVal strA$, ByVal strB$, ByVal nRecord&)
On Error GoTo ErrSection:

    Dim strTemp$, i&, j&
    Dim bFoundA As Boolean, bFoundB As Boolean
    
    'aardvark 3501 fix
    If Len(strB) = 0 Then
        bFoundB = True
    ElseIf UCase(strB) = "TRUE" Then
        m.tblIndInfo(eFld_IndIdxB, nRecord) = 9994
        bFoundB = True
    ElseIf UCase(strB) = "FALSE" Then
        m.tblIndInfo(eFld_IndIdxB, nRecord) = 9993
        bFoundB = True
    End If
    
    For i = 0 To aItems.Size - 1
        If Not bFoundA Then
            j = InStr(aItems(i), strA)
            If j > 0 Then
                strTemp = Mid(aItems(i), 2, j - 3)
                m.tblIndInfo(eFld_IndIdxA, nRecord) = Val(strTemp)
                bFoundA = True
            End If
        End If
        
        If Not bFoundB Then
            j = InStr(aItems(i), strB)
            If j > 0 Then
                strTemp = Mid(aItems(i), 2, j - 3)
                m.tblIndInfo(eFld_IndIdxB, nRecord) = Val(strTemp)
                bFoundB = True
            End If
        End If
        
        If bFoundA And bFoundB Then Exit For
        
    Next
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConditionBuilder.FakeComboCloseUp"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadIndExpr
'' Description: Load indicator expressions
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadIndExpr()
On Error GoTo ErrSection:

    Dim i&, j&, nOpIdx&
    Dim Ind As cIndicator
    Dim aExpr As New cGdArray
    Dim aItems As New cGdArray
    Dim strLogicalOp$, strOp$, strTemp$
    Dim strA$, strBarsA$, strB$, strBarsB$
    
    If m.oAlert Is Nothing Then Exit Sub
    Set Ind = m.oAlert.Indicator
    If Ind Is Nothing Then Exit Sub
    
    If m.Chart Is Nothing Then
        Set m.Chart = Ind.IndChart
    End If
    
    If Not m.Chart Is Nothing Then
        InitChartItemsList
        m.strExpression = Ind.Expression
        txtExpression.Text = m.strExpression
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConditionBuilder.LoadIndExpr"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveCondBuilderString
'' Description: Save the condition builder string
'' Inputs:      Indicator
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveCondBuilderString(Ind As cIndicator)
On Error GoTo ErrSection:

    Dim i&, strSave$
    Dim strLogical$, strA$, strBarsA$, strOp$, strB$, strBarsB$
    
    If Ind Is Nothing Then Exit Sub
    
    With fgExpression
        For i = .FixedRows To .Rows - 1
            If .MergeRow(i) = False Then
                strLogical = .TextMatrix(i, eCol_LogicalOp)
                strA = .TextMatrix(i, eCol_ExprA)
                strBarsA = .TextMatrix(i, eCol_BarsAgoA)
                strOp = .TextMatrix(i, eCol_Operator)
                strB = .TextMatrix(i, eCol_ExprB)
                strBarsB = .TextMatrix(i, eCol_BarsAgoB)
                strSave = strSave & strLogical & "~" & strA & "~" & strBarsA & "~" & strOp & "~" & strB & "~" & strBarsB & "|"
            End If
        Next
    End With
    
    Ind.CondBuilderStr = strSave

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConditionBuilder.SaveCondBuilderString"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadCondBuilderString
'' Description: Load a condition builder string
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadCondBuilderString()
On Error GoTo ErrSection:

    Dim Ind As cIndicator
    Dim i&, j&, k&, strSave$
    Dim aSave As New cGdArray
    Dim aItems As New cGdArray
    Dim aRow As New cGdArray
    
    If m.oAlert Is Nothing Then Exit Sub
    Set Ind = m.oAlert.Indicator
    If Ind Is Nothing Then Exit Sub
    strSave = Ind.CondBuilderStr
    If Len(strSave) = 0 Then Exit Sub
    
    aSave.SplitFields strSave, "|"
    aItems.SplitFields m.strAllItems, "|"
    m.bPartial = True
    For i = 0 To aSave.Size - 1
        aRow.SplitFields aSave(i), "~"
        With fgExpression
            k = .Rows - 1
            .TextMatrix(k, eCol_LogicalOp) = aRow(eCol_LogicalOp)
            .TextMatrix(k, eCol_ExprA) = aRow(eCol_ExprA)
            .TextMatrix(k, eCol_BarsAgoA) = aRow(eCol_BarsAgoA)
            .TextMatrix(k, eCol_Operator) = aRow(eCol_Operator)
            .TextMatrix(k, eCol_ExprB) = aRow(eCol_ExprB)
            .TextMatrix(k, eCol_BarsAgoB) = aRow(eCol_BarsAgoB)
            FakeComboCloseUp aItems, aRow(eCol_ExprA), aRow(eCol_ExprB), k
            .Rows = .Rows + 1
        End With
    Next
    
    m.bInitialShow = False
    fgExpression.Rows = m.tblIndInfo.NumRecords
    BuildExpression
    cmdAdd_Click
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmConditionBuilder.LoadCondBuilderString"
    
End Sub

