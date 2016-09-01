VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmPatProfitSettings 
   Caption         =   "Settings"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   7275
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniButtonImageXP cmdOK 
      Default         =   -1  'True
      Height          =   390
      Left            =   2085
      TabIndex        =   7
      Top             =   6840
      Width           =   795
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
      Caption         =   "frmPatProfitSettings.frx":0000
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmPatProfitSettings.frx":0026
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmPatProfitSettings.frx":008A
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
      Height          =   390
      Left            =   3030
      TabIndex        =   8
      Top             =   6840
      Width           =   795
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
      Caption         =   "frmPatProfitSettings.frx":00A6
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      ShowFocus       =   -1  'True
      Tristate        =   0   'False
      Pressed         =   0   'False
      Tip             =   "frmPatProfitSettings.frx":00D4
      Style           =   -1
      RoundedBorders  =   -1  'True
      xTranspColor    =   0
      yTranspColor    =   0
      MousePointer    =   0
      MouseIcon       =   "frmPatProfitSettings.frx":0138
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniFrameWL fraMethods 
      Height          =   4545
      Left            =   53
      TabIndex        =   0
      Top             =   2160
      Width           =   5805
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
      Caption         =   "frmPatProfitSettings.frx":0154
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmPatProfitSettings.frx":01A4
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmPatProfitSettings.frx":01C4
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtPercentWeight 
         Height          =   300
         Left            =   3720
         TabIndex        =   11
         Top             =   3435
         Width           =   420
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmPatProfitSettings.frx":01E0
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
         Tip             =   "frmPatProfitSettings.frx":0204
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPatProfitSettings.frx":0224
      End
      Begin HexUniControls.ctlUniCheckXP chkIndDifferences 
         Height          =   220
         Left            =   180
         TabIndex        =   10
         Top             =   3480
         Width           =   3540
         _ExtentX        =   6244
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
         Caption         =   "frmPatProfitSettings.frx":0240
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmPatProfitSettings.frx":02BC
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmPatProfitSettings.frx":02DC
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin VSFlex7LCtl.VSFlexGrid fg 
         Height          =   1485
         Left            =   90
         TabIndex        =   9
         Top             =   675
         Width           =   5580
         _cx             =   9842
         _cy             =   2619
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
      Begin HexUniControls.ctlUniLabelXP Label3 
         Height          =   285
         Left            =   90
         Top             =   405
         Width           =   5130
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
         Caption         =   "frmPatProfitSettings.frx":02F8
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmPatProfitSettings.frx":0394
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPatProfitSettings.frx":03B4
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblPercentWeight 
         Height          =   210
         Left            =   4215
         Top             =   3480
         Width           =   1185
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
         Caption         =   "frmPatProfitSettings.frx":03D0
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmPatProfitSettings.frx":0406
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPatProfitSettings.frx":0426
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP lblIndDiffDesc 
         Height          =   960
         Left            =   150
         Top             =   2340
         Width           =   5130
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
         Caption         =   "frmPatProfitSettings.frx":0442
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmPatProfitSettings.frx":04C0
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPatProfitSettings.frx":04E0
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
   Begin HexUniControls.ctlUniFrameWL fraGeneral 
      Height          =   1935
      Left            =   735
      TabIndex        =   1
      Top             =   105
      Width           =   4440
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
      Caption         =   "frmPatProfitSettings.frx":04FC
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmPatProfitSettings.frx":052A
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmPatProfitSettings.frx":054A
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtStdDevPFP 
         Height          =   315
         Left            =   2520
         TabIndex        =   4
         Top             =   270
         Width           =   525
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "frmPatProfitSettings.frx":0566
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
         Alignment       =   2
         ScrollBars      =   0
         PasswordChar    =   ""
         TrapTab         =   0   'False
         EnableContextMenu=   -1  'True
         RaiseChangeEvent=   -1  'True
         Tip             =   "frmPatProfitSettings.frx":0588
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPatProfitSettings.frx":05A8
      End
      Begin HexUniControls.ctlUniComboImageXP cboLineStyle 
         Height          =   315
         Left            =   2520
         TabIndex        =   3
         Top             =   680
         Width           =   1572
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
         Tip             =   "frmPatProfitSettings.frx":05C4
         Sorted          =   0   'False
         HScroll         =   0   'False
         RoundedBorders  =   -1  'True
         IconDim         =   16
         MousePointer    =   0
         MouseIcon       =   "frmPatProfitSettings.frx":05E4
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniCheckXP chkPfpFill 
         Height          =   220
         Left            =   210
         TabIndex        =   2
         Top             =   1560
         Width           =   2235
         _ExtentX        =   3942
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
         Caption         =   "frmPatProfitSettings.frx":0600
         Enabled         =   -1  'True
         Align           =   0
         CheckBackColor  =   -2147483643
         CheckForeColor  =   -2147483640
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Pressed         =   0   'False
         Tip             =   "frmPatProfitSettings.frx":065C
         Style           =   -1
         MousePointer    =   0
         MouseIcon       =   "frmPatProfitSettings.frx":067C
         ShowFocus       =   -1  'True
         RightToLeft     =   0   'False
      End
      Begin gdOCX.gdSelectColor gdPfpForecastColor 
         Height          =   315
         Left            =   2520
         TabIndex        =   5
         Top             =   1090
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   556
         CustomColor     =   255
      End
      Begin gdOCX.gdSelectColor gdPfpFillColor 
         Height          =   315
         Left            =   2520
         TabIndex        =   6
         Top             =   1500
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   556
         CustomColor     =   255
      End
      Begin HexUniControls.ctlUniLabelXP Label2 
         Height          =   195
         Left            =   210
         Top             =   740
         Width           =   1650
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
         Caption         =   "frmPatProfitSettings.frx":0698
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmPatProfitSettings.frx":06D8
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPatProfitSettings.frx":06F8
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label11 
         Height          =   195
         Left            =   210
         Top             =   330
         Width           =   1650
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
         Caption         =   "frmPatProfitSettings.frx":0714
         BackColor       =   12632256
         ForeColor       =   -2147483640
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   0
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmPatProfitSettings.frx":075A
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPatProfitSettings.frx":077A
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
      Begin HexUniControls.ctlUniLabelXP Label1 
         Height          =   195
         Left            =   210
         Top             =   1150
         Width           =   1650
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
         Caption         =   "frmPatProfitSettings.frx":0796
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         Alignment       =   0
         VAlignment      =   0
         BackStyle       =   1
         BorderStyle     =   0
         AutoSize        =   0   'False
         Tip             =   "frmPatProfitSettings.frx":07D6
         Style           =   0
         Enabled         =   -1  'True
         Margin          =   0
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "frmPatProfitSettings.frx":07F6
         RightToLeft     =   0   'False
         WordWrap        =   0   'False
      End
   End
End
Attribute VB_Name = "frmPatProfitSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const kName = "Patterns for Profit Settings"

Private Const kCorrDesc = "the standard correlation formula measures the strength of the linear relationship"
Private Const kShapeDesc = "measures the fit of the overall shape (based on a 'Percent R' type of comparison)"
Private Const kDirDesc = "compares how similar the direction of each value is (above/below its previous value)"
Private Const kSignsDesc = "compares how similiar the sign of each value is (i.e. above/below zero)"
Private Const kPeaksDesc = "compares the bar# of the highest and lowest peaks (e.g. converging/diverging peaks)"
Private Const kDiffDesc = "With multiple indicators in the same pane, also comparing the difference between indicators will typically provide better matches (e.g. crossovers and divergence):"

Private Enum eGridCols
    eGridCols_Use = 0
    eGridCols_Weight
    eGridCols_Type
    eGridCols_Desc
End Enum

Private Type mPrivate
    frm As Form
    oPFP As cPatternProfit
    bMethodChanged As Boolean
End Type

Private m As mPrivate

Private Sub chkIndDifferences_Click()

    Dim d#

    If Me.Visible Then
        If chkIndDifferences.Value = vbChecked Then
            d = ValOfText(txtPercentWeight.Text)
            If d > 0 Then
                d = d / 100
            Else
                d = Abs(m.oPFP.MethodWeightGet(eMethod_IndDiff, True))
            End If
            
            UpdateIndDiffOption d, True
        Else
            UpdateIndDiffOption 0, True
        End If
    End If

End Sub

Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    Unload Me

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPatProfitSettings.cmdCancel_Click"

End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    Dim i&, d#
    Dim nColor&, nFillColor&, nFillStyle&, nLineStyle&, dStandardDev#
    
    nColor = -1
    nFillColor = -1
    nFillStyle = -1
    dStandardDev = -1#
    nLineStyle = -1
    
        
    If ValOfText(txtStdDevPFP.Text) <> m.oPFP.StandardDev Then dStandardDev = ValOfText(txtStdDevPFP.Text)
    If chkPfpFill.Value <> m.oPFP.FillPattern Then nFillStyle = chkPfpFill.Value
    If gdPfpFillColor.Color <> m.oPFP.FillColor Then nFillColor = gdPfpFillColor.Color
    If gdPfpForecastColor.Color <> m.oPFP.ForecastColor Then nColor = gdPfpForecastColor.Color
    If i = m.oPFP.ForecastLineStyle Then i = -1
    
    If m.bMethodChanged Then
        If chkIndDifferences.Value = vbChecked Then
            If IndDiffWeightOk() Then
                d = ValOfText(txtPercentWeight.Text)
                m.oPFP.MethodWeightSet eMethod_IndDiff, d / 100
            Else
                Exit Sub
            End If
        Else
            m.oPFP.MethodWeightSet eMethod_IndDiff, 0
        End If
        If Not ValidateMethodWeight Then Exit Sub
    End If
    
    i = cboLineStyle.ListIndex + 1
    
    If nColor > 0 Or nFillColor > 0 Or nFillStyle <> -1 Or dStandardDev > 0 Or i <> -1 Or m.bMethodChanged Then
        m.oPFP.UpdateProperty m.frm.Chart, True, nColor, nFillColor, nFillStyle, dStandardDev, i, m.bMethodChanged
    End If
    
    Unload Me

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPatProfitSettings.cmdOk_Click"

End Sub

Private Sub fg_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrExit:

    With fg
        If Col = eGridCols_Use Then
            If .Cell(flexcpChecked, Row, Col) = flexChecked Then
                .Cell(flexcpFontBold, Row, eGridCols_Type) = True
                .Cell(flexcpFontBold, Row, eGridCols_Weight) = True
                .TextMatrix(Row, eGridCols_Weight) = Abs(m.oPFP.MethodWeightGet(Row, True))
            Else
                .Cell(flexcpFontBold, Row, eGridCols_Type) = False
                .Cell(flexcpFontBold, Row, eGridCols_Weight) = False
                .TextMatrix(Row, eGridCols_Weight) = ""
            End If
            m.bMethodChanged = True
        ElseIf eGridCols_Weight Then
            If ValOfText(fg.TextMatrix(Row, Col)) = 0 Then
                .Cell(flexcpChecked, Row, eGridCols_Use) = flexUnchecked
                .Cell(flexcpFontBold, Row, eGridCols_Type) = False
                .Cell(flexcpFontBold, Row, eGridCols_Weight) = False
                .TextMatrix(Row, eGridCols_Weight) = ""
            End If
            m.bMethodChanged = True
        End If
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPatProfitSettings.fg_AfterEdit"

End Sub

Private Sub fg_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    If Col = eGridCols_Weight Then
        If Len(fg.TextMatrix(Row, Col)) = 0 Then Cancel = True
    ElseIf Col <> eGridCols_Use Then
        Cancel = True
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPatProfitSettings.fg_BeforeEdit"

End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    Me.Icon = Picture16(ToolbarIcon("ID_IndAnalyst"), , True)
    
    g.Styler.StyleForm Me
    
    With cboLineStyle
        .AddItem "Thin"
        .AddItem "Medium Thin"
        .AddItem "Medium"
        .AddItem "Medium Thick"
        .AddItem "Thick"
        .AddItem "Extra Thick"
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPatProfitSettings.Form_Load"

End Sub

Public Sub ShowMe(frm As Form)
On Error GoTo ErrSection:

    Dim bLocked As Boolean

    If frm Is Nothing Then GoTo ErrExit
    If Not IsFrmChart(frm) Then GoTo ErrExit
    
    Set m.frm = frm
    
    Set m.oPFP = m.frm.PatternProfitObj
    If m.oPFP Is Nothing Then GoTo ErrExit
    
    InitControls

    CenterTheForm Me
    
    If TypeOf m.frm Is frmChart2 Then
        LockWindowUpdate 0                          '6386
        bLocked = LockWindowUpdate(m.frm.hWnd)
    End If
        
    ShowForm Me, eForm_Modal, frmMain, , ALT_GRID_ROW_COLOR
    
    If bLocked Then LockWindowUpdate 0

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPatProfitSettings.ShowMe"

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    Set m.frm = Nothing
    Set m.oPFP = Nothing

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPatProfitSettings.Form_Unload"

End Sub

Private Sub InitControls()
On Error GoTo ErrExit:

    Dim i&, d#

    'General settings
    txtStdDevPFP.Text = Str(m.oPFP.StandardDev)
    gdPfpForecastColor.Color = m.oPFP.ForecastColor
    gdPfpFillColor.Color = m.oPFP.FillColor
    chkPfpFill.Value = m.oPFP.FillPattern
    
    i = m.oPFP.ForecastLineStyle - 1
    
    If i >= 0 And i < cboLineStyle.ListCount Then
        cboLineStyle.ListIndex = m.oPFP.ForecastLineStyle - 1
    Else
        cboLineStyle.ListIndex = 0
    End If
    
    'Methods grid
    With fg
        .Redraw = flexRDNone
        
        SetupGrid fg, eGridMode_Grid
        .AllowUserResizing = flexResizeNone
        .ExplorerBar = flexExNone
        .ScrollBars = flexScrollBarNone
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .SelectionMode = flexSelectionFree
        .HighLight = flexHighlightNever
        .Editable = flexEDKbdMouse
        .ExtendLastCol = True
        .WordWrap = True
        .FixedCols = 0
        .FixedRows = 1
        .Rows = 6
        .Cols = 4
        'alignment
        .ColAlignment(eGridCols_Use) = flexAlignCenterCenter
        .ColAlignment(eGridCols_Weight) = flexAlignCenterCenter
        .ColAlignment(eGridCols_Type) = flexAlignCenterCenter
        .ColAlignment(eGridCols_Desc) = flexAlignLeftCenter
        'column headers
        .TextMatrix(0, eColsPFP_Use) = "Use"
        .TextMatrix(0, eGridCols_Weight) = "Weight"
        .TextMatrix(0, eGridCols_Type) = "Type"
        .TextMatrix(0, eGridCols_Desc) = "Description"
        'data type
        .ColDataType(eGridCols_Use) = flexDTBoolean
        .ColDataType(eGridCols_Weight) = flexDTDouble
        
        'text for type
        .TextMatrix(1, eGridCols_Type) = "Correlation"
        .TextMatrix(2, eGridCols_Type) = "Shape"
        .TextMatrix(3, eGridCols_Type) = "Directional"
        .TextMatrix(4, eGridCols_Type) = "Signs"
        .TextMatrix(5, eGridCols_Type) = "Peaks"
        
        'text for description
        .TextMatrix(1, eGridCols_Desc) = kCorrDesc
        .TextMatrix(2, eGridCols_Desc) = kShapeDesc
        .TextMatrix(3, eGridCols_Desc) = kDirDesc
        .TextMatrix(4, eGridCols_Desc) = kSignsDesc
        .TextMatrix(5, eGridCols_Desc) = kPeaksDesc
        
        'weight
        For i = 1 To 5
            d = m.oPFP.MethodWeightGet(i)
            If d > 0 Then
                .TextMatrix(i, eGridCols_Weight) = d
                .Cell(flexcpChecked, i, eGridCols_Use) = flexChecked
                .Cell(flexcpFontBold, i, eGridCols_Weight) = True
                .Cell(flexcpFontBold, i, eGridCols_Type) = True
            Else
                .TextMatrix(i, eGridCols_Weight) = ""
                .Cell(flexcpChecked, i, eGridCols_Use) = flexUnchecked
                .Cell(flexcpFontBold, i, eGridCols_Weight) = False
                .Cell(flexcpFontBold, i, eGridCols_Type) = False
            End If
        Next
        
        'columns width
        .ColWidth(eGridCols_Use) = 375
        .ColWidth(eGridCols_Weight) = 615
        .ColWidth(eGridCols_Type) = 1230
        
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize eGridCols_Desc
        
        .RowHeight(2) = .RowHeight(1)
        .RowHeight(3) = .RowHeight(1)
        .RowHeight(4) = .RowHeight(1)
        .RowHeight(5) = .RowHeight(1)
        
        .Height = .RowHeight(5) * 5 + .RowHeight(0) + 75
        
        .Redraw = flexRDBuffered
    End With
    
    'Indicators differences
    With lblIndDiffDesc
        .BorderStyle = 0
        .WordWrap = True
        .AutoSize = True
        .Top = fg.Top + fg.Height + fg.RowHeight(0)
        .Caption = kDiffDesc
        
        chkIndDifferences.Top = .Top + .Height + fg.RowHeight(0) / 2
        chkIndDifferences.Left = .Left
        
        txtPercentWeight.Top = chkIndDifferences.Top - 45
        lblPercentWeight.Top = chkIndDifferences.Top
    End With
    
    UpdateIndDiffOption
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPatProfitSettings.InitControls"

End Sub

Private Sub UpdateIndDiffOption(Optional ByVal dValue# = kNullData, Optional ByVal bSave As Boolean = False)
On Error GoTo ErrSection:

    Dim d#
    
    If dValue = kNullData Then
        d = m.oPFP.MethodWeightGet(eMethod_IndDiff)
    Else
        d = dValue
        If bSave Then
            If m.oPFP.MethodWeightGet(eMethod_IndDiff) <> d Then m.bMethodChanged = True
        End If
    End If
    
    If d > 0 Then
        If Not Me.Visible Then chkIndDifferences.Value = vbChecked
        txtPercentWeight.Enabled = True
        txtPercentWeight.Text = Str(d * 100)
    Else
        If Not Me.Visible Then chkIndDifferences.Value = vbUnchecked
        txtPercentWeight.Enabled = False
        txtPercentWeight.Text = ""
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmPatProfitSettings.UpdateIndDiffOption"

End Sub

Private Function IndDiffWeightOk() As Boolean
On Error GoTo ErrSection:

    Dim d#

    d = ValOfText(txtPercentWeight.Text)
    If d > 0 And d <= 100 Then
        IndDiffWeightOk = True
    Else
        InfBox "The weighting for indicator differences needs to be a percentage value between 1 and 100.", "I", "Ok", kName
    End If

ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmPatProfitSettings.IndDiffWeightOk"

End Function

Private Function ValidateMethodWeight() As Boolean
On Error GoTo ErrExit

    Dim i&, d#
    Dim bOK As Boolean
    
    With fg
        If .Cell(flexcpChecked, eMethod_Corr, eGridCols_Use) = flexChecked Or _
           .Cell(flexcpChecked, eMethod_Shape, eGridCols_Use) = flexChecked Or _
           .Cell(flexcpChecked, eMehod_Directional, eGridCols_Use) = flexChecked Then
           bOK = True
        End If
        
        If Not bOK Then
            InfBox "One of these methods:" & vbCrLf & "Correlation, Shape or Directional must be selected.", "I", "Okay"
            Exit Function
        End If
        
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, eGridCols_Use) = flexChecked Then
                d = ValOfText(.TextMatrix(i, eGridCols_Weight))
                m.oPFP.MethodWeightSet i, d
            Else
                m.oPFP.MethodWeightSet i, 0
            End If
        Next
    End With

    ValidateMethodWeight = bOK

ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmPatProfitSettings.ValidateMethodWeight"

End Function
























