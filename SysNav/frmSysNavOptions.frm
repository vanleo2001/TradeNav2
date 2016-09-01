VERSION 5.00
Object = "{3B008041-905A-11D1-B4AE-444553540000}#1.0#0"; "Vsocx6.ocx"
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.13#0"; "gdOCX.OCx"
Begin VB.Form frmSysNavOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "System Navigator Options"
   ClientHeight    =   5685
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   9105
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Corner 
      Caption         =   "Corner"
      Height          =   375
      Left            =   7815
      TabIndex        =   22
      Top             =   5340
      Visible         =   0   'False
      Width           =   1335
   End
   Begin vsOcx6LibCtl.vsIndexTab vsIndexTab1 
      Height          =   5385
      Left            =   120
      TabIndex        =   2
      Top             =   135
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   9499
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
      Caption         =   "&Editor|&Strategies|&General"
      Align           =   0
      Appearance      =   1
      CurrTab         =   1
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
      Begin vsOcx6LibCtl.vsElastic vsElastic4 
         Height          =   5010
         Left            =   8115
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   330
         Width           =   7380
         _ExtentX        =   13018
         _ExtentY        =   8837
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
         Appearance      =   0
         MousePointer    =   0
         _ConvInfo       =   1
         Version         =   600
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   192
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         Appearance      =   0
         AutoSizeChildren=   0
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
         GridRows        =   0
         GridCols        =   0
         _GridInfo       =   ""
         Begin VB.TextBox txtDeveloper 
            Height          =   345
            Left            =   390
            TabIndex        =   20
            Top             =   1650
            Width           =   3990
         End
         Begin VB.CheckBox chkShowSave 
            Caption         =   "Show SAVE confirmation message box throughout application"
            Height          =   390
            Left            =   435
            TabIndex        =   19
            Top             =   615
            Width           =   4905
         End
         Begin VB.Label Label5 
            Caption         =   "Default Developer Name"
            Height          =   240
            Left            =   390
            TabIndex        =   21
            Top             =   1395
            Width           =   3735
         End
      End
      Begin vsOcx6LibCtl.vsElastic vsElastic2 
         Height          =   5010
         Left            =   45
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   330
         Width           =   7380
         _ExtentX        =   13018
         _ExtentY        =   8837
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
         Appearance      =   0
         MousePointer    =   0
         _ConvInfo       =   1
         Version         =   600
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   192
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         Appearance      =   0
         AutoSizeChildren=   0
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
         GridRows        =   0
         GridCols        =   0
         _GridInfo       =   ""
         Begin VB.Frame Frame3 
            Caption         =   "Back Testing"
            Height          =   1980
            Left            =   285
            TabIndex        =   14
            Top             =   330
            Width           =   6765
            Begin VB.CheckBox chkDirtyFuncLib 
               Caption         =   "Reload function library on every test (will run a little slower)"
               Height          =   255
               Left            =   240
               TabIndex        =   41
               Top             =   1500
               Width           =   4875
            End
            Begin VB.Frame frTestingMethod 
               Caption         =   "Assume High was hit before Low of bar"
               Height          =   855
               Left            =   240
               TabIndex        =   15
               Tag             =   "GENESIS"
               Top             =   420
               Width           =   3720
               Begin VB.OptionButton OptTestingMethod 
                  Caption         =   "if Open > Close of bar"
                  Height          =   255
                  Index           =   0
                  Left            =   360
                  TabIndex        =   17
                  Tag             =   "GENESIS"
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   2055
               End
               Begin VB.OptionButton OptTestingMethod 
                  Caption         =   "if Open > Midpoint of bar"
                  Height          =   255
                  Index           =   1
                  Left            =   360
                  TabIndex        =   16
                  Tag             =   "OMEGA"
                  Top             =   495
                  Width           =   2175
               End
            End
         End
      End
      Begin vsOcx6LibCtl.vsElastic vsElastic1 
         Height          =   5010
         Left            =   -8025
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   330
         Width           =   7380
         _ExtentX        =   13018
         _ExtentY        =   8837
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
         Appearance      =   0
         MousePointer    =   0
         _ConvInfo       =   1
         Version         =   600
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   192
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         Appearance      =   0
         AutoSizeChildren=   0
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
         GridRows        =   0
         GridCols        =   0
         _GridInfo       =   ""
         Begin VB.Frame Frame2 
            Caption         =   "TradeSense Formatting"
            Height          =   2460
            Left            =   180
            TabIndex        =   8
            Top             =   2070
            Width           =   6780
            Begin VB.TextBox txtFontName 
               Enabled         =   0   'False
               Height          =   285
               Left            =   4365
               TabIndex        =   39
               Text            =   "Arial"
               Top             =   1605
               Width           =   1905
            End
            Begin VB.TextBox txtFontSize 
               Enabled         =   0   'False
               Height          =   300
               Left            =   4380
               TabIndex        =   37
               Text            =   "9"
               Top             =   825
               Width           =   495
            End
            Begin VB.CheckBox chkbold 
               Caption         =   "Bold"
               Enabled         =   0   'False
               Height          =   210
               Index           =   3
               Left            =   3360
               TabIndex        =   36
               Top             =   1905
               Width           =   720
            End
            Begin VB.CheckBox chkbold 
               Caption         =   "Bold"
               Enabled         =   0   'False
               Height          =   210
               Index           =   2
               Left            =   3360
               TabIndex        =   35
               Top             =   1470
               Width           =   720
            End
            Begin VB.CheckBox chkbold 
               Caption         =   "Bold"
               Enabled         =   0   'False
               Height          =   210
               Index           =   1
               Left            =   3360
               TabIndex        =   34
               Top             =   1050
               Width           =   720
            End
            Begin VB.CheckBox chkItalic 
               Caption         =   "Italic"
               Enabled         =   0   'False
               Height          =   210
               Index           =   3
               Left            =   2520
               TabIndex        =   33
               Top             =   1905
               Width           =   675
            End
            Begin VB.CheckBox chkItalic 
               Caption         =   "Italic"
               Enabled         =   0   'False
               Height          =   210
               Index           =   2
               Left            =   2520
               TabIndex        =   32
               Top             =   1470
               Width           =   675
            End
            Begin VB.CheckBox chkItalic 
               Caption         =   "Italic"
               Enabled         =   0   'False
               Height          =   210
               Index           =   1
               Left            =   2520
               TabIndex        =   31
               Top             =   1050
               Width           =   675
            End
            Begin gdOCX.gdSelectColor gdColor 
               Height          =   315
               Index           =   0
               Left            =   1350
               TabIndex        =   23
               Top             =   585
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   556
               Enabled         =   0   'False
               Color           =   0
               CustomColor     =   0
            End
            Begin VB.CheckBox chkItalic 
               Caption         =   "Italic"
               Enabled         =   0   'False
               Height          =   210
               Index           =   0
               Left            =   2520
               TabIndex        =   10
               Top             =   645
               Width           =   675
            End
            Begin VB.CheckBox chkbold 
               Caption         =   "Bold"
               Enabled         =   0   'False
               Height          =   210
               Index           =   0
               Left            =   3360
               TabIndex        =   9
               Top             =   645
               Width           =   720
            End
            Begin gdOCX.gdSelectColor gdColor 
               Height          =   315
               Index           =   1
               Left            =   1350
               TabIndex        =   24
               Top             =   1005
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   556
               Enabled         =   0   'False
               Color           =   0
               CustomColor     =   0
            End
            Begin gdOCX.gdSelectColor gdColor 
               Height          =   315
               Index           =   2
               Left            =   1350
               TabIndex        =   25
               Top             =   1425
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   556
               Enabled         =   0   'False
               Color           =   16711680
               CustomColor     =   16711680
            End
            Begin gdOCX.gdSelectColor gdColor 
               Height          =   315
               Index           =   3
               Left            =   1350
               TabIndex        =   26
               Top             =   1860
               Width           =   945
               _ExtentX        =   1667
               _ExtentY        =   556
               Enabled         =   0   'False
               CustomColor     =   255
            End
            Begin VB.Label Label4 
               Caption         =   "Font &name"
               Height          =   195
               Left            =   4365
               TabIndex        =   40
               Top             =   1380
               Width           =   840
            End
            Begin VB.Label Label3 
               Caption         =   "Font &size"
               Height          =   240
               Left            =   4395
               TabIndex        =   38
               Top             =   600
               Width           =   900
            End
            Begin VB.Label Label8 
               Caption         =   "Errors"
               Height          =   285
               Index           =   3
               Left            =   405
               TabIndex        =   30
               Top             =   1920
               Width           =   765
            End
            Begin VB.Label Label8 
               Caption         =   "Operators"
               Height          =   285
               Index           =   2
               Left            =   375
               TabIndex        =   29
               Top             =   1455
               Width           =   810
            End
            Begin VB.Label Label8 
               Caption         =   "Inputs"
               Height          =   285
               Index           =   1
               Left            =   390
               TabIndex        =   28
               Top             =   1020
               Width           =   825
            End
            Begin VB.Label Label8 
               Caption         =   "Functions"
               Height          =   285
               Index           =   0
               Left            =   375
               TabIndex        =   27
               Top             =   615
               Width           =   855
            End
            Begin VB.Label Label2 
               Caption         =   "Color"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               Left            =   1530
               TabIndex        =   11
               Top             =   315
               Width           =   540
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "TradeSense Options"
            Height          =   1545
            Left            =   180
            TabIndex        =   4
            Top             =   270
            Width           =   6795
            Begin VB.CheckBox chkTradeSenseOptions 
               Caption         =   "Show Fill Words"
               Enabled         =   0   'False
               Height          =   240
               Index           =   1
               Left            =   2865
               TabIndex        =   12
               Top             =   405
               Width           =   2430
            End
            Begin VB.TextBox txtRowsToDisplay 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   300
               Left            =   240
               TabIndex        =   6
               Text            =   "7"
               Top             =   780
               Width           =   495
            End
            Begin VB.CheckBox chkTradeSenseOptions 
               Caption         =   "Show TradeSense Lists"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   5
               Top             =   360
               Width           =   2220
            End
            Begin VB.Label Label1 
               Caption         =   "Rows displayed in TradeSense Listbox"
               Height          =   285
               Left            =   855
               TabIndex        =   7
               Top             =   810
               Width           =   2790
            End
         End
      End
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7770
      TabIndex        =   1
      Top             =   930
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   7785
      TabIndex        =   0
      Top             =   450
      Width           =   1215
   End
End
Attribute VB_Name = "frmSysNavOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    EditorOptions   As cEditorOptions
End Type
Private m As mPrivate

'TradeSense Formating
Private Const C_FUNCS = 0
Private Const C_INPUTS = 1
Private Const C_OPERATORS = 2
Private Const C_ERRORS = 3

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub chkbold_Click(Index As Integer)
    OKButton.Enabled = True
End Sub

Private Sub chkGeneral_Click()
    OKButton.Enabled = True
End Sub

Private Sub chkDirtyFuncLib_Click()
    OKButton.Enabled = True
End Sub

Private Sub chkItalic_Click(Index As Integer)
    OKButton.Enabled = True
End Sub

Private Sub chkShowSave_Click()
    OKButton.Enabled = True
End Sub

Private Sub chkTradeSenseOptions_Click(Index As Integer)
    OKButton.Enabled = True
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
    RaiseError "frmSysNavOptions.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m.EditorOptions = Nothing
End Sub

Private Sub gdColor_Validate(Index As Integer, Cancel As Boolean)
    OKButton.Enabled = True
End Sub

Private Sub txtDeveloper_Change()
    OKButton.Enabled = True
End Sub

Private Sub txtFontSize_Change()
    OKButton.Enabled = True
End Sub

Private Sub txtFontSize_LostFocus()
    
    If ValOfText(txtFontSize.Text) > 14 Then
        txtFontSize.Text = 14
    End If
    
    If ValOfText(txtFontSize.Text) < 8 Then
        txtFontSize.Text = 8
    End If

End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:
    Dim TestingMethod       As String
    

    ReSizeMDIChildForm Me, Corner
    CenterTheForm Me

    With vsIndexTab1
        .TabVisible(0) = False
        .CurrTab = 1
    End With

    Set m.EditorOptions = New cEditorOptions
    m.EditorOptions.Load
    With m.EditorOptions
        chkTradeSenseOptions(0).Value = .EditorOn * -1
        chkTradeSenseOptions(1).Value = .FillWordsOn * -1
        txtRowsToDisplay.Text = .RowsToDisplay
        txtFontSize.Text = .eFontSize
        txtFontName.Text = .eFontName
        gdColor(C_FUNCS).Color = .FunctionsColor
        gdColor(C_INPUTS).Color = .ParmColor
        gdColor(C_OPERATORS).Color = .OperatorsColor
        gdColor(C_ERRORS).Color = .ErrorColor
        chkItalic(C_FUNCS) = .FunctionsItalics * -1
        chkItalic(C_INPUTS) = .ParmItalics * -1
        chkItalic(C_OPERATORS) = .OperatorsItalics * -1
        chkItalic(C_ERRORS) = .ErrorItalics * -1
        chkbold(C_FUNCS) = .FunctionsBoldFace * -1
        chkbold(C_INPUTS) = .ParmBoldFace * -1
        chkbold(C_OPERATORS) = .OperatorsBoldFace * -1
        chkbold(C_ERRORS) = .ErrorBoldFace * -1
    End With
    
    chkDirtyFuncLib = GetIniFileProperty("DirtyFuncLib", 0, "Systems", g.strIniFile)
    
    TestingMethod = GetIniFileProperty("TestingMethod", "GENESIS", "Systems", _
        g.strIniFile)
    If TestingMethod = "GENESIS" Then
        OptTestingMethod(0).Value = 1
    Else
        OptTestingMethod(1).Value = 1
    End If
    
    'Fix this later (it is stored as 1 or 0, to be consistent it should change to being
    'stored as -1 for true, 0 for false)
    chkShowSave = GetIniFileProperty("DontShowSaveConfirmation", _
        1, "Misc", g.strIniFile)
    
    txtDeveloper = GetIniFileProperty("Developer", _
        "", "Misc", g.strIniFile)
    
    OKButton.Enabled = False
    
ErrExit:
    Exit Sub
ErrSection:
    ShowMsg
    Resume ErrExit:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:
    
    Dim Response        As Variant
    If OKButton.Enabled Then
        Response = MsgBox("Do you want to save your changes?", vbQuestion + vbYesNoCancel, "Confirmation")
        If Response = vbYes Then
            Save
        Else
            If Response = vbCancel Then
                Cancel = True
            End If
        End If
    End If
    
ErrExit:
    Exit Sub
ErrSection:
    ShowMsg
    Resume ErrExit:
End Sub

Private Sub gdColor_Changed(Index As Integer)
    OKButton.Enabled = True
End Sub

Private Sub OKButton_Click()
On Error GoTo ErrSection:
    
    Save
    OKButton.Enabled = False
    
    Unload Me
    
ErrExit:
    Exit Sub
ErrSection:
    ShowMsg
    Resume ErrExit:
End Sub

Private Sub Save()
On Error GoTo ErrSection:
    Dim RetVal      As Variant
    
    With m.EditorOptions
        .EditorOn = chkTradeSenseOptions(0).Value * -1
        .FillWordsOn = chkTradeSenseOptions(1).Value * -1
        .RowsToDisplay = ValOfText(txtRowsToDisplay.Text)
        .eFontName = txtFontName.Text
        .eFontSize = txtFontSize.Text
        .FunctionsColor = gdColor(C_FUNCS).Color
        .ParmColor = gdColor(C_INPUTS).Color
        .OperatorsColor = gdColor(C_OPERATORS).Color
        .ErrorColor = gdColor(C_ERRORS).Color
        .FunctionsItalics = chkItalic(C_FUNCS) * -1
        .ParmItalics = chkItalic(C_INPUTS) * -1
        .OperatorsItalics = chkItalic(C_OPERATORS) * -1
        .ErrorItalics = chkItalic(C_ERRORS) * -1
        .FunctionsBoldFace = chkbold(C_FUNCS) * -1
        .ParmBoldFace = chkbold(C_INPUTS) * -1
        .OperatorsBoldFace = chkbold(C_OPERATORS) * -1
        .ErrorBoldFace = chkbold(C_ERRORS) * -1
        .Save
    End With
    
    SetIniFileProperty "DirtyFuncLib", chkDirtyFuncLib, "Systems", g.strIniFile
    RetVal = SetIniFileProperty("TestingMethod", frTestingMethod.Tag, _
            "Systems", g.strIniFile)
    RetVal = SetIniFileProperty("DontShowSaveConfirmation", _
        chkShowSave, "Misc", g.strIniFile)
    RetVal = SetIniFileProperty("Developer", _
        txtDeveloper.Text, "Misc", g.strIniFile)
        
    OKButton.Enabled = False
    
ErrExit:
    Exit Sub
ErrSection:
    ShowMsg
    Resume ErrExit:
End Sub

Private Sub OptTestingMethod_Click(Index As Integer)
    OKButton.Enabled = True
    frTestingMethod.Tag = OptTestingMethod(Index).Tag
End Sub

Private Sub txtExpenses_Change()
    OKButton.Enabled = True
End Sub

Private Sub txtRowsToDisplay_Change()
    OKButton.Enabled = True
End Sub

