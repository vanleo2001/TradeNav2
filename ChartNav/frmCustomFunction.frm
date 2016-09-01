VERSION 5.00
Object = "{C0F09D2D-1125-11D7-8DA9-0004757A4B66}#1.9#0"; "NavTradeSenseOCXV3.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmCustomFunction 
   Caption         =   "Custom Indicator"
   ClientHeight    =   1650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7095
   Icon            =   "frmCustomFunction.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin NavTradeSenseV3.Editor Editor1 
      Height          =   1095
      Left            =   60
      TabIndex        =   0
      Top             =   420
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   1931
   End
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   390
      Left            =   4800
      TabIndex        =   1
      Top             =   0
      Width           =   2115
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
      Caption         =   "frmCustomFunction.frx":000C
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmCustomFunction.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmCustomFunction.frx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Height          =   330
         Left            =   1140
         TabIndex        =   3
         Top             =   60
         Width           =   900
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
         Caption         =   "frmCustomFunction.frx":0068
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmCustomFunction.frx":0096
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmCustomFunction.frx":00B6
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Default         =   -1  'True
         Height          =   330
         Left            =   120
         TabIndex        =   2
         Top             =   60
         Width           =   900
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
         Caption         =   "frmCustomFunction.frx":00D2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmCustomFunction.frx":00F8
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmCustomFunction.frx":0118
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniLabelXP lblEditor 
      Height          =   225
      Left            =   120
      Top             =   180
      Width           =   4995
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
      Caption         =   "frmCustomFunction.frx":0134
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmCustomFunction.frx":01CE
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmCustomFunction.frx":01EE
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmCustomFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmCustomFunction.frm
'' Description: Allows the user to enter a custom expression for an indicator
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    strUserText As String
    strCodedText As String
    bIsBoolean As Boolean
    bSkipAutoIf As Boolean
    ListLoading As cListLoading
    Function As cFunction
    bCancelled As Boolean
    bAllowOtherMarkets As Boolean
    bAllowMacros As Boolean             ' Do we allow macros in this instance?
End Type
Private m As mPrivate

Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    'Dim frm As New frmCriteria
    'frm.ShowMe App.Path & "\Custom", "", True
    'Exit Sub

    m.bCancelled = True
    Me.Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCustomFunction.cmdCancel_Click"
    
End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    m.bCancelled = False
    If Verify Then Me.Hide
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCustomFunction.cmdOK_Click"
    
End Sub

Private Sub Editor1_Change()
On Error GoTo ErrSection:

    If (m.bAllowMacros = False) Then
        If (InStr(Editor1.Text, ":=") <> 0) Then
            InfBox "Assignment operators are not allowed in this expression", "!", , "Expression Error"
            Editor1.Text = Replace(Editor1.Text, ":=", "")
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCustomFunction.Editor1_Change"
    
End Sub

Private Sub Editor1_GotFocus()
On Error GoTo ErrSection:

    Set g.ActiveEditor = Editor1
    With Editor1
        .FunctionsRef = g.Functions
        .Lists = m.ListLoading.Lists
        .DisableEnterKey = Not m.bAllowMacros
        .ShowNewFunction = False
        .Usage = 6             'Usage Mask: 2=means bit 2 turned on
        .TurnOnEditing
        .Refresh
    End With
    
    If Len(Trim(Editor1.Text)) = 0 And Not m.bSkipAutoIf Then
        Editor1.Text = ""
        SendKeys " "
    End If
    
    m.bSkipAutoIf = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCustomFunction.Editor1_GotFocus"
    
End Sub

Private Sub Editor1_LostFocus()
On Error GoTo ErrSection:
    
    Set g.ActiveEditor = Nothing
    Editor1.RemoveTradeSense

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCustomFunction.Editor1_LostFocus"
    
End Sub

Private Sub Form_Activate()
On Error GoTo ErrSection:
    
    MoveFocus Editor1

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCustomFunction.Form_Activate"
    
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
    RaiseError "frmCustomFunction.Form_KeyDown"
    
End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    CenterTheForm Me
    
    g.Styler.StyleForm Me
    
    Me.Icon = Picture16("kBlank")

    Set m.Function = New cFunction
    With m.Function
        .FunctionID = 0
        .Load
    End With
        
    'Load internally generated TradeSense lists (Symbols, etc.)
    ' (when activate, in case list has changed)
    Set m.ListLoading = New cListLoading
    m.ListLoading.Load
    
    With Editor1
        .AppPath = App.Path
        .FunctionsRef = g.Functions
        .Lists = m.ListLoading.Lists
        .DisableEnterKey = False ' True
        .Usage = 6             'Usage Mask: 2=means bit 2 turned on
        .TurnOnEditing
        .Refresh
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCustomFunction.Form_Load"
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode = 0 Then
        Cancel = True
        Me.Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCustomFunction.Form_QueryUnload"
    
End Sub

Private Sub Form_Resize()
On Error Resume Next

    If LimitFormSize(Me, fraButtons.Width, fraButtons.Height * 3) Then Exit Sub

    With fraButtons
        '.Move (Me.ScaleWidth - .Width) \ 2, Me.ScaleHeight - .Height
        .Move (Me.ScaleWidth - .Width) ', Me.ScaleHeight - .Height
    End With
    
    With Editor1
        .Move .Left, .Top, Me.ScaleWidth - .Left * 2, _
            Me.ScaleHeight - .Top - .Left ' fraButtons.Top - .Top
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    Set m.Function = Nothing
    Set m.ListLoading = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCustomFunction.Form_Unload"
    
End Sub

Private Function Verify(Optional ByVal bInvisible As Boolean = False) As Boolean
On Error GoTo ErrSection:
   
    Dim i&, strChk$
    Dim lNumDays As Long
    Dim strNotKnown As String
    Dim bExtraInputs As Boolean
    Dim strMsg          As String
    Dim wrkText         As String
    Dim Expr            As cExpression
    Dim Inputs          As cInputs
 
    If Len(Trim(Editor1.Text)) = 0 Then
        m.strCodedText = ""
        Verify = True
        Exit Function
    End If
 
    'Shut things off, get ready for verifying rule
    If Not bInvisible Then
        Screen.MousePointer = vbHourglass
        LockWindowUpdate Me.hWnd
    End If
    m.strCodedText = ""
    
    'Verify...
    Set Expr = New cExpression
    With Expr
        .PortfolioNavigator = False
        .Functions = g.Functions
        .ValidateFunctionRule Editor1.Text
        
        'Convert to rich text
        If Not bInvisible Then
            'Load internally generated TradeSense lists (Symbols, etc.)
            If m.ListLoading Is Nothing Then
                Set m.ListLoading = New cListLoading
                m.ListLoading.Load
            End If
            Editor1.TurnOffEditing
            wrkText = .EditText
            Editor1.TextRTF = m.Function.GetRTF(wrkText)
            Editor1.ExprIsFormatted = True
            Editor1.SelStart = 999999
        End If
    
        'Save verify settings
        'mFunction.FunctionIDs = .GetFIDs
        'mFunction.Formatted = .EditText
        'mFunction.FormattedWithFillWords = .Preview
        'mFunction.CodedText = .CodedText
        'mFunction.FunctionIDs = .GetFIDs
        'mFunction.DataTypeID = .FunctionReturnType
        'mFunction.ReturnTypeID = .FunctionReturnType
        
        'Save Late calculating flags (borrows "LateCondition" property)
        'If .LateCondition Then
        '    mFunction.LateCalculating = True
        'Else
        '    mFunction.LateCalculating = False
        'End If
    End With
        
    ' see if unwanted inputs exist
    bExtraInputs = False
    strNotKnown = ""
    If Not Expr.Inputs Is Nothing Then
        'ShowParmLine TradeSense
        'mFunction.TradeSenseUsage = TradeSense.Tag
        Set Inputs = Expr.Inputs
        For i = 1 To Expr.Inputs.Count
            strChk = UCase(Inputs.Item(i).ParmName)
            If Inputs.Item(i).ParmTypeID <> 5 Then
                strNotKnown = strNotKnown & "|" & Inputs.Item(i).ParmName
                bExtraInputs = True
            ElseIf Not m.bAllowOtherMarkets And Left(strChk, 6) <> "MARKET" Then
                If strChk <> "WEEKLY" And strChk <> "DAILY" And strChk <> "MONTHLY" And _
                        strChk <> "GC" And strChk <> "TQ" And Left(strChk, 1) <> Chr(34) Then
                    strNotKnown = strNotKnown & "|" & Inputs.Item(i).ParmName
                    bExtraInputs = True
                End If
            End If
        Next
    End If
    If bExtraInputs Then
        If Not bInvisible Then
            InfBox "Error: Unrecognized items in expression:|" & strNotKnown & "|", _
                "!", , "Error"
            'EnableButtons False
            'cmdVerify.Enabled = True
        End If
    Else
        ' successful
        m.strCodedText = Expr.CodedText
        i = Expr.FunctionReturnType
        If i = 3 Or i = 6 Then
            m.bIsBoolean = True
        Else
            m.bIsBoolean = False
        End If
        
        If Not bInvisible Then
            LockWindowUpdate 0
            'EnableButtons True
            'cmdVerify.Enabled = False
        End If
        
        Verify = True
    End If
    
    
ErrExit:
    If Not bInvisible Then
        Screen.MousePointer = vbDefault
        LockWindowUpdate 0
    End If
    Set Expr = Nothing
    Exit Function

ErrSection:
    If bInvisible Then Resume ErrExit
    
    Screen.MousePointer = vbDefault
    LockWindowUpdate 0
    
    'TradeSense error occurred...
    If Err.Number < 0 Or Left(Err.Source, 5) = "Class" Then
        'svErr = Err.Number
        'svSource = Err.Source
        'svErrDesc = Err.Description
        
        'Highlight error in advanced editor...
        If Expr.EditText <> "" Then
            With Editor1
                .TurnOffEditing
                wrkText = Expr.EditText
                .ExprIsFormatted = False
                .TextRTF = m.Function.GetRTF(wrkText)
                .ExprIsFormatted = True
            End With
            Editor1.TurnOnEditing
        End If
        
        Set Expr = Nothing
        'Err.Raise svErr, svSource, svErrDesc
        InfBox Err.Description, "e", , "Invalid Expression"
    Else
        Set Expr = Nothing
        RaiseError "frmCustomFunction.Verify"
    End If

End Function

' return: -1=Cancelled, 0=Invalid, 1=Numeric, 2=Boolean
Public Function ShowMe(strUserText$, strCodedText$, _
    Optional ByVal bInvisible As Boolean = False, _
    Optional ByVal bAllowOtherMarkets = False, Optional ByVal bAllowMacros As Boolean = True, _
    Optional ByRef Chart As cChart) As Long
On Error GoTo ErrSection:

    m.bCancelled = False
    m.bAllowOtherMarkets = bAllowOtherMarkets
    m.bAllowMacros = bAllowMacros
    Editor1.Text = strUserText
    
    If Not Chart Is Nothing Then CenterFormOnChart Me, Chart        '6499
    
    Verify bInvisible
    If Not bInvisible Then
        If Me.WindowState <> 0 Then Me.WindowState = 0
        'ShowForm Me, True
        ShowForm Me, eForm_ActModal
    End If
    
    If m.bCancelled Then
        ShowMe = -1
    ElseIf Len(m.strCodedText) > 0 Then
        strUserText = Editor1.Text
        strCodedText = m.strCodedText
        If m.bIsBoolean Then
            ShowMe = 2
        Else
            ShowMe = 1
        End If
    Else
        ShowMe = 0
    End If
    
ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmCustomFunction.ShowMe"
    
End Function

