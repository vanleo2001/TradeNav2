VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cInputsLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text

Private Type mPrivate
    'Properties referencing cInputsLine control
    ParmText As RichTextBox
    strEditText As String
    
    'Properties used to control position/highlighting of cInputsLine
    iCaretPos As Integer
    Func As cFunction
    iFunctionBegin As Integer
    iFunctionLeftParen As Integer
    iCurParmPos As Integer
    vRetVal As Variant
End Type
Private m As mPrivate

'Current parameter position relative to 0 in RTF box...
Property Get CurParmPos() As Integer
    CurParmPos = m.iCurParmPos
End Property

'These properties are set just before the show method is executed
Property Let FunctionRef(pData As cFunction)
    Set m.Func = pData
End Property
Property Let FunctionBegin(pData As Integer)
    m.iFunctionBegin = pData
End Property
Property Let FunctionLeftParen(pData As Integer)
    m.iFunctionLeftParen = pData
End Property

'Format cInputsLine
Property Get EditText()
On Error GoTo ErrSection:
    
    Dim Cnt     As Integer
    Dim X       As Integer
    Dim s       As String
    Dim iNonHiddenParms As Integer
    
    'All parameters start with a Left paren...
    s = m.Func.FunctionName & " ("
    
    With m.Func.Inputs
    
        'If no parms exist for function strip out left paren
        Cnt = .Count
        If Cnt = 0 Then
            s = Left(s, Len(s) - 2)
            EditText = s
            m.strEditText = s
            Exit Property
        Else
            'otherwise string together parms separated by commas.  Don't
            'show Hidden inputs.
            For X = 1 To Cnt
                If .Item(X).ParmTypeUsage <> gInternalData Then
                    s = s & .Item(X).ParmName & ", "
                    iNonHiddenParms = iNonHiddenParms + 1
                End If
            Next X
        End If
    End With
    
    If iNonHiddenParms > 0 Then
        'Remove extra comma and append Right paren to complete cInputsLine...
        s = Left(s, Len(s) - 2) & ")"
    Else
        'Remove extra Left paren to complete cInputsLine...
        s = Left(s, Len(s) - 2)
    End If
    EditText = s
    m.strEditText = s
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "TSOCX.cInputsLine.EditText", eGDRaiseError_Raise, g.strAppPath

End Property

'Show the floating cInputsLine at position X,Y...
Public Sub Show(pX As Long, pY As Long)
On Error GoTo ErrSection:
    
    Dim TopPos          As Long
    Dim W               As Long
    
    'Suppress window display...
    m.vRetVal = LockWindowUpdate(frmInputsLine.hWnd)

    'Show cInputsLine form X,Y coordinates
    With frmInputsLine
        .Move pX, pY
        FormOnTop .hWnd, True, 0, 0
        .Visible = True
    End With
    
    'Move caret to parm line and retrieve X position for the beginning and
    'ending of FunctionName.  This determines the width of mFunctionName.
    m.ParmText.Text = " " & m.strEditText & " "
    frmInputsLine.SizeParmLine m.ParmText.Text, 25
    
    'Show first parm (if it exists) in bold
    If m.Func.Inputs.Count > 0 Then
        HighlightParm 1 + m.Func.HiddenInputs, True
    End If
    
    'Refresh window updates...
    m.vRetVal = LockWindowUpdate(0)
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cInputsLine.Show", eGDRaiseError_Raise, g.strAppPath

End Sub

'Highlights the selected parm in the cInputsLine text in bold
Public Sub HighlightParm(pParmNbr As Byte, pHighlight As Boolean)
On Error GoTo ErrSection:
    
    Dim ParmStr As String
    
    If pParmNbr = 0 Then Exit Sub
    
    'Suppress window display...
    m.vRetVal = LockWindowUpdate(frmInputsLine.hWnd)
    
    If m.Func.Inputs.Count < pParmNbr Then GoTo ErrExit
    
    ParmStr = m.Func.Inputs.Item(pParmNbr).ParmName
    
    With m.ParmText
        'Unhighlight entire line
        .SelStart = 0
        .SelLength = Len(.Text)
        .SelBold = False
        .SelStart = 0
        
        'Highlight input
        If .Find(ParmStr, 1) > -1 Then
            .SelStart = .Find(ParmStr, 1)
            m.iCurParmPos = .SelStart
            .SelLength = Len(ParmStr)
            .SelBold = pHighlight
            .SelItalic = pHighlight
            .SelLength = 0
        End If
    End With
    
ErrExit:
    'Refresh window updates...
    m.vRetVal = LockWindowUpdate(0)
    
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cInputsLine.HighlightParm", eGDRaiseError_Raise, g.strAppPath

End Sub

'Highlights the current function
Public Sub HighlightFunction(pBold As Boolean)
On Error GoTo ErrSection:

    'Suppress window display...
    m.vRetVal = LockWindowUpdate(frmInputsLine.hWnd)
    
    With m.ParmText
        'Unhighlight entire line
        .SelStart = 0
        .SelLength = Len(.Text)
        .SelBold = False
        .SelStart = 0
        
        'Highlight function
        .SelStart = 1
        .SelLength = m.iFunctionLeftParen - m.iFunctionBegin
        .SelBold = pBold
        .SelItalic = False
        .SelLength = 0
    End With
    
ErrExit:
    'Refresh window updates...
    m.vRetVal = LockWindowUpdate(0)
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cInputsLine.HighlightFunction", eGDRaiseError_Raise, g.strAppPath

End Sub

Private Sub Class_Initialize()

    With frmInputsLine
        Set m.ParmText = !ParmText
    End With

End Sub

Private Sub Class_Terminate()
    
    Set m.ParmText = Nothing

End Sub
