VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private cCalledFrom         As String
Private cSourceID           As Byte
Private cEditorOCX          As Editor
Private cRule               As cRule

'Information returned to caller
Property Get FormattedFillWords() As String
    FormattedFillWords = gExpression.FormattedFillWords
End Property
Property Get Formatted() As String
    Formatted = gExpression.Formatted
End Property
Property Get CodedText() As String
    CodedText = gExpression.CodedText
End Property
Property Get Inputs() As cInputs
    Set Inputs = gExpression.Inputs
End Property

Property Let CalledFrom(pData As String)
    cCalledFrom = pData
End Property
Property Let EditorOCX(pData As Editor)
    Set cEditorOCX = pData
End Property
Property Let SourceID(pData As Byte)
    cSourceID = pData
End Property


Public Sub Verify(pMode As Byte)
On Error GoTo ErrSection:
    Dim RetVal      As Long
    Dim i&
    
    'Turn off editing
    With cEditorOCX
        .TurnOffEditing
        .RemoveTradeSense
    End With
    
    With gExpression
        .mode = pMode
If 1 Then
        .ValidateRule cEditorOCX.Text
Else
    'BenchMark ""
    For i = 1 To 10
        .ValidateRule cEditorOCX.Text
    Next
    'BenchMark "Validated"
End If
        cEditorOCX.ExprIsFormatted = True
        cEditorOCX.TextRtf = cRule.GetRTF(.Formatted)
        If Not .ProcessedOK Then
            MsgBox .ErrMessage, vbInformation, "Message"
            Screen.MousePointer = vbDefault
            GoTo ErrExit
        End If
    End With
    
    'Turn on editing
    cEditorOCX.TurnOnEditing
    
ErrExit:
    Exit Sub
ErrSection:
    ShowMsg
    Resume ErrExit:
End Sub

Public Sub InitEditor(pActionCodedName As String)
On Error GoTo ErrSection
    
    'If the functions collection is not loaded then do it now
    ''If gFunctions Is Nothing Then
    ''    InitFunctions
    ''End If
    
    If gExpression Is Nothing Then
        InitRuleEditor
    End If
    
    gExpression.ActionCodedName = pActionCodedName
    gExpression.SourceID = cSourceID
    With cEditorOCX
        .FunctionsRef = gFunctions
        .ActionCodedName = pActionCodedName
        .SourceID = cSourceID
        .TurnOnEditing
        .Refresh
    End With
    
ErrExit:
    Exit Sub
ErrSection:
    ShowMsg
    Resume ErrExit:
End Sub

Private Sub InitRuleEditor()
    
    Set gExpression = New cExpression
    With gExpression
        .SourceID = C_Both
        .cFunctionsRef = gFunctions
    End With

End Sub

Private Sub Class_Initialize()
    Set cRule = New cRule
End Sub

Private Sub Class_Terminate()
    Set cRule = Nothing
End Sub
