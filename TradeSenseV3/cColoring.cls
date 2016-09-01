VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cColoring"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text

Private Type mPrivate
    EditorOptions As cEditorOptions
End Type
Private m As mPrivate


Public Function Color(ByVal strText As String) As String
On Error GoTo ErrSection:

    Dim lBegOfStr As Long
    Dim lPos As Long
    Dim lStrLen As Long
    Dim alColor As cGdArray
    Dim alStart As cGdArray
    Dim alLength As cGdArray
    Dim lTokens As Long
    Dim strRichText As String
    Dim lindex As Long
    
    'Used to store string information based on lTokens...
    Set alColor = New cGdArray
    alColor.Create eGDARRAY_Longs
    Set alStart = New cGdArray
    alStart.Create eGDARRAY_Longs
    Set alLength = New cGdArray
    alLength.Create eGDARRAY_Longs
    
    'Build color information arrays...
    lPos = 1: lTokens = 0
    Do Until lPos = 0
        lPos = InStr(lPos, strText, "~")
        If lPos > 0 Then
            lTokens = lTokens + 1
            alColor.Add Mid(strText, lPos + 1, 2)
            lStrLen = Val(Mid(strText, lPos + 3, 3))
            alLength.Add lStrLen
            
            'Calculate the beginning of the tag...
            'Calculate the beginning of the string (offset from the tag begin)
            lBegOfStr = (lPos + gTokenLen) - (lTokens * gTokenLen) - 1 'minus1 to adj for zero based
            'lBegOfStr = lBegOfStr + Val(Mid(.Text, lPos + 3, 3)) - 1
            alStart.Add lBegOfStr
            
            lPos = lPos + gTokenLen
        End If
    Loop
    
    'Strip out lTokens
    lPos = 1
    Do Until lPos = 0
        lPos = InStr(1, strText, "~")
        If lPos > 0 Then
            If lPos > 1 Then
                strText = Left(strText, lPos - 1) & _
                        Right(strText, Len(strText) - _
                        (lPos + gTokenLen - 1))
            Else
                strText = Right(strText, Len(strText) - gTokenLen)
            End If
        End If
    Loop
    
    FormatRTFHeader strRichText
    
    'Color strings based on array values
    For lindex = 0 To alColor.Size - 1
        With m.EditorOptions
            Select Case alColor.Item(lindex)
                
                Case gPT_Enter, gPT_EnterFormatting
                    strRichText = strRichText & "\par "
                
                Case gPT_Tab
                    strRichText = strRichText & "\tab "
                
                Case gPT_Comment
                    HighlightSymbols strText, strRichText, lindex, alStart, _
                        alLength, vbGreen, False, False
                        
                'Show first error only...
                Case gPT_Error
                    HighlightSymbols strText, strRichText, lindex, alStart, _
                        alLength, .ErrorColor, .ErrorBoldFace, _
                        .ErrorItalics
                
                Case gPT_FCompiled, gPT_FCompiledAction, _
                     gPT_FTradeSense, gPT_FInternal
                    HighlightSymbols strText, strRichText, lindex, alStart, alLength, _
                        .FunctionsColor, .FunctionsBoldFace, _
                        .FunctionsItalics
                
                Case gPT_Parm
                    HighlightSymbols strText, strRichText, lindex, alStart, alLength, _
                        .ParmColor, .ParmBoldFace, .ParmItalics
                
                Case gPT_Text
                    HighlightSymbols strText, strRichText, lindex, alStart, alLength, _
                        .ParmColor, False, False
                
                Case gPT_Add, gPT_Sub, gPT_Div, gPT_Mult, gPT_GT, gPT_GE, gPT_LT, _
                     gPT_LE, gPT_NE, gPT_EQ, gPT_And, gPT_Or, gPT_Of, gPT_Not, _
                     gPT_LeftPar, gPT_RightPar, gPT_FLParen, gPT_FRParen, _
                     gPT_Offset
                    HighlightSymbols strText, strRichText, lindex, alStart, alLength, _
                        .OperatorsColor, .OperatorsBoldFace, _
                        .OperatorsItalics
                
                Case gPT_If, gPT_Then, gPT_Else, gPT_ElseIf, _
                     gPT_EndIf, gPT_DoUntil, gPT_EndDo
                    HighlightSymbols strText, strRichText, lindex, alStart, alLength, _
                        .OperatorsColor, .OperatorsBoldFace, _
                        .OperatorsItalics
                    
                Case Else
                    HighlightSymbols strText, strRichText, lindex, alStart, _
                        alLength, 0, False, False
            End Select
        End With
    Next lindex
    
    'Enclose with trailing brace and stip out last blank format
    strRichText = strRichText & "}"
    Color = strRichText

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "TSOCX.cColoring.Color", eGDRaiseError_Raise, g.strAppPath

End Function

Private Sub FormatRTFHeader(strRichText As String)

    strRichText = "{\rtf1\ansi\ansicpg1252\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fswiss Arial;}{\f3\fswiss Arial;}}" & _
            "{\colortbl\red0\green0\blue0;\red0\green0\blue255;\red255\green0\blue0;\red0\green150\blue0;}" & _
            "\deflang1033\pard"

End Sub

Private Sub HighlightSymbols(strText As String, strRTF As String, lCurPos As Long, _
        alStart As cGdArray, alLength As cGdArray, lColor As Long, bBold As Boolean, _
        bItalic As Boolean)
On Error GoTo ErrSection:

    Dim strWorkRTF As String
    Dim strWorkPlainText As String
    Dim lStartPos As Long
    Dim lLength As Long
    Dim lSlashPos As Long
    
    'Use for getting the next phrase.  (Searching for right paren).
    Dim lNextStartPos As Long
    Dim lNextLength As Long
    Dim strNextPlainText As String
    
    'Get current text information
    strWorkPlainText = Mid(strText, alStart.Item(lCurPos) + 1, alLength.Item(lCurPos))
    lStartPos = alStart.Item(lCurPos)
    lLength = alLength.Item(lCurPos)
    If lCurPos + 1 <= alStart.Size Then
        lNextStartPos = alStart.Item(lCurPos + 1)
        lNextLength = alLength.Item(lCurPos + 1)
        strNextPlainText = Mid(strText, lNextStartPos + 1, lNextLength)
    Else
        lNextStartPos = 0
    End If
    
    'font size 9> fs18  10>fs20)
    'font name f2> points to font number 2 in header
    strWorkRTF = "\plain\f2\fs20"
    
    'Insert color code...
    Select Case lColor
        Case vbBlue: strWorkRTF = strWorkRTF & "\cf1"
        Case vbRed: strWorkRTF = strWorkRTF & "\cf2"
        Case vbGreen: strWorkRTF = strWorkRTF & "\cf3"
    End Select
    
    'Insert bold
    If bBold Then
        strWorkRTF = strWorkRTF & "\b"
    End If

    'Insert Italic
    If bItalic Then
        strWorkRTF = strWorkRTF & "\i"
    End If

    'if text has a "\" then make it "\\"...
    lSlashPos = InStr(1, strWorkPlainText, "\")
    Do Until lSlashPos = 0
        strWorkPlainText = Left(strWorkPlainText, lSlashPos) & "\" & _
            Right(strWorkPlainText, Len(strWorkPlainText) - lSlashPos)
        lSlashPos = InStr(lSlashPos + 2, strWorkPlainText, "\")
    Loop
    
    
    'Insert text after RTF codes
    strRTF = strRTF & strWorkRTF & " " & strWorkPlainText
    
    'Insert blank after text (unless text is a parenthesis, decimal pt
    If strWorkPlainText <> "(" And strNextPlainText <> ")" And _
        strWorkPlainText <> "." And strNextPlainText <> "." Then
        strRTF = strRTF & "\plain\f2\fs20" & "  "
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cColoring.HighlightSymbols", eGDRaiseError_Raise, g.strAppPath
    
End Sub

Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.EditorOptions = New cEditorOptions
    m.EditorOptions.Load
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cColoring.Initialize", eGDRaiseError_Raise, g.strAppPath

End Sub

Private Sub Class_Terminate()
    Set m.EditorOptions = Nothing
End Sub
