VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEnterKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text
'Ensures the format of enter key characters used in expression paragraphs

Private Type mPrivate
    iBlankPos As Integer
    strCurChar As String
    iCurPos As Integer
    iEnterKeyPos As Integer
    iTextLen As Integer
End Type
Private m As mPrivate

Public Sub ReFormat(pStr As String)
On Error GoTo ErrSection:

    RemoveLeadingEnterKeys pStr
    RemoveTrailingEnterKeys pStr
    RemoveBlanksAroundEnterKeys pStr
    RemoveDuplicateEnterKeys pStr
    AppendToEnd pStr
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cEnterKey.ReFormat", eGDRaiseError_Raise, g.strAppPath

End Sub

'This was create for ValidateRuleAdv.  It was create to maintain binary
'comptibility for now.  Eventually, open up all routines in this class and
'make public.  Let the caller control the enterkey parsing.
Public Sub ReFormat2(pStr As String)
On Error GoTo ErrSection:

    RemoveLeadingEnterKeys pStr
    RemoveTrailingEnterKeys pStr
    RemoveBlanksAroundEnterKeys pStr
    RemoveDoubleSpacing pStr
    AppendToEnd pStr
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cEnterKey.ReFormat2", eGDRaiseError_Raise, g.strAppPath

End Sub

'Removes double spacing
Private Sub RemoveDoubleSpacing(pStr As String)
On Error GoTo ErrSection:
    
    Dim ConsecEK          As String
    Dim ConsecEKPos       As Integer
    Dim tmpstr            As String
    Dim TextLen           As Long
    
    tmpstr = ""
    TextLen = Len(pStr)
    
    ConsecEK = Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(13) & Chr(10)
    ConsecEKPos = InStr(1, pStr, ConsecEK)
    Do Until ConsecEKPos = 0
        tmpstr = Left(pStr, ConsecEKPos + 1)
        tmpstr = tmpstr & Right(pStr, m.iTextLen - (ConsecEKPos + 3))
        pStr = tmpstr
        TextLen = Len(pStr)
        ConsecEKPos = InStr(1, pStr, ConsecEK)
    Loop
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cEnterKey.RemoveDoubleSpacing", eGDRaiseError_Raise, g.strAppPath

End Sub

'Replace all enter keys with a blank character
Private Sub RemoveAllEnterKeys(pStr As String)
On Error GoTo ErrSection:
    
    Dim FirstHalf       As String
    Dim SecondHalf      As String
    
    If Len(pStr) > 0 Then
        m.iCurPos = 1
        Do Until m.iCurPos = 0
            m.iCurPos = InStr(1, pStr, Chr(13))
            If m.iCurPos > 0 Then
                FirstHalf = Left(pStr, m.iCurPos - 1)
                SecondHalf = Right(pStr, Len(pStr) - (m.iCurPos + 1))
                pStr = FirstHalf & " " & SecondHalf
            End If
        Loop
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "TSOCX.cEnterKey.RemoveAllEnterKeys", eGDRaiseError_Raise, g.strAppPath

End Sub

'Remove any extra line feeds(13), carriage returns(10)
Private Sub RemoveLeadingEnterKeys(pStr As String)
On Error GoTo ErrSection:
    
    Dim CurChar     As String
    Dim EnterPos    As Long
    Dim X           As Long
    
    m.iTextLen = Len(pStr)
    m.iCurPos = 1
    
    If m.iTextLen > 0 Then
    
        'Remove leading spaces and tabs preceeding <Enter> key.  Only
        'if they occur before the first valid tradesense statement.
        For X = 1 To m.iTextLen
            CurChar = Mid(pStr, X, 1)
            If CurChar = Chr(32) Or CurChar = Chr(9) Then
                m.iCurPos = m.iCurPos + 1
            Else
                If CurChar = Chr(13) Then
                    m.iCurPos = X
                Else
                    If CurChar <> Chr(10) Then
                        Exit For
                    End If
                End If
            End If
        Next X
        
        Do Until Mid(pStr, m.iCurPos, 1) <> Chr(13) And _
                 Mid(pStr, m.iCurPos, 1) <> Chr(10) And _
                 m.iCurPos <= m.iTextLen
            m.iCurPos = m.iCurPos + 1
        Loop
        
        If m.iCurPos > 1 Then
            pStr = Right(pStr, m.iTextLen - m.iCurPos + 1)
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cEnterKey.RemoveLeadingEnterKeys", eGDRaiseError_Raise, g.strAppPath

End Sub

'Remove any extra enter keys pressed after the last enter key in the clause
'Find last enter key and then go backwards until a non enter key is pressed.
Private Sub RemoveTrailingEnterKeys(pStr As String)
On Error GoTo ErrSection:

    Dim X       As Integer
    
    pStr = RTrim(pStr)
    m.iTextLen = Len(pStr)
    
    For X = m.iTextLen To 1 Step -1
        m.strCurChar = Mid(pStr, X, 1)
        If m.strCurChar <> Chr(13) And m.strCurChar <> Chr(10) Then Exit For
    Next X
    pStr = Left(pStr, X)
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cEnterKey.RemoveTrailingEnterKeys", eGDRaiseError_Raise, g.strAppPath

End Sub

'Remove blanks in front of Enter Keys...
Private Sub RemoveBlanksAroundEnterKeys(pStr As String)
On Error GoTo ErrSection:

    ' Check 1st character before enter keys.  remove blanks if found
    m.iBlankPos = InStr(1, pStr, Chr(32) & Chr(13))
    Do Until m.iBlankPos = 0
        pStr = Left(pStr, m.iBlankPos - 1) & Right(pStr, Len(pStr) - m.iBlankPos)
        m.iBlankPos = InStr(1, pStr, Chr(32) & Chr(13))
    Loop
    
    ' Remove blank characters after the Carriage control character
    m.iBlankPos = InStr(1, pStr, Chr(10) & Chr(32))
    Do Until m.iBlankPos = 0
        pStr = Left(pStr, m.iBlankPos) & Right(pStr, Len(pStr) - (m.iBlankPos + 1))
        m.iBlankPos = InStr(1, pStr, Chr(10) & Chr(32))
    Loop
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cEnterKey.RemoveBlanksAroundEnterKeys", eGDRaiseError_Raise, g.strAppPath

End Sub

'Removes duplicate enter key combinations
Private Sub RemoveDuplicateEnterKeys(pStr As String)
On Error GoTo ErrSection:
    
    Dim TwoConsecEnterKeys          As String
    Dim TwoConsecEnterKeysPos       As Integer
    Dim tmpstr                      As String

    tmpstr = ""
    m.iTextLen = Len(pStr)
    
    TwoConsecEnterKeys = Chr(13) & Chr(10) & Chr(13) & Chr(10)
    TwoConsecEnterKeysPos = InStr(1, pStr, TwoConsecEnterKeys)
    Do Until TwoConsecEnterKeysPos = 0
        tmpstr = Left(pStr, TwoConsecEnterKeysPos + 1)
        tmpstr = tmpstr & Right(pStr, m.iTextLen - (TwoConsecEnterKeysPos + 3))
        pStr = tmpstr
        m.iTextLen = Len(pStr)
        TwoConsecEnterKeysPos = InStr(1, pStr, TwoConsecEnterKeys)
    Loop
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cEnterKey.RemoveDuplicateEnterKeys", eGDRaiseError_Raise, g.strAppPath

End Sub

' Inserts one blank line between If/Then clauses.  Phrases take the format
' If/Action, blankline, If/Action, blankline, etc.
Private Sub InsertBlankLines(pStr As String)
On Error GoTo ErrSection:
    
    Dim First           As Boolean
    Dim tmpstr          As String
    
    tmpstr = ""
    m.iTextLen = Len(pStr)
    First = True
    m.iEnterKeyPos = InStr(1, pStr, Chr(13))
    
    Do Until m.iEnterKeyPos = 0
        If First Then
            First = False
            m.iEnterKeyPos = m.iEnterKeyPos + 2
        Else
            tmpstr = Left(pStr, m.iEnterKeyPos + 1)
            tmpstr = tmpstr & Chr(13) & Chr(10) & Right(pStr, m.iTextLen - (m.iEnterKeyPos + 1))
            pStr = tmpstr
            m.iTextLen = Len(pStr)
            First = True
            m.iEnterKeyPos = m.iEnterKeyPos + 4
        End If
        m.iEnterKeyPos = InStr(m.iEnterKeyPos, pStr, Chr(13))
    Loop
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "TSOCX.cEnterKey.InsertBlankLines", eGDRaiseError_Raise, g.strAppPath

End Sub

' Insert an enter key sequence if one is not found at the end of paragraph.
Private Sub AppendToEnd(pStr As String)
On Error GoTo ErrSection:
    
    If Len(pStr) > 0 Then
        If Right(pStr, 2) <> Chr(13) & Chr(10) Then
            pStr = pStr & Chr(13) & Chr(10)
        End If
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "TSOCX.cEnterKey.AppendToEnd", eGDRaiseError_Raise, g.strAppPath

End Sub
