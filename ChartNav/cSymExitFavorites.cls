VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSymExitFavorites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cSymExitFavorites.cls
'' Description: Class for exit favorites that are symbol specific
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 01/25/2012   MJM         Create
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    strBaseSym As String
    aFavoriteLetter As cGdArray         'contains 4 elements with letters A,B,C,D
    aFavoriteExits As cGdArray          'contains 4 auto exit objects corresponding to A,B,C,D faovite letters
End Type
Private m As mPrivate

Private Sub Class_Initialize()
    Set m.aFavoriteLetter = New cGdArray
    Set m.aFavoriteExits = New cGdArray
End Sub

Private Sub Class_Terminate()
    Set m.aFavoriteLetter = Nothing
    Set m.aFavoriteExits = Nothing
End Sub

Public Sub ExitFavoritesSet(ByVal strBaseSym$, ByRef aExitsInfo As cGdArray)
On Error GoTo ErrExit:

    Dim i&, j&
    Dim strLetter$, strName$
    
    Dim oExit As cExitStrategy

    If Len(strBaseSym) = 0 Then Exit Sub
    If aExitsInfo Is Nothing Then Exit Sub
    If aExitsInfo.Size < 1 Then Exit Sub
    
    If aExitsInfo.Size > 4 Then
        j = 3
    Else
        j = aExitsInfo.Size - 1
    End If
    
    aExitsInfo.Sort
    
    m.strBaseSym = strBaseSym
    m.aFavoriteLetter.Clear
    m.aFavoriteExits.Clear
    
    For i = 0 To j
        strLetter = Parse(aExitsInfo(i), "=", 1)
        strName = Parse(aExitsInfo(i), "=", 2)
        
        Set oExit = New cExitStrategy
        If Len(strName) > 0 Then
            'if load fails then revert to new object with default info
            If Not oExit.Load(strName) Then Set oExit = New cExitStrategy
        End If
        
        Select Case strLetter
            Case "A", "B", "C", "D"
                m.aFavoriteLetter.Add strLetter
                m.aFavoriteExits.Add oExit
        End Select
    Next

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSymExitFavorites.ExitFavoritesSet"
    
End Sub

Public Sub ExitFavoritesGet(ByVal strBaseSym$, ByRef aExitsInfo As cGdArray)
On Error GoTo ErrExit:

    Dim i&
    
    If aExitsInfo Is Nothing Then Exit Sub
    If strBaseSym <> m.strBaseSym Then Exit Sub
    
    aExitsInfo.Clear
    
    'precautionary checks, theoretically should never be true
    If m.aFavoriteLetter.Size <= 0 Or m.aFavoriteExits.Size <= 0 Then Exit Sub
    If m.aFavoriteLetter.Size <> m.aFavoriteExits.Size Then Exit Sub
    
    For i = 0 To m.aFavoriteLetter.Size
        If Not m.aFavoriteExits(i) Is Nothing Then
            If Len(m.aFavoriteExits(i).FileName) > 0 Then
                aExitsInfo.Add m.aFavoriteLetter(i) & "=" & m.aFavoriteExits(i).FileName
            End If
        End If
    Next
    

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSymExitFavorites.ExitFavoritesGet"
    
End Sub

Public Function ExitFavoritesAssigned(ByVal strBaseSym$) As String
On Error GoTo ErrExit

    Dim i&, strFavorites$
    
    If strBaseSym = m.strBaseSym Then
        For i = 0 To m.aFavoriteLetter.Size - 1
            If Not m.aFavoriteExits(i) Is Nothing Then
                If Len(m.aFavoriteExits(i).FileName) > 0 Then
                    If i = 0 Then
                        strFavorites = m.aFavoriteLetter(i)
                    Else
                        strFavorites = strFavorites & "|" & m.aFavoriteLetter(i)
                    End If
                End If
            End If
        Next
    End If
    
    ExitFavoritesAssigned = strFavorites

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSymExitFavorites.ExitFavoritesAssigned"
    
End Function

Public Property Get BaseSym() As String
    BaseSym = m.strBaseSym
End Property

Public Function ExitObjectGet(ByVal strBaseSym$, ByVal strLetter$) As cExitStrategy
On Error GoTo ErrExit

    Dim i&
    
    Set ExitObjectGet = Nothing

    If strBaseSym <> m.strBaseSym Then Exit Function    'precautionary
    
    For i = 0 To m.aFavoriteLetter.Size - 1
        If strLetter = m.aFavoriteLetter(i) Then
            Set ExitObjectGet = m.aFavoriteExits(i)
            Exit For
        End If
    Next

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSymExitFavorites.ExitObjectGet"
    
End Function

Public Sub ExitObjectClear(ByVal strBaseSym$, ByVal strLetter$)
On Error GoTo ErrExit
    
    Dim i&
    
    If strBaseSym <> m.strBaseSym Then Exit Sub    'precautionary
    
    For i = 0 To m.aFavoriteLetter.Size - 1
        If strLetter = m.aFavoriteLetter(i) Then
            m.aFavoriteExits.Remove i
            m.aFavoriteLetter.Remove i
            Exit For
        End If
    Next

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSymExitFavorites.ExitObjectClear"
    
End Sub

Public Sub ExitObjectSet(ByVal strBaseSym$, ByVal strLetter$, ByRef AutoExit As cExitStrategy)
On Error GoTo ErrExit

    If strBaseSym <> m.strBaseSym Then Exit Sub    'precautionary
    If AutoExit Is Nothing Then Exit Sub
    
    ExitObjectClear strBaseSym, strLetter           'remove so don't have duplicate
    
    m.aFavoriteLetter.Add strLetter
    m.aFavoriteExits.Add AutoExit

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSymExitFavorites.ExitObjectSet"
    
End Sub
