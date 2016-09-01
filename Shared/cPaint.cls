VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    hWnd As Long
    hDC As Long
    bNeedToReleaseDC As Boolean
    
    hOrigPen As Long
    hOrigBrush As Long
    hNewPen As Long
    hNewBrush As Long
    
    nCurX As Long
    nCurY As Long
End Type
Private m As mPrivate

Private Sub Class_Terminate()
    
    Release

End Sub

Public Sub Init(ByVal hWnd As Long, Optional ByVal hDC As Long = 0)

    Release
    
    m.hWnd = hWnd
    If hDC = 0 Then
        m.hDC = GetWindowDC(hWnd)
        m.bNeedToReleaseDC = True
    Else
        m.hDC = hDC
        m.bNeedToReleaseDC = False
    End If

End Sub

Public Sub Release()
    
    If m.hDC <> 0 And m.hWnd <> 0 Then
        If m.hNewPen <> 0 Then
            SelectObject m.hDC, m.hOrigPen
            DeleteObject m.hNewPen
            m.hNewPen = 0
        End If
        If m.hNewBrush <> 0 Then
            SelectObject m.hDC, m.hOrigBrush
            DeleteObject m.hNewBrush
            m.hNewBrush = 0
        End If
        
        If m.bNeedToReleaseDC Then
            ReleaseDC m.hWnd, m.hDC
        End If
        m.hDC = 0
        m.hWnd = 0
    End If

End Sub

Public Property Get hDC() As Long
    hDC = m.hDC
End Property

Public Property Get hWnd() As Long
    hWnd = m.hWnd
End Property

Public Sub SetPen(ByVal nColor&, Optional ByVal nPixelWidth& = 1)

    If m.hDC = 0 Then Exit Sub
    
    If m.hNewPen <> 0 Then
        SelectObject m.hDC, m.hOrigPen
        DeleteObject m.hNewPen
    End If
    m.hNewPen = CreatePen(0, nPixelWidth, nColor)
    m.hOrigPen = SelectObject(m.hDC, m.hNewPen)

End Sub

Public Sub SetBrush(ByVal nColor&, Optional ByVal bSetPenAlso As Boolean = False)

    If m.hDC = 0 Then Exit Sub
    
    If m.hNewBrush <> 0 Then
        SelectObject m.hDC, m.hOrigBrush
        DeleteObject m.hNewBrush
    End If
    m.hNewBrush = CreateSolidBrush(nColor)
    m.hOrigBrush = SelectObject(m.hDC, m.hNewBrush)
    
    If bSetPenAlso Then
        SetPen nColor
    End If

End Sub

Public Sub MoveTo(ByVal nX&, ByVal nY&, ByVal bRelative As Boolean)

    Dim pt As POINTAPI

    If m.hDC = 0 Then Exit Sub
    
    If bRelative Then
        m.nCurX = m.nCurX + nX
        m.nCurY = m.nCurY + nY
    Else
        m.nCurX = nX
        m.nCurY = nY
    End If
    MoveToEx m.hDC, m.nCurX, m.nCurY, pt

End Sub

Public Function DrawLine(ByVal nX&, ByVal nY&, Optional ByVal bRelative As Boolean = True) As Boolean

    If m.hDC = 0 Then Exit Function
        
    If bRelative Then
        m.nCurX = m.nCurX + nX
        m.nCurY = m.nCurY + nY
    Else
        m.nCurX = nX
        m.nCurY = nY
    End If
    
    If LineTo(m.hDC, m.nCurX, m.nCurY) <> 0 Then
        DrawLine = True
    End If

End Function

Public Function DrawRectangle(ByVal nX&, ByVal nY&, Optional ByVal bRelative As Boolean = True) As Boolean

    Dim nFromX&, nFromY&

    If m.hDC = 0 Then Exit Function

    nFromX = m.nCurX
    nFromY = m.nCurY

    If bRelative Then
        m.nCurX = m.nCurX + nX
        m.nCurY = m.nCurY + nY
    Else
        m.nCurX = nX
        m.nCurY = nY
    End If

    If Rectangle(m.hDC, nFromX, nFromY, m.nCurX, m.nCurY) <> 0 Then
        DrawRectangle = True
    End If

End Function
