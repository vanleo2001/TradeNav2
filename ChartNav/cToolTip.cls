VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cToolTip.class
'' Description: Implementation for windows registered class TOOLTIP_CLASS
''
'' Origination: Code downloaded from CodeGuru.com
''              Original files had no credit/author information
''
'' Modified:    Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'' Technical Reference:  This file and its companion file mTooltip.bas
''      implement wrappers for a subset of the functionalities of the TOOLTIP_CLASS
''      window class as documented in Microsoft MSDN libary. For details and/or
''      background information refer to MSDN: key word "tooltip controls"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author          Description
'' 05-12-2005   J. Muchmore     adapted original for use by NavSuite project
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Usage:
''  a) add an object of this class to a form
''  b) call the Create() function of this class prior to using the object
''  c) to have tooltip for a control on the form call the AddTool() function
''  d) to have tooltip for a rectangular region on the form call the AddRect() function
''  e) to stop showing tooltip call RemoveTool() or RemoveRect() accordingly
''  f) set the object to nothing to trigger the class terminate routine and clean up
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' Defaults:
' DelayInitial   = 500  (1/2 sec)
' DelayAutoPopup = 5000  (5 secs)
' DelayReshow    = 100 (1/10 sec)
' MaxTipWidth    = 0
' all Margins    = 0

Private Type mPrivate
    hwndToolTip As Long
    nMaxTipLen As Long
    tbRects As cGdTable
End Type
Private m As mPrivate

Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.tbRects = New cGdTable
    
    m.tbRects.CreateField eGDARRAY_Longs, 0, "hWnds"
    m.tbRects.CreateField eGDARRAY_Longs, 1, "left"
    m.tbRects.CreateField eGDARRAY_Longs, 2, "top"
    m.tbRects.CreateField eGDARRAY_Longs, 3, "right"
    m.tbRects.CreateField eGDARRAY_Longs, 4, "bottom"
    m.tbRects.CreateField eGDARRAY_Strings, 5, "tiptext"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cToolTip.Class_Initialize", eGDRaiseError_Raise

End Sub

Private Sub Class_Terminate()
On Error GoTo ErrSection:

    If m.hwndToolTip > 0 Then Call DestroyWindow(m.hwndToolTip)
    Set m.tbRects = Nothing
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cToolTip.Class_Terminate", eGDRaiseError_Raise

End Sub

Public Function Create(ByRef frm As Form) As Boolean
On Error GoTo ErrSection:
  
    Dim ccex As COMMONCONTROLSEX
    
    If frm Is Nothing Then Exit Function
    
    If (m.hwndToolTip = 0) Then
        ccex.dwSize = Len(ccex)
        ccex.dwICC = ICC_TAB_CLASSES
        If InitCommonControlsEx(ccex) Then
            '
            ' The hwndParent param lets the tooltip window
            ' be owned by the specified form and be destroyed
            ' along with it. We'll cleanup in Class_Terminate anyway.
            ' No WS_EX_TOPMOST or TTS_ALWAYSTIP per Win95 UI rules.
            '
            m.hwndToolTip = CreateWindowEx(0, TOOLTIPS_CLASS, vbNullString, TTS_ALWAYSTIP Or TTS_NOPREFIX, 0, 0, 0, 0, frm.hWnd, 0, App.hInstance, ByVal 0)
        End If
    End If
    
    Create = CBool(m.hwndToolTip)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cToolTip.Create", eGDRaiseError_Raise

End Function

Public Function AddTool(ByRef ctrl As Control, Optional ByVal strText As String) As Boolean
On Error GoTo ErrSection:

    Dim ti As TOOLINFO
  
    If m.hwndToolTip = 0 Or ctrl Is Nothing Then Exit Function
    If ctrl.hWnd = 0 Then Exit Function
    
    If (fGetToolInfo(ctrl.hWnd, ti) = False) Then
        With ti
            .cbSize = Len(ti)
            '
            ' TTF_IDISHWND must be set to tell the tooltip
            ' control to retrieve the control's rect from
            ' it's hWnd specified in uId.
            '
            .uFlags = TTF_SUBCLASS Or TTF_IDISHWND
            .hWnd = ctrl.Container.hWnd
            .uId = ctrl.hWnd
            
            If Len(strText) > 0 Then
                .lpszText = strText
'            Else
'                .lpszText = "Tool" & ToolCount + 1
            End If
            
            '
            ' Maintain the maximun tip text
            ' length for fGetToolInfo.
            '
            m.nMaxTipLen = fMax(m.nMaxTipLen, Len(.lpszText) + 1)
        End With
        
        '
        ' Returns 1 on success, 0 on failure
        '
        AddTool = SendMessageT(m.hwndToolTip, TTM_ADDTOOL, 0, ti)
    End If
  
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cToolTip.AddTool", eGDRaiseError_Raise

End Function

Public Function AddRect(ByRef vFormOrCtrl As Variant, _
    ByVal nLeft&, nTop&, nRight&, nBottom&, _
    Optional ByVal strText As String, _
    Optional ByVal bMultirect As Boolean) As Boolean
On Error GoTo ErrSection:

    Dim ti As TOOLINFO
  
    If m.hwndToolTip = 0 Or vFormOrCtrl Is Nothing Then Exit Function
    If Not TypeOf vFormOrCtrl Is Form And Not TypeOf vFormOrCtrl Is Control Then Exit Function
    If vFormOrCtrl.hWnd = 0 Then Exit Function
    
    If (fGetToolInfo(vFormOrCtrl.hWnd, ti) = True) Then
        If bMultirect Then
            AddRect = AddSubRect(vFormOrCtrl, nLeft, nTop, nRight, nBottom, strText)
        End If
        Exit Function
    End If
    
    With ti
        .cbSize = Len(ti)
        '
        ' values for the rRect must be in pixels and in client coordinates
        ' - i.e. coordinates are relative to the form or control
        ' for forms/controls, client coordinates of upperleft is always 0,0
        ' - the Left, Top, Width, Height shown in the VB IDE properties window
        '   for forms and controls are screen coordinates and generally in twips
        '
        .uFlags = TTF_SUBCLASS
        .hWnd = vFormOrCtrl.hWnd
        .uId = vFormOrCtrl.hWnd
        .rRect.Left = nLeft
        .rRect.Top = nTop
        .rRect.Right = nRight
        .rRect.Bottom = nBottom
        
        If Len(strText) > 0 Then
            .lpszText = strText
'            Else
'                .lpszText = "Tool" & ToolCount + 1
        End If
        
        '
        ' Maintain the maximun tip text
        ' length for fGetToolInfo.
        '
        m.nMaxTipLen = fMax(m.nMaxTipLen, Len(.lpszText) + 1)
    End With
    
    '
    ' Returns 1 on success, 0 on failure
    '
    AddRect = SendMessageT(m.hwndToolTip, TTM_ADDTOOL, 0, ti)
  
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cToolTip.AddRect", eGDRaiseError_Raise

End Function


Private Function fMax(ByVal lngParm1 As Long, ByVal lngParm2 As Long) As Long
On Error GoTo ErrSection:
    '
    ' Returns the larger of the two values.
    '
    If lngParm1 > lngParm2 Then
        fMax = lngParm1
    Else
        fMax = lngParm2
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cToolTip.fMax", eGDRaiseError_Raise

End Function

Public Function RemoveTool(ByRef ctrl As Control) As Boolean
On Error GoTo ErrSection:

    Dim ti As TOOLINFO

    If m.hwndToolTip = 0 Or ctrl Is Nothing Then Exit Function
    If ctrl.hWnd = 0 Then Exit Function
    
    If fGetToolInfo(ctrl.hWnd, ti) Then
        Call SendMessageT(m.hwndToolTip, TTM_DELTOOL, 0, ti)
        RemoveTool = True
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cToolTip.RemoveTool", eGDRaiseError_Raise

End Function

Public Function RemoveRect(ByRef vFormOrControl As Variant) As Boolean
On Error GoTo ErrSection:

    Dim ti As TOOLINFO
    Dim frmForm As Form
    Dim ctlControl As Control

    If m.hwndToolTip = 0 Or vFormOrControl Is Nothing Then Exit Function
    If Not TypeOf vFormOrControl Is Form And Not TypeOf vFormOrControl Is Control Then Exit Function
    If vFormOrControl.hWnd = 0 Then Exit Function
    
    If fGetToolInfo(vFormOrControl.hWnd, ti) Then
        Call SendMessageT(m.hwndToolTip, TTM_DELTOOL, 0, ti)
        RemoveRect = True
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cToolTip.RemoveRect", eGDRaiseError_Raise

End Function

Public Property Get BackColor() As OLE_COLOR
On Error GoTo ErrSection:
    
    If m.hwndToolTip = 0 Then Exit Property
    
    '
    ' OLE_COLOR is defined in stdole2.tlb
    '
    BackColor = SendMessageT(m.hwndToolTip, TTM_GETTIPBKCOLOR, 0, 0)
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cToolTip.BackColor.Get", eGDRaiseError_Raise

End Property

Public Property Let BackColor(Clr As OLE_COLOR)
On Error GoTo ErrSection:
    
    If m.hwndToolTip = 0 Then Exit Property
    Call SendMessageT(m.hwndToolTip, TTM_SETTIPBKCOLOR, Clr, 0)

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cToolTip.BackColor.Let", eGDRaiseError_Raise

End Property

Public Property Get DelayTime(dwType As ttDelayTimeConstants) As Long
On Error GoTo ErrSection:
    
    If m.hwndToolTip = 0 Then Exit Property
    DelayTime = SendMessageT(m.hwndToolTip, TTM_GETDELAYTIME, (dwType And ttDelayMask), 0&)
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cToolTip.DelayTime.Get", eGDRaiseError_Raise

End Property

Public Property Let DelayTime(dwType As ttDelayTimeConstants, dwMilliSecs As Long)
On Error GoTo ErrSection:
    
    If m.hwndToolTip = 0 Then Exit Property
    Call SendMessageT(m.hwndToolTip, TTM_SETDELAYTIME, (dwType And ttDelayMask), ByVal dwMilliSecs)  ' no rtn val

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cToolTip.DelayTime.Let", eGDRaiseError_Raise

End Property

Public Property Get ForeColor() As OLE_COLOR
On Error GoTo ErrSection:
    
    If m.hwndToolTip = 0 Then Exit Property
    ForeColor = SendMessageT(m.hwndToolTip, TTM_SETTIPTEXTCOLOR, 0, 0)

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cToolTip.ForeColor.Get", eGDRaiseError_Raise

End Property

Public Property Let ForeColor(Clr As OLE_COLOR)
On Error GoTo ErrSection:
    
    If m.hwndToolTip = 0 Then Exit Property
    Call SendMessageT(m.hwndToolTip, TTM_SETTIPTEXTCOLOR, Clr, 0)   ' no rtn val

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cToolTip.ForeColor.Let", eGDRaiseError_Raise

End Property

Public Property Get hWnd() As Long
On Error GoTo ErrSection:

    hWnd = m.hwndToolTip
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cToolTip.hWnd.Get", eGDRaiseError_Raise

End Property

Public Property Get Margin(dwType As ttMarginConstants) As Long
On Error GoTo ErrSection:

    Dim rc As Rect

    If m.hwndToolTip = 0 Then Exit Property
    
    Call SendMessageT(m.hwndToolTip, TTM_GETMARGIN, 0, rc)
    
    Select Case dwType
        Case ttMarginLeft
            Margin = rc.Left
        
        Case ttMarginTop
            Margin = rc.Top
        
        Case ttMarginRight
            Margin = rc.Right
        
        Case ttMarginBottom
            Margin = rc.Bottom
    End Select

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cToolTip.Margin.Get", eGDRaiseError_Raise

End Property

Public Property Let Margin(dwType As ttMarginConstants, cPixels As Long)
On Error GoTo ErrSection:
    
    Dim rc As Rect
    
    If m.hwndToolTip = 0 Then Exit Property
    
    Call SendMessageT(m.hwndToolTip, TTM_GETMARGIN, 0, rc)
    
    Select Case dwType
        Case ttMarginLeft
            rc.Left = cPixels
            
        Case ttMarginTop
            rc.Top = cPixels
            
        Case ttMarginRight
            rc.Right = cPixels
            
        Case ttMarginBottom
            rc.Bottom = cPixels
    End Select
    
    Call SendMessageT(m.hwndToolTip, TTM_SETMARGIN, 0, rc)

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cToolTip.Margin.Let", eGDRaiseError_Raise

End Property

Public Property Get MaxTipWidth() As Long
On Error GoTo ErrSection:

    '
    ' If MaxTipWidth is -1, there is no word wrapping and
    ' text control characters are printed and not
    ' evaluated (i.e. a vbCrLf shows up as "||")
    '
    If m.hwndToolTip = 0 Then Exit Property
    MaxTipWidth = fLowWord(SendMessageT(m.hwndToolTip, TTM_GETMAXTIPWIDTH, 0, 0))

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cToolTip.MaxTipWidth.Get", eGDRaiseError_Raise

End Property

Public Property Let MaxTipWidth(ByVal lngWidth As Long)
On Error GoTo ErrSection:
    '
    ' If MaxTipWidth is -1, there is no word wrapping and
    ' text control characters are printed and not
    ' evaluated (i.e. a vbCrLf shows up as "||")
    '
    If m.hwndToolTip = 0 Then Exit Property
    If lngWidth < 1 Then lngWidth = -1
    Call SendMessageT(m.hwndToolTip, TTM_SETMAXTIPWIDTH, 0, lngWidth)

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cToolTip.MaxTipWidth.Let", eGDRaiseError_Raise

End Property

Private Function fLowWord(ByVal lngValue As Long) As Integer
On Error GoTo ErrSection:
    '
    ' Returns the low-order word from a 32-bit value.
    '
    Call MoveMemory(fLowWord, lngValue, 2)
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cToolTip.fLowWord", eGDRaiseError_Raise

End Function

Public Property Get ToolCount() As Long
On Error GoTo ErrSection:
    
    If m.hwndToolTip = 0 Then Exit Property
    ToolCount = SendMessageT(m.hwndToolTip, TTM_GETTOOLCOUNT, 0, 0)

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cToolTip.ToolCount.Get", eGDRaiseError_Raise

End Property

Public Property Get ToolText(ByRef ctrl As Control) As String
On Error GoTo ErrSection:

    Dim ti As TOOLINFO
  
    If m.hwndToolTip = 0 Or ctrl Is Nothing Then Exit Property
    If ctrl.hWnd = 0 Then Exit Property
    
    If fGetToolInfo(ctrl.hWnd, ti, True) Then
        ToolText = fGetStrFromBuffer(ti.lpszText)
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cToolTip.ToolText.Get", eGDRaiseError_Raise

End Property

Public Property Let ToolText(ByRef ctrl As Control, ByVal strText As String)
On Error GoTo ErrSection:

    Dim ti As TOOLINFO
  
    If m.hwndToolTip = 0 Or ctrl Is Nothing Then Exit Property
    If ctrl.hWnd = 0 Then Exit Property
    If ToolText(ctrl) = strText Then Exit Property
    
    If fGetToolInfo(ctrl.hWnd, ti) Then
        ti.lpszText = strText
        m.nMaxTipLen = fMax(m.nMaxTipLen, Len(strText) + 1)
        '
        ' The tooltip won't appear for the control
        ' if lpszText is an empty string
        '
        Call SendMessageT(m.hwndToolTip, TTM_UPDATETIPTEXT, 0, ti)
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cToolTip.ToolText.Let", eGDRaiseError_Raise

End Property

Public Property Get RectText(ByRef vFormOrControl As Variant) As String
On Error GoTo ErrSection:

    Dim ti As TOOLINFO
    Dim frmForm As Form
    Dim ctlControl As Control
  
    If m.hwndToolTip = 0 Or vFormOrControl Is Nothing Then Exit Property
    If Not TypeOf vFormOrControl Is Form And Not TypeOf vFormOrControl Is Control Then Exit Function
    If vFormOrControl.hWnd = 0 Then Exit Function
    
    If fGetToolInfo(vFormOrControl.hWnd, ti, True) Then
        RectText = fGetStrFromBuffer(ti.lpszText)
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cToolTip.RectText.Get", eGDRaiseError_Raise

End Property

Public Property Let RectText(ByRef vFormOrControl As Variant, ByVal strText As String)
On Error GoTo ErrSection:

    Dim ti As TOOLINFO
    Dim strCurrTip$
    
    If m.hwndToolTip = 0 Or vFormOrControl Is Nothing Then Exit Property
    If Not TypeOf vFormOrControl Is Form And Not TypeOf vFormOrControl Is Control Then Exit Property
    If vFormOrControl.hWnd = 0 Then Exit Property
    
    If fGetToolInfo(vFormOrControl.hWnd, ti, True) Then
        strCurrTip = fGetStrFromBuffer(ti.lpszText)
        If strCurrTip <> strText Then
            ti.lpszText = strText
            m.nMaxTipLen = fMax(m.nMaxTipLen, Len(strText) + 1)
            Call SendMessageT(m.hwndToolTip, TTM_UPDATETIPTEXT, 0, ti)
        End If
    End If
    
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cToolTip.RectText.Let", eGDRaiseError_Raise

End Property

Private Function fGetStrFromBuffer(ByVal strValue As String) As String
On Error GoTo ErrSection:
    
    If InStr(strValue, vbNullChar) Then
        fGetStrFromBuffer = Left$(strValue, InStr(strValue, vbNullChar) - 1)
    Else
        '
        ' If strValue had no null char, the Left$ function
        ' above would rtn a zero length string ("").
        '
        fGetStrFromBuffer = strValue
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cToolTip.fGetStrFromBuffer", eGDRaiseError_Raise

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''original code (appears not needed, leaving here for reference)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''Private Function fIsWindow(ByRef ctrl As Control) As Boolean
''On Error GoTo ErrSection:
''
''    fIsWindow = CBool(ctrl.hWnd)
''
''ErrExit:
''    Exit Function
''
''ErrSection:
''    RaiseError "cToolTip.fIsWindow", eGDRaiseError_Raise
''
''End Function
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function fGetToolInfo(ByVal lnghwndTool As Long, ti As TOOLINFO, _
    Optional fGetText As Boolean = False) As Boolean
On Error GoTo ErrSection:
  
    Dim nItems As Long
    Dim i      As Integer

    ti.cbSize = Len(ti)
    If fGetText Then ti.lpszText = String$(m.nMaxTipLen, 0)
    
    nItems = ToolCount
    
    For i = 0 To nItems - 1
        '
        ' Returns 1 on success, 0 on failure.
        '
        If SendMessageT(m.hwndToolTip, TTM_ENUMTOOLS, (i), ti) Then
            If (lnghwndTool = ti.uId) Then
                fGetToolInfo = True
                Exit Function
            End If
        End If
    Next

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cToolTip.fGetToolInfo", eGDRaiseError_Raise

End Function

Public Function ResizeRect(ByRef vFormOrControl As Variant, _
    ByVal nLeft&, ByVal nTop&, ByVal nRight&, ByVal nBottom&) As Boolean
On Error GoTo ErrSection:

    Dim ti As TOOLINFO
    Dim frm As Form
    Dim ctl As Control
    
    Dim bChanged As Boolean

    If m.hwndToolTip = 0 Or vFormOrControl Is Nothing Then Exit Function
    If Not TypeOf vFormOrControl Is Form And Not TypeOf vFormOrControl Is Control Then Exit Function
    If vFormOrControl.hWnd = 0 Then Exit Function
           
    If (fGetToolInfo(vFormOrControl.hWnd, ti) = True) Then
        If ti.rRect.Left <> nLeft Or _
           ti.rRect.Top <> nTop Or _
           ti.rRect.Right <> nRight Or _
           ti.rRect.Bottom <> nBottom Then
           
           bChanged = True
        End If
    End If
    
    If bChanged Then
        ti.rRect.Left = nLeft
        ti.rRect.Top = nTop
        ti.rRect.Right = nRight
        ti.rRect.Bottom = nBottom
        ResizeRect = SendMessageT(m.hwndToolTip, TTM_NEWTOOLRECT, 0, ti)
    End If
    

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cToolTip.ResizeRect", eGDRaiseError_Raise
    
End Function

Private Function IsPtInRect(rRect As Rect, pt As POINTAPI) As Boolean
On Error GoTo ErrSection:

    Dim bIsInRect As Boolean
    Dim r As Rect, i&
    
    r.Left = rRect.Left
    r.Top = rRect.Top
    r.Right = rRect.Right
    r.Bottom = rRect.Bottom

    'normalize the rect
    If r.Left > r.Right Then
        i = r.Left
        r.Left = r.Right
        r.Right = i
    End If
    If r.Top > r.Bottom Then
        i = r.Top
        r.Top = r.Bottom
        r.Bottom = i
    End If
    
    If pt.X >= r.Left And pt.X <= r.Right Then
        If pt.Y >= r.Top And pt.Y <= r.Bottom Then
            bIsInRect = True
        End If
    End If
    
    IsPtInRect = bIsInRect

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cToolTip.IsPtInRect", eGDRaiseError_Raise
    Resume ErrExit

End Function

Private Function AddSubRect(ByRef vFormOrControl As Variant, _
    ByVal nLeft&, ByVal nTop&, ByVal nRight&, ByVal nBottom&, _
    ByVal strTip$) As Boolean
On Error GoTo ErrSection:
    
    Dim aIdx As cGdArray
    Dim r As Rect, rNew As Rect
    Dim bFound As Boolean
    Dim i&
    
    
    SetRect rNew, nLeft, nTop, nRight, nBottom
    
    
    Set aIdx = m.tbRects.CreateSortedIndex(0, eGdSort_Default)
    If aIdx Is Nothing Then
        Exit Function
    End If
        
    For i = 0 To aIdx.Size - 1
        If vFormOrControl.hWnd = m.tbRects(0, aIdx(i)) Then
            SetRect r, m.tbRects(1, aIdx(i)), m.tbRects(2, aIdx(i)), m.tbRects(3, aIdx(i)), m.tbRects(4, aIdx(i))
            If r.Left = rNew.Left Then
                If r.Top = rNew.Top Then
                    If r.Right = rNew.Right Then
                        If r.Bottom = rNew.Bottom Then
                            bFound = True
                            Exit For
                        End If
                    End If
                End If
            End If
        End If
    Next
    
    If Not bFound Then
        m.tbRects.AddRecord ""
        i = m.tbRects.NumRecords - 1
        m.tbRects(0, i) = vFormOrControl.hWnd
        m.tbRects(1, i) = rNew.Left
        m.tbRects(2, i) = rNew.Top
        m.tbRects(3, i) = rNew.Right
        m.tbRects(4, i) = rNew.Bottom
        m.tbRects(5, i) = strTip
    End If
    
    Set aIdx = Nothing
    
    AddSubRect = True
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cToolTip.AddSubRect", eGDRaiseError_Raise

End Function

Public Sub UpdateMultiRectText(ByRef vFormOrControl As Variant)
On Error GoTo ErrSection:

    Dim pt As POINTAPI
    Dim aIdx As cGdArray
    Dim r As Rect
    Dim bFound As Boolean
    Dim i&
    
    If vFormOrControl Is Nothing Then Exit Sub
            
    Set aIdx = m.tbRects.CreateSortedIndex(0, eGdSort_Default)
    If aIdx Is Nothing Then Exit Sub
    
    If GetCursorPos(pt) <> 0 Then
        If ScreenToClient(vFormOrControl.hWnd, pt) <> 0 Then
            For i = 0 To aIdx.Size - 1
                If vFormOrControl.hWnd = m.tbRects(0, aIdx(i)) Then
                    SetRect r, m.tbRects(1, aIdx(i)), m.tbRects(2, aIdx(i)), m.tbRects(3, aIdx(i)), m.tbRects(4, aIdx(i))
                    If IsPtInRect(r, pt) Then
                        RectText(vFormOrControl) = m.tbRects(5, aIdx(i))
                        bFound = True
                        Exit For
                    End If
                Else
                    StatusMsg Str(pt.X) & ", " & Str(pt.Y)
                End If
            Next
        End If
    End If
    
    If Not bFound Then RectText(vFormOrControl) = ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cToolTip.UpdateMultiRectText", eGDRaiseError_Raise

End Sub

Public Sub ClearSubRects()
On Error GoTo ErrSection:

    m.tbRects.NumRecords = 0

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cToolTip.ClearSubRects", eGDRaiseError_Raise

End Sub

