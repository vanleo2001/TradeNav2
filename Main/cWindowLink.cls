VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWindowLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    nOldFormWinProc As Long
    frm As Form
    
    nSymbolColor As Long
    nPeriodColor As Long
    
    ToolTip As cToolTip
    
    bTestMethod As Boolean
    frmButtons As frmWindowLinkButtons
End Type
Private m As mPrivate

Public Sub Init(frm As Form)
        
    Dim i&
            
    If IsIDE Then
        'm.bTestMethod = True
    End If
        
    ' don't re-init if already initialized
    If frm.hWnd <> 0 And m.nOldFormWinProc = 0 Then
        Set m.frm = frm
        If SubclassingEnabled Then
            m.nOldFormWinProc = ezSubclass(frm.hWnd, Me)
        
            Set m.ToolTip = New cToolTip
            'm.ToolTip.Create frm.hWnd
            
            ' Draw buttons now if form is visible (need to do this when show a dockable window after hidden)
            DrawTitleBarButtons
            
            If m.bTestMethod Then
                Set m.frmButtons = New frmWindowLinkButtons
                m.frmButtons.InitButtons
                m.frmButtons.Top = -9000
                If TypeOf m.frm Is frmChart Then
                    'm.frmButtons.Show False, frmMain
                    'i = SetWindowLong(m.frmButtons.hWnd, -8&, m.frm.hWnd)
                    'm.frmButtons.Visible = True
                Else
                    m.frmButtons.Show False, m.frm
                End If
            End If
        End If
    End If
    
End Sub

Public Sub Unhook()

    If m.nOldFormWinProc <> 0 Then
        ezUnSubclass m.frm.hWnd, m.nOldFormWinProc
        m.nOldFormWinProc = 0
        Set m.ToolTip = Nothing
    End If
    Set m.frm = Nothing
    Set m.frmButtons = Nothing

End Sub

Private Sub Class_Terminate()
    Unhook
End Sub

'TLB: to draw buttons on the title bar
Private Function DrawTitleBarButtons(Optional ByVal nXclicked& = kNullData, Optional ByVal bButtonDown As Boolean = False) As Long

#If 1 Then
    
    Dim bChart As Boolean, bSmallCaption As Boolean, bDetached As Boolean, bDrawIt As Boolean
    Dim hDC&, hWnd&, xSize&, ySize&, nButtonsOver&, iButton&
    Dim BorderX&, BorderY&, nBtnLeft&
    Dim uFlags As Long
    Dim r As Rect
    Dim Paint As New cPaint

    If g.bUnloading Or g.bLoadingChartPage Or Not SubclassingEnabled Then Exit Function
    hWnd = m.frm.hWnd
    If IsWindowVisible(hWnd) = 0 Then Exit Function
    If IsFrmChart(m.frm) Then
        Select Case m.frm.DetachStatus
        Case eNotDetached
            If m.frm.WindowState <> 0 Then Exit Function
        Case eDetached
            bDetached = True
        End Select
        bChart = True
    End If
    
    uFlags = GetWindowLong(hWnd, GWL_EXSTYLE)
    If (uFlags And WS_EX_TOOLWINDOW) = 0 Then
        ' normal size buttons
        If IsAtLeastVista Then
            ' if THEME is disabled for the non-client areas
            xSize = GetSystemMetrics(SM_CXSMSIZE) + 1
            ySize = GetSystemMetrics(SM_CYSIZE)
            If xSize < ySize Then
                xSize = ySize ' so these buttons are never narrower than height
            End If
        Else
            ' if THEME is enabled for the non-client areas
            xSize = GetSystemMetrics(SM_CXSIZE)
            ySize = GetSystemMetrics(SM_CYSIZE)
        End If
        nButtonsOver = 3
    Else
        ' small buttons for thin window title area
        xSize = GetSystemMetrics(SM_CXSMSIZE)
        ySize = GetSystemMetrics(SM_CYSMSIZE)
        nButtonsOver = 2
        bSmallCaption = True
    End If

    BorderX = GetSystemMetrics(SM_CXFRAME)
    BorderY = GetSystemMetrics(SM_CYFRAME)

    GetWindowRect hWnd, r
    nBtnLeft = r.Right - r.Left - xSize * nButtonsOver - BorderX + 1 - 2
    
    If nXclicked <> kNullData Then
        nXclicked = nXclicked + BorderX
    End If
    
    With Paint
        .Init hWnd
        
        If 0 Then ' TypeOf m.frm Is frmChart Then
            r.Left = nBtnLeft '+ xSize
            r.Top = BorderY + 2
            r.Right = r.Left + xSize - 2
            r.Bottom = r.Top + ySize - 4
            'DrawFrameControl .hDC, r, DFC_CAPTION, DFCS_CAPTIONMAX
        End If
        
        For iButton = 3 To 1 Step -1
            bDrawIt = False
            
            If bChart Then
                If iButton = 3 Then
                    bDrawIt = m.frm.AllowDetach
                Else
                    bDrawIt = Not m.frm.IsInGameMode And m.frm.Chart.TypeOfChart <> eTypeChart_Seasonal         '5048
                End If
            ElseIf iButton = 1 Then
                bDrawIt = True
            End If
            
            If bDrawIt Then
                ' set the button position
                nBtnLeft = nBtnLeft - xSize + 1 '2
                r.Left = nBtnLeft
                r.Top = BorderY + 2
                r.Right = r.Left + xSize - 2
                r.Bottom = r.Top + ySize - 4
    
                ' draw the button
                uFlags = DFCS_BUTTONPUSH
                If nXclicked <> kNullData Then
                    If nXclicked > r.Left And nXclicked < r.Right Then
                        If bButtonDown Then
                            uFlags = uFlags Or DFCS_PUSHED
                        End If
                        DrawTitleBarButtons = iButton '(pushed)
                    End If
                    ''Debug.Print "X = " & CStr(nXclicked) & ", L = " & CStr(r.Left) & ", R = " & CStr(r.Right)
                End If
                If bSmallCaption Then ' iButton >= 2 Then
                    DrawFrameControl .hDC, r, DFC_BUTTON, uFlags
                Else
                    .MoveTo r.Left, r.Top, False
                    .SetBrush GetSysColor(COLOR_3DFACE), True
                    .DrawRectangle r.Right - 1, r.Bottom - 1, False
                
                    .MoveTo r.Left, r.Bottom - 1, False
                    .SetPen GetSysColor(COLOR_3DHILIGHT)
                    .DrawLine r.Left, r.Top, False
                    .DrawLine r.Right - 1, r.Top, False
                    .SetPen GetSysColor(COLOR_3DDKSHADOW)
                    .DrawLine r.Right - 1, r.Bottom - 1, False
                    .DrawLine r.Left - 1, r.Bottom - 1, False
                    
                    .MoveTo r.Left + 1, r.Bottom - 2, False
                    .SetPen GetSysColor(COLOR_3DLIGHT)
                    .DrawLine r.Left + 1, r.Top + 1, False
                    .DrawLine r.Right - 2, r.Top + 1, False
                    .SetPen GetSysColor(COLOR_3DSHADOW)
                    .DrawLine r.Right - 2, r.Bottom - 2, False
                    .DrawLine r.Left, r.Bottom - 2, False
                End If
            
#If 1 Then
                Select Case iButton
                Case 1 'Symbol Link
                    ' draw the colored rectangle in the button
                    If m.nSymbolColor > 0 Then
                        .SetBrush m.nSymbolColor, True
                        .MoveTo r.Left + 1, r.Top + 1, False
                        .DrawRectangle r.Right - 2, r.Bottom - 2, False
                        .SetPen RGB(255, 255, 255)
                    Else
                        .SetPen 0
                    End If
            
                    ' draw the "S" (7x7) in the middle of the rectangle
                    .MoveTo (r.Left + r.Right) \ 2 + 2, (r.Top + r.Bottom + 1) \ 2 - 3, False
                    .DrawLine -1, 0
                    .DrawLine 0, -1
                    .DrawLine -4, 0
                    .DrawLine 0, 1
                    .DrawLine -1, 0
                    .DrawLine 0, 1
                    .DrawLine 1, 0
                    
                    '.DrawLine 1, 1
                    '.DrawLine 2, 0
                    '.DrawLine 1, 1
                    .DrawLine 0, 1
                    .DrawLine 4, 0
                    .DrawLine 0, 1
                    
                    .DrawLine 1, 0
                    .DrawLine 0, 1
                    .DrawLine -1, 0
                    .DrawLine 0, 1
                    .DrawLine -4, 0
                    .DrawLine 0, -1
                    .DrawLine -2, 0
                    
                Case 2 'Period link
                    ' draw the colored rectangle in the button
                    If m.nPeriodColor > 0 Then
                        .SetBrush m.nPeriodColor, True
                        .MoveTo r.Left + 1, r.Top + 1, False
                        .DrawRectangle r.Right - 2, r.Bottom - 2, False
                        .SetPen RGB(255, 255, 255)
                    Else
                        .SetPen 0
                    End If
                
If 0 Then
                    ' draw the "A" (7x7) in the middle of the rectangle
                    .MoveTo (r.Left + r.Right) \ 2 - 2, (r.Top + r.Bottom + 1) \ 2 - 4, False
                    .DrawLine 2, 0
                    .DrawLine 0, 1
                    .DrawLine -2, 0
                    .DrawLine 0, 3
                    .DrawLine -2, 0
                    .DrawLine 0, 2
                    .DrawLine 1, 0
                    .DrawLine 0, -4
                    .DrawLine 1, 0
                    .DrawLine 0, 2
                    .DrawLine 4, 0
                    .DrawLine 0, 2
                    .DrawLine -1, 0
                    .DrawLine 0, -4
                    .DrawLine -1, 0
                    .DrawLine 0, 2
ElseIf 0 Then
                    ' draw the "D" (7x7) in the middle of the rectangle
                    .MoveTo (r.Left + r.Right) \ 2 - 4, (r.Top + r.Bottom + 1) \ 2 - 4, False
                    .DrawLine 0, 6
                    .DrawLine 1, 0
                    .DrawLine 0, -6
                    .DrawLine 3, 0
                    .DrawLine 0, 1
                    .DrawLine 1, 0
                    .DrawLine 0, 3
                    .DrawLine 1, 0
                    .DrawLine 0, -2
                    .DrawLine -1, 0
                    .DrawLine 0, 3
                    .DrawLine -1, 0
                    .DrawLine 0, 1
                    .DrawLine -4, 0
Else
                    ' draw the "P" (7x7) in the middle of the rectangle
                    .MoveTo (r.Left + r.Right) \ 2 + 2, (r.Top + r.Bottom + 1) \ 2 - 3, False
                    .DrawLine -1, 0
                    .DrawLine 0, -1
                    .DrawLine -5, 0
                    .DrawLine 0, 6
                    .DrawLine 1, 0
                    .DrawLine 0, -5
                    .DrawLine 0, 2
                    .DrawLine 4, 0
                    .DrawLine 0, -1
                    .DrawLine 2, 0
End If
                
                Case 3 'Detached status
If 1 Then
                    If bDetached Then
                        ' draw the Attach (7x7) in the middle of the rectangle
                        '.SetPen vbBlack
                        .SetPen RGB(96, 96, 96)
                        .MoveTo (r.Left + r.Right) \ 2 - 5, (r.Top + r.Bottom + 1) \ 2 - 4, False
                        .DrawRectangle 8, 8
                        .MoveTo (r.Left + r.Right) \ 2 - 5, (r.Top + r.Bottom + 1) \ 2 - 3, False
                        .DrawLine 8, 0
                        '.SetPen vbRed
                        .SetPen RGB(0, 192, 0)
                        .MoveTo (r.Left + r.Right) \ 2 - 3, (r.Top + r.Bottom + 1) \ 2 - 2, False
                        .DrawLine 0, 3
                        .DrawLine 6, -6
                        .DrawLine 0, 1
                        .DrawLine -5, 5
                        .DrawLine 2, 0
                        .DrawLine -3, -3
                    Else
                        ' draw the Detach (7x7) in the middle of the rectangle
                        '.SetPen vbBlack
                        .SetPen RGB(96, 96, 96)
                        .MoveTo (r.Left + r.Right) \ 2 - 5, (r.Top + r.Bottom + 1) \ 2 - 3, False
                        .DrawRectangle 7, 7
                        .MoveTo (r.Left + r.Right) \ 2 - 5, (r.Top + r.Bottom + 1) \ 2 - 2, False
                        .DrawLine 7, 0
                        .SetPen vbRed
                        .MoveTo (r.Left + r.Right) \ 2 - 0, (r.Top + r.Bottom + 1) \ 2 - 5, False
                        .DrawLine 3, 0
                        .DrawLine -5, 5
                        .DrawLine 0, 1
                        .DrawLine 5, -5
                        .DrawLine 0, 2
                        .DrawLine -3, -3
                    End If
Else
                    If bDetached Then
                        ' draw the Attach (7x7) in the middle of the rectangle
                        .SetBrush RGB(160, 160, 160)
                        .SetPen vbBlack
                        .MoveTo (r.Left + r.Right) \ 2 - 5, (r.Top + r.Bottom + 1) \ 2 - 5, False
                        .DrawRectangle 9, 9
                        .SetBrush vbWhite
                        .SetPen vbBlue
                        .MoveTo (r.Left + r.Right) \ 2 - 3, (r.Top + r.Bottom + 1) \ 2 - 3, False
                        .DrawRectangle 6, 6
                        .SetPen vbRed
                        .MoveTo (r.Left + r.Right) \ 2 - 2, (r.Top + r.Bottom + 1) \ 2 + 0, False
                        .DrawLine 1, -1
                        .DrawLine 1, 1
                        .DrawLine 2, -2
                    Else
                        ' draw the Detach (7x7) in the middle of the rectangle
                        .SetBrush RGB(160, 160, 160)
                        .SetPen vbBlack
                        .MoveTo (r.Left + r.Right) \ 2 - 5, (r.Top + r.Bottom + 1) \ 2 - 5, False
                        .DrawRectangle 7, 7
                        .SetBrush vbWhite
                        .SetPen vbBlue
                        .MoveTo (r.Left + r.Right) \ 2 - 2, (r.Top + r.Bottom + 1) \ 2 - 2, False
                        .DrawRectangle 6, 6
                        '.MoveTo (r.Left + r.Right) \ 2 - 2, (r.Top + r.Bottom + 1) \ 2 - 1, False
                        '.DrawLine 5, 0
                        .SetPen vbRed
                        .MoveTo (r.Left + r.Right) \ 2 - 1, (r.Top + r.Bottom + 1) \ 2 + 1, False
                        .DrawLine 1, -1
                        .DrawLine 1, 1
                        .DrawLine 2, -2
                    End If
End If
                End Select
#End If
            End If
        Next 'iButton
    End With
    Set Paint = Nothing
    
'    m.frm.SetBtnTipRect r.Left - 5, -1 * (ySize + BorderY), r.Right - 5, 0

#End If

End Function

' convert lParam of windows message to a Point structure, and optionally converts from screen to client window
Private Function ConvertToPoint(ByVal lParam As Long, Optional ByVal hWndClient As Long = 0) As POINTAPI

    ' convert long to signed int's
    ConvertToPoint.X = LoWord(lParam, False)
    ConvertToPoint.Y = HiWord(lParam, False)
    
    If hWndClient <> 0 Then
        ScreenToClient hWndClient, ConvertToPoint
    End If

End Function

' returns bottom of button relative to client area of form
' (will be about 0 if no toolbar, but negative if a toolbar exists)
Private Function ButtonBottom() As Long
    
    Dim ySize&
    
    If (GetWindowLong(m.frm.hWnd, GWL_EXSTYLE) And WS_EX_TOOLWINDOW) = 0 Then
        ' normal size buttons
        ySize = GetSystemMetrics(SM_CYCAPTION)
    Else
        ' small buttons for thin window title area
        ySize = GetSystemMetrics(SM_CYSMCAPTION)
    End If
    ySize = ySize + GetSystemMetrics(SM_CYFRAME) + 2

    ButtonBottom = ySize * Screen.TwipsPerPixelY - (m.frm.Height - m.frm.ScaleHeight)
    
End Function

Private Function FixCaptionStyle() As Boolean
        
    Dim uFlags&, hWnd&
    
If 0 Then
'Public Const WS_MAXIMIZEBOX = &H10000
'Public Const WS_MINIMIZEBOX = &H20000
'Public Const WM_WINDOWPOSCHANGING = &H46
    If IsFrmChart(m.frm) Then
        hWnd = m.frm.hWnd
        uFlags = GetWindowLong(hWnd, GWL_EXSTYLE)
        If m.frm.WindowState = 0 Then
        'If 1 Then
            If (uFlags And WS_EX_TOOLWINDOW) = 0 Then
                SetWindowLong hWnd, GWL_EXSTYLE, (uFlags Or WS_EX_TOOLWINDOW)
                uFlags = GetWindowLong(hWnd, GWL_STYLE)
                SetWindowLong hWnd, GWL_STYLE, (uFlags Or &H10000 Or &H20000)
                SetWindowPos hWnd, 0, 0, 0, 0, 0, _
                        SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
                FixCaptionStyle = True
            End If
        Else 'If 0 Then
            If (uFlags And WS_EX_TOOLWINDOW) <> 0 Then
                SetWindowLong hWnd, GWL_EXSTYLE, (uFlags And Not WS_EX_TOOLWINDOW)
                uFlags = GetWindowLong(hWnd, GWL_STYLE)
                SetWindowLong hWnd, GWL_STYLE, (uFlags Or &H10000 Or &H20000)
                SetWindowPos hWnd, 0, 0, 0, 0, 0, _
                        SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
                FixCaptionStyle = True
            End If
        End If
        'If FixCaptionStyle Then FormResize m.frm
    End If
End If

End Function

Public Function WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Dim bCancel As Boolean, bShowSymbolLink As Boolean, i&
    Dim ptMouse As POINTAPI
    Static dMinMouseWheelMove#
    
If Not IsIDE Then On Error Resume Next

#If 0 Then
    If IsIDE And TypeOf m.frm Is frmSnapshot Then
        If frmTest.Visible And iMsg <> 132 And iMsg <> 32 And iMsg <> 160 Then
            frmTest.AddList Str(iMsg) & vbTab & Str(wParam) & vbTab & Str(lParam)
        End If
    End If
#End If
    
    Select Case iMsg
    'Case WM_NCPAINT ', WM_NCACTIVATE, WM_SETTEXT, WM_STYLECHANGED  ', WM_SIZE
    '    FixCaptionStyle
    
    'Case &H47 ', &H46  ', &H18  ', &H225 ' WM_WINDOWPOSCHANGING, WM_WINDOWPOSCHANGED, WM_SHOWWINDOW
    '    FixCaptionStyle
        
    Case WM_NCACTIVATE              '4883
        If Not g.bStarting And Not g.bUnloading And Not g.bLoadingChartPage Then
            If Not m.frm Is Nothing Then
                If wParam <> 0 Then
                    If IsFrmChart(m.frm) Then
                        If Not ActiveChart Is Nothing Then
                            If ActiveChart.DetachStatus = eDetached Then
                                If Not ActiveChart.SkipFocusFix Then
                                    If m.frm.DetachStatus = eNotDetached Then
                                        bCancel = True
                                        SendMessage m.frm.hWnd, WM_ACTIVATE, 0, 0
                                        SendMessage ActiveChart.hWnd, WM_ACTIVATE, 1, 0
                                        SendMessage ActiveChart.hWnd, WM_NCACTIVATE, 1, 0
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
            
    Case WM_MOUSEACTIVATE
        If Not m.frm Is Nothing Then
            If IsFrmChart(m.frm) Then ActiveChartFormSet m.frm
        End If

        If GetActiveWindow <> hWnd Then
            ' Can't do this for charts
            'SetFocus hWnd
            'bCancel = True
        End If
       
'    Case WM_NCLBUTTONDOWN
'        If wParam = HTCAPTION Then
'            ptMouse = ConvertToPoint(lParam, hWnd)
'            If DrawTitleBarButtons(ptMouse.X, True) <> 0 Then
'                bCancel = True
'            End If
'        End If
        
    ' TLB 4/2/2007: the ButtonUp does not trigger when the form is in negative territory,
    ' but the ButtonDown does -- so we have to use that event to trigger this
    ' (but still have to Cancel for either ButtonUp or ButtonDown due to dockable forms)
    Case WM_NCLBUTTONDOWN, WM_NCLBUTTONUP, WM_NCLBUTTONDBLCLK
        
        If FormIsLoaded("frmTemplatePage") Then frmTemplatePage.ToolBarNotify
        If FormIsLoaded("frmTbMoreButtons") Then Unload frmTbMoreButtons
        
        If wParam = HTCAPTION Then
            ptMouse = ConvertToPoint(lParam, hWnd)
            Select Case DrawTitleBarButtons(ptMouse.X)
            Case 1:
                bCancel = True
                If iMsg = WM_NCLBUTTONDOWN And Not m.frm Is Nothing Then
                    frmWindowLink.ShowMe m.frm, eLink_Symbol, ptMouse.X * Screen.TwipsPerPixelX, ButtonBottom
                End If
            Case 2:
                bCancel = True
                If iMsg = WM_NCLBUTTONDOWN And Not m.frm Is Nothing Then
                    frmWindowLink.ShowMe m.frm, eLink_Period, ptMouse.X * Screen.TwipsPerPixelX, ButtonBottom
                End If
            Case 3:
                bCancel = True
                If iMsg = WM_NCLBUTTONDOWN And Not m.frm Is Nothing Then
                    If m.frm.DetachStatus = eDetached Then
                        'If InfBox("Attach this chart back into the main application window?", "?", "+Attach|-Cancel", "Attach Chart") <> "C" Then
                        AttachChart m.frm
                        'End If
                    Else
                        'If InfBox("Detach this chart from the main application window?", "?", "+Detach|-Cancel", "Detach Chart") <> "C" Then
                        DetachChart m.frm
                        'End If
                    End If
                End If
            End Select
        End If
        
    Case WM_MOUSEWHEEL ' TLB 9/25/2009: allow the mouse wheel to scroll the chart
        If IsFrmChart(m.frm) Then
            'StatusMsg "Mousewheel: " & Str(wParam) & " " & Str(lParam) & " " & Str(HiWord(wParam, False)) & " " & Str(LoWord(wParam, False))
            bCancel = True
            i = HiWord(wParam, False) ' (can be negative or positive)
            If i <> 0 Then
                If KeyIsPressed(VK_CONTROL) Then
                    ' TLB 1/22/2014: if Ctrl key is pressed, mouse scroll will do more/less bars
                    If i < 0 Then
                        ' more bars
                        m.frm.Chart.PixelsPerBar = -2
                        m.frm.Chart.GenerateChart eRedo1_Scrolled
                    Else
                        ' less bars
                        m.frm.Chart.PixelsPerBar = -1
                        m.frm.Chart.GenerateChart eRedo1_Scrolled
                    End If
                Else
                    ' allow faster wheel rotations to scroll more bars at a time
                    ' (so keep track of smallest rotation and associate that with a "1 bar scroll")
                    If Abs(i) < dMinMouseWheelMove Or dMinMouseWheelMove = 0 Then
                        dMinMouseWheelMove = Abs(i)
                    End If
                    i = m.frm.hsb.Value + Round(m.frm.hsb.SmallChange * (i / dMinMouseWheelMove))
                    If i > m.frm.hsb.Max Then
                        i = m.frm.hsb.Max
                    ElseIf i < m.frm.hsb.Min Then
                        i = m.frm.hsb.Min
                    End If
                    m.frm.hsb.Value = i
                End If
            End If
        End If
    End Select
    
    If Not bCancel Then WindowProc = CallWindowProc(m.nOldFormWinProc, hWnd, iMsg, wParam, ByVal lParam)

    Select Case iMsg
    Case WM_NCPAINT, WM_NCACTIVATE, WM_SETTEXT, WM_STYLECHANGED  ', WM_SIZE
        DrawTitleBarButtons
        MoveButtons

    Case WM_NCMOUSEMOVE
        If wParam = HTCAPTION Then
            'StatusMsg Str(wParam) & ", " & Str(lParam)
        End If
        
    Case WM_WINDOWPOSCHANGED, WM_SIZE
        MoveButtons
    End Select

End Function

' Symbol Link Color:
' 0 = unlinked
' 1 = link to active chart
' RGB = link to windows with same RGB color
Public Property Get SymbolColor() As Long
    SymbolColor = m.nSymbolColor
End Property

Public Property Let SymbolColor(ByVal nColor As Long)
    If nColor = &H101010 Then nColor = 0 '(old black color no longer used)
    m.nSymbolColor = nColor
End Property

Public Property Get PeriodColor() As Long
    PeriodColor = m.nPeriodColor
End Property

Public Property Let PeriodColor(ByVal nColor As Long)
    m.nPeriodColor = nColor
End Property


Private Sub MoveButtons()
    
    Dim pt As POINTAPI
    Dim i&, iX&, iY&
    
    If m.bTestMethod And Not m.frmButtons Is Nothing Then
        pt.X = 0
        pt.Y = 0
        ClientToScreen m.frm.hWnd, pt
        
        iX = pt.X * Screen.TwipsPerPixelX
        
        iY = (pt.Y + GetSystemMetrics(SM_CYFRAME) * 2) * Screen.TwipsPerPixelY
        iY = iY - (m.frm.Height - m.frm.ScaleHeight)
        
        i = (GetSystemMetrics(SM_CYCAPTION) * Screen.TwipsPerPixelY - m.frmButtons.Height) / 2
        If i < 0 Then i = 0
        
        m.frmButtons.Move iX, iY
    End If

End Sub
