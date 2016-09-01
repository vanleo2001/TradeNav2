VERSION 5.00
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmTbMoreButtons 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1485
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3885
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   885
      Top             =   915
   End
   Begin VB.PictureBox pbTbBack 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   585
      Index           =   0
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   3885
      TabIndex        =   0
      Top             =   0
      Width           =   3885
      Begin HexUniControls.ctlUniComboBoxXP cboBarPeriod 
         Height          =   315
         Left            =   3060
         TabIndex        =   1
         Top             =   120
         Visible         =   0   'False
         Width           =   885
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Tip             =   "frmTbMoreButtons.frx":0000
         Sorted          =   0   'False
         HScroll         =   0   'False
         Style           =   2
         ButtonBackColor =   -2147483633
         ButtonForeColor =   0
         ButtonWidth     =   17
         Locked          =   0   'False
         SelBackColor    =   -2147483635
         SelForeColor    =   -2147483634
         TrapTab         =   0   'False
         ButtonStyle     =   -1
         SelectorStyle   =   -1
         MousePointer    =   0
         MouseIcon       =   "frmTbMoreButtons.frx":0020
         DropDownOnTextClick=   -1  'True
         DropDownWidth   =   -1
         ManualStart     =   0   'False
         MaxLength       =   0
         RightToLeft     =   0   'False
         LeftMargin      =   0
         RightMargin     =   0
         SelectOnFocus   =   0   'False
      End
      Begin VB.Image imgTbBack 
         Height          =   585
         Index           =   0
         Left            =   780
         Stretch         =   -1  'True
         Top             =   105
         Width           =   11250
      End
   End
End
Attribute VB_Name = "frmTbMoreButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type mPrivate
    frm As Form                     'form that toolbar is on
    aButtons As New cGdArray
    oBtnMouseLast As cPicBoxButton
End Type
Private m As mPrivate

Public Function ShowMe(frmSource As Form, ByVal strToolbar$, ByVal nX&, ByVal nY&)
On Error GoTo ErrSection:

    Dim pt As POINTAPI
    
    Dim iTbWidth&, iDiff&, i&, j&, k&
    
    Dim aButtons As cGdArray
    Dim oBtn As cPicBoxButton
            
    If frmSource Is Nothing Then Exit Function
    
    Set m.frm = frmSource
    Set m.oBtnMouseLast = Nothing       '5224
    Set m.aButtons = Nothing
    
    Set aButtons = m.frm.TbButtonsArray(strToolbar)
    If aButtons Is Nothing Then Exit Function
    Set oBtn = aButtons(0)
    If oBtn Is Nothing Then Exit Function
    
    pt.X = nX
    pt.Y = nY
    
    If strToolbar = kTbDraw And m.frm.pbTbBackDraw(0).Visible Then
        ClientToScreen m.frm.pbTbBackDraw(0).hWnd, pt
    Else
        ClientToScreen m.frm.pbTbBack(0).hWnd, pt
    End If
    pt.X = pt.X * Screen.TwipsPerPixelX
    pt.Y = pt.Y * Screen.TwipsPerPixelY
    
    Me.cboBarPeriod.Visible = False
    
    InitForm2 aButtons, strToolbar
    
    j = 0
    For i = 0 To Me.pbTbBack.UBound
        j = j + Me.pbTbBack(i).Height
    Next
    
    k = 0
    For i = 0 To m.frm.pbTbBack.UBound
        If m.frm.pbTbBack(i).Visible Then
            k = k + m.frm.pbTbBack(i).Height    '5225
        Else
            Exit For
        End If
    Next
    
    Me.Height = j
    
    j = Me.pbTbBack.UBound + 1
        
    If oBtn.ToolBarPos = vbAlignBottom Then
        Me.Left = pt.X - Me.Width
        Me.Top = pt.Y + oBtn.BtnHeight
    ElseIf oBtn.ToolBarPos = vbAlignTop Then
        Me.Left = pt.X - Me.Width
        If oBtn.ToolBarName = kTbDraw Then
            Me.Top = pt.Y - k
        Else
            Me.Top = pt.Y + oBtn.BtnHeight
        End If
    ElseIf oBtn.ToolBarPos = vbAlignLeft Then
        'drawing toolbar on left
        Me.Left = m.frm.Left + (oBtn.BtnWidth + 3) * Screen.TwipsPerPixelX
        Me.Top = pt.Y - (oBtn.BtnHeight * j * Screen.TwipsPerPixelY) - (3 * Screen.TwipsPerPixelY)
    ElseIf oBtn.ToolBarPos = vbAlignRight Then
        'drawing toolbar on right
        Me.Left = (m.frm.Left + m.frm.Width) - ((oBtn.BtnWidth + 3) * Screen.TwipsPerPixelX) - Me.Width
        Me.Top = pt.Y - (oBtn.BtnHeight * j * Screen.TwipsPerPixelY) - (3 * Screen.TwipsPerPixelY)
    End If
    
    ShowForm Me
            
ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmTbMoreButtons.ShowMe"
    Unload Me
            
End Function

Private Sub InitForm2(aSourceBtns As cGdArray, strToolbar$)
On Error GoTo ErrSection:

    Dim i&
    
    Dim Chart As cChart
    Dim aShow As New cGdArray
    Dim oBtn As cPicBoxButton
    
    Dim bDisplayTypeDone As Boolean
    Dim bBarPeriodDone As Boolean
    
    Dim strSector$, strSub$, strCmp$, strSecDesc$, strSubDesc$, strCmpDesc$
    
    Set oBtn = aSourceBtns(0)
    
    If IsFrmChart(m.frm) Then
        Set Chart = m.frm.Chart
    ElseIf Not ActiveChart Is Nothing Then
        Set Chart = ActiveChart.Chart
    End If
    
    For i = oBtn.LastVisibleBtnIndex + 1 To aSourceBtns.Size - 2
        Set oBtn = aSourceBtns(i)
        m.aButtons.Add oBtn
        aShow.Add oBtn.BtnID
        
        If oBtn.BtnID = "ID_BarPeriod" Then
            BarPeriodCboInit cboBarPeriod
            cboBarPeriod.Text = ""
            
            If Not Chart Is Nothing Then
                cboBarPeriod.Text = GetPeriodStr(Chart.Periodicity)
            End If
        End If
        
    Next
    
    ToolbarInit2 Me, m.aButtons, aShow, strToolbar
    
    For i = 0 To m.aButtons.Size - 1
        Set oBtn = m.aButtons(i)
        If oBtn.BtnGroup = eBtnGroup_BarDisplayType Then
            If Not bDisplayTypeDone And Not Chart Is Nothing Then
                If Chart.tbToolbar.Tools(oBtn.BtnID).State = ssChecked Then
                    oBtn.BtnState = eBtnState_Selected
                    bDisplayTypeDone = True
                End If
            End If
        ElseIf oBtn.BtnGroup = eBtnGroup_BarPeriods Then
            If Not bBarPeriodDone And Not Chart Is Nothing Then
                If Chart.tbToolbar.Tools(oBtn.BtnID).State = ssChecked Then
                    oBtn.BtnState = eBtnState_Selected
                    bBarPeriodDone = True
                End If
            End If
        Else
            Select Case oBtn.BtnID
                Case "ID_Quote"
                    If frmQuotes.Visible Then oBtn.BtnState = eBtnState_Selected
                Case "ID_SymbolGrid"
                    If frmSymbolGrid.Visible Then oBtn.BtnState = eBtnState_Selected
                Case "ID_SectorBrowser"
                    If frmSectorTree.Visible Then oBtn.BtnState = eBtnState_Selected
                Case "ID_ChartOnOff"
                    If frmChartOnOff.Visible Then oBtn.BtnState = eBtnState_Selected
                Case "ID_ChartData"
                    If frmChartData.Visible Then oBtn.BtnState = eBtnState_Selected
                Case "ID_PlanetData"
                    If frmPlanetData.Visible Then oBtn.BtnState = eBtnState_Selected
                Case "ID_Snapshot"
                    If frmSnapshot.Visible Then oBtn.BtnState = eBtnState_Selected
                Case "ID_TradeTracker"
                    If frmTTSummary.Visible Then oBtn.BtnState = eBtnState_Selected
                Case "ID_Orders"
                    If frmNextBar.Visible Then oBtn.BtnState = eBtnState_Selected
                Case "ID_WhatIf"
                    If Not Chart Is Nothing Then
                        If Chart.IsInWhatIfMode Then oBtn.BtnState = eBtnState_Selected         '5447
                    End If
                Case "ID_ChartOrderbar"
                    If Not Chart Is Nothing Then
                        If Chart.ShowTrades = 2 Then oBtn.BtnState = eBtnState_Selected
                    End If
                'JP buttons
                Case "ID_JPDaily"
                    If Chart.tbToolbar.Tools("ID_JPDaily").State = ssChecked Then oBtn.BtnState = eBtnState_Selected
                Case "ID_JPWeekly"
                    If Chart.tbToolbar.Tools("ID_JPWeekly").State = ssChecked Then oBtn.BtnState = eBtnState_Selected
                Case "ID_JPMonthly"
                    If Chart.tbToolbar.Tools("ID_JPMonthly").State = ssChecked Then oBtn.BtnState = eBtnState_Selected
                Case "ID_JPQuarterly"
                    If Chart.tbToolbar.Tools("ID_JPQuarterly").State = ssChecked Then oBtn.BtnState = eBtnState_Selected
                Case "ID_JPExpiration"
                    If Chart.tbToolbar.Tools("ID_JPExpiration").State = ssChecked Then oBtn.BtnState = eBtnState_Selected
                'dinapoli buttons
                Case "ID_OscPredictor"
                    If Chart.tbToolbar.Tools("ID_OscPredictor").State = ssChecked Then oBtn.BtnState = eBtnState_Selected
                Case "ID_MacdPredictor"
                    If Chart.tbToolbar.Tools("ID_MacdPredictor").State = ssChecked Then oBtn.BtnState = eBtnState_Selected
                Case "ID_DisplacedMA"
                    If Chart.tbToolbar.Tools("ID_DisplacedMA").State = ssChecked Then oBtn.BtnState = eBtnState_Selected
                Case "ID_DiNapoliMACD"
                    If Chart.tbToolbar.Tools("ID_DiNapoliMACD").State = ssChecked Then oBtn.BtnState = eBtnState_Selected
                Case "ID_PrefStoch"
                    If Chart.tbToolbar.Tools("ID_PrefStoch").State = ssChecked Then oBtn.BtnState = eBtnState_Selected
                Case "ID_DetrendOsc"
                    If Chart.tbToolbar.Tools("ID_DetrendOsc").State = ssChecked Then oBtn.BtnState = eBtnState_Selected
                Case "ID_Sectors", "ID_Subsectors", "ID_Components"
                    If Not Chart Is Nothing Then
                        Chart.GetSectorInfo strSector, strSub, strCmp, strSecDesc, strSubDesc, strCmpDesc
                        If oBtn.BtnID = "ID_Sectors" Then
                            oBtn.BtnCaption = strSecDesc
                        ElseIf oBtn.BtnID = "ID_Subsectors" Then
                            oBtn.BtnCaption = strSubDesc
                        Else
                            oBtn.BtnCaption = strCmp
                        End If
                    End If
            End Select
        End If
    Next

    ToolbarResize2 Me, Me.pbTbBack, Me.imgTbBack, m.aButtons, True, Me.Left, Me.Top

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTbMoreButtons.InitForm2"
    Unload Me

End Sub

Private Sub cboBarPeriod_Click()
On Error GoTo ErrSection:

    Dim frm As frmChart

    If cboBarPeriod.Visible Then            '4939
        If Not m.frm Is Nothing Then
            If IsFrmChart(m.frm) Then
                Set frm = m.frm
            ElseIf TypeOf m.frm Is frmMain Then
                Set frm = g.ChartGlobals.frmActiveNonDetached
            End If

            If Not frm Is Nothing Then
                frm.Chart.ChangeBarPeriod cboBarPeriod.Text
            End If
        End If
        
        Me.Hide
        tmr.Interval = 500
        tmr.Enabled = True      'cannot unload form here so must do it in timer event
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTbMoreButtons.cboBarPeriod_Click"

End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    g.Styler.StyleForm Me
    
    Set imgTbBack(0).Picture = frmMain.imgTbBack(0).Picture

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmTbMoreButtons.Form_Load"
    Unload Me

End Sub

Public Function ToolBarWrapGet(ByVal strToolbar$) As Boolean
On Error GoTo ErrSection:

    If Not m.frm Is Nothing Then
        m.frm.ToolBarWrapGet strToolbar
    End If

ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmTbMoreButtons.ToolBarWrapGet"

End Function

Public Property Get FormSource() As Form
    Set FormSource = m.frm
End Property

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Set m.aButtons = Nothing
    Set m.oBtnMouseLast = Nothing
    Set m.frm = Nothing

End Sub

Private Sub imgTbBack_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    Dim i&
    Dim oButton As cPicBoxButton

    Set m.oBtnMouseLast = ToolbarMouseEvent(Me, m.oBtnMouseLast, WM_LBUTTONDOWN, Button, Index, X, Y, False)
    
    If Not m.oBtnMouseLast Is Nothing Then
        If m.oBtnMouseLast.ToolBarName = kTbDraw Then
            If IsFrmChart(m.frm) Then
                m.frm.DrawToolSelect Me
            ElseIf Not g.ChartGlobals.frmActiveNonDetached Is Nothing Then
                g.ChartGlobals.frmActiveNonDetached.DrawToolSelect Me
            End If
        End If
    End If

End Sub

Private Sub imgTbBack_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    Dim oButton As cPicBoxButton

    Set oButton = ToolbarMouseEvent(Me, m.oBtnMouseLast, WM_MOUSEMOVE, Button, Index, X, Y, False)
    If Not oButton Is m.oBtnMouseLast Then
        ClearLastMouseButton Me, m.oBtnMouseLast
        Set m.oBtnMouseLast = oButton
        If Not m.oBtnMouseLast Is Nothing Then
            m.oBtnMouseLast.BtnToolTipShow Me, pbTbBack(Index)
        End If
    End If
    
End Sub

Private Sub imgTbBack_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

    m.oBtnMouseLast.MouseUp Me, pbTbBack(Index), m.aButtons, Nothing
    Me.Hide
    tmr.Interval = 1000
    tmr.Enabled = True

End Sub

Private Sub tmr_Timer()
On Error Resume Next

    If Not Me.Visible Then
        Unload Me
    End If

End Sub

Public Property Get TbButtonsArray(ByVal strToolbar$) As cGdArray
On Error GoTo ErrSection:

    Set TbButtonsArray = m.aButtons

ErrExit:
    Exit Property

ErrSection:
    RaiseError "frmTbMoreButtons.TbButtonsArray"

End Property

