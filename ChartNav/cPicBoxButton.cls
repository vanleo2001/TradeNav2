VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPicBoxButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cPicButton.class
'' Description: Implementation for custom tool bar buttons done with picture boxes
''
'' Modified:    Genesis Financial Technologies
''              4775 Centennial Blvd, Suite 150
''              Colorado Springs, 80919
''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author          Description
'' 06-24-2008   J. Muchmore     class creation
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    frm As Form
    strID As String
    strIconID As String
    
    strToolbar As String
    nToolbarPos As Long
    
    strCaption As String
    strTip As String
    
    eStyle As eBtnStyle
    eGroup As eBtnGroup
    eCategory As eBtnCategory
    
    nTextOnly As Long
    nTextColor As Long
    nLeft As Long                       'left in pixels
    nTop As Long                        'top in pixels
    nWidth As Long                      'width in pixels
    nHeight As Long                     'height in pixels
    nIndex As Long                      'index into gdArray that holds button object
    nDrawIndex As Long                  'index of draw position of button
    nPicboxIndex As Long                'index of the picturebox control that button is drawn on
    
    strStatus As String
    nLastVisibleBtnIdx As Long          'index into gdArray of last button visible on screen
    
    oToolTip As cToolTip
End Type

Private m As mPrivate

Private Sub Class_Initialize()
On Error Resume Next

    PropertiesClear

End Sub

Private Sub Class_Terminate()
On Error Resume Next
    
    PropertiesClear
    Set m.oToolTip = Nothing

End Sub

Private Sub PropertiesClear()
On Error GoTo ErrSection:

    m.strID = ""
    m.strIconID = ""
    
    m.strToolbar = ""
    m.nToolbarPos = vbAlignNone
    
    m.eStyle = eBtnStyle_Unknown
    m.eGroup = eBtnGroup_Unknown
    m.eCategory = eBtnCat_Unknown
    
    m.nTextOnly = 0
    m.nLeft = -1
    m.nTop = -1
    m.nWidth = -1
    m.nHeight = -1
    m.nIndex = -1
    m.nDrawIndex = -1
    m.nPicboxIndex = -1
    
    m.strCaption = ""
    m.strTip = ""
    
    m.strStatus = kKeyMouseNone
    m.nLastVisibleBtnIdx = -1
    
    Set m.frm = Nothing

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cPicBoxButton.PropertiesClear"

End Sub

Private Sub BtnStatusSync()
On Error Resume Next:
    
    m.strStatus = kKeyMouseNone
    
    Select Case m.strID
        Case "ID_ChartMove"
            If g.ChartGlobals.eChartMode = eMode_Move Then m.strStatus = kBtnStateDown
        Case "ID_ZoomIn"
            If g.ChartGlobals.eChartMode = eMode_Zoom Then m.strStatus = kBtnStateDown
        Case "ID_ZoomOut"
            m.strStatus = kBtnStateDisabled
        Case "ID_Eraser"
            If g.ChartGlobals.eChartMode = eMode_Erase Then m.strStatus = kBtnStateDown
        Case "ID_Magnet", "ID_RepeatDraw", "ID_DragModeY"
            If frmMain.tbToolbar.Tools(m.strID).State = ssChecked Then m.strStatus = kBtnStateDown
    End Select

End Sub


Public Function PropertiesSet(frmSource As Form, tblConfig As cGdTable, aIndex As cGdArray, _
    ByVal strToolID$, ByVal strToolbar$, ByVal vbeAlign&, ByVal idx&) As Boolean
On Error Resume Next:
    
    Dim i&, j&, X&, Y&
    
    PropertiesClear
    
    If Not tblConfig Is Nothing And Not aIndex Is Nothing Then
        frmMain.ToolBarBtnSizeGet strToolbar, X, Y
        
        If tblConfig.SearchAsIndex(aIndex, eBtnCfgField_ID, strToolID, i) Then
            j = aIndex(i)
            
            Set m.frm = frmSource
            
            m.nWidth = X
            m.nHeight = Y
            m.nIndex = idx
            
            m.strID = strToolID
            If strToolID = "ID_ZoomOut" Then
                m.strStatus = kBtnStateDisabled
            ElseIf strToolID = "ID_TextIncrease" Or strToolID = "ID_TextDecrease" Then
                If TypeOf frmSource Is frmMain Then
                    m.strStatus = kBtnStateDisabled
                Else
                    m.strStatus = kKeyMouseNone     '+T/-T on detached chart's toolbar should never be disabled
                End If
            Else
                m.strStatus = kKeyMouseNone
            End If
            
            m.strToolbar = strToolbar
            m.nToolbarPos = vbeAlign
            
            If strToolID = "ID_Chain" Then
'JM (11-04-2009): original code leave awhile then remove if all okay
'                If HasModule("OPTNAV") And Len(OptNavIP) > 0 And FileExist(OptNavExeFile) Then
'                    m.eStyle = eBtnStyle_Push
'                Else
                    m.eStyle = tblConfig(eBtnCfgField_eStyle, j)
'                End If
            Else
                m.eStyle = tblConfig(eBtnCfgField_eStyle, j)
            End If
            
            m.eGroup = tblConfig(eBtnCfgField_eGroup, j)
            m.eCategory = tblConfig(eBtnCfgField_eCategory, j)
            
            m.nTextOnly = Val(tblConfig(eBtnCfgField_TextOnly, j))
            
            m.strCaption = tblConfig(eBtnCfgField_Name, j)
            m.strTip = tblConfig(eBtnCfgField_Tooltip, j)
            If Len(m.strTip) = 0 Then m.strTip = m.strCaption
            
            If m.strID = "ID_FibTimeZones" And HasModule("JGREEN") Then
                m.strCaption = "Greenblatt"
                m.strTip = "Greenblatt Time Zone"       '5360
            ElseIf m.strID = "ID_DNExpansion" And Not UseDiNapFib() Then        'Not HasModule("FIB") Then
                m.strCaption = "Fib Ext"
                m.strTip = "Fibonacci Extension"
            ElseIf m.strID = "ID_AndrewFork" Then
                If HasModule("ADVFIB") Then
                    m.strTip = "Andrews/Schiff Pitchfork"
                Else
                    m.strTip = "Andrews Pitchfork"
                End If
            End If
                        
            BtnStatusSync
            If m.nTextOnly = 1 Or m.strID = "ID_TradeTracker" Then m.nWidth = BtnBestWidth
                        
            PropertiesSet = True
        End If
    End If

End Function

Public Property Get BtnID() As String
    BtnID = m.strID
End Property

Public Property Get BtnCaption() As String
    BtnCaption = m.strCaption
End Property

Public Property Let BtnCaption(ByVal strText$)
On Error Resume Next:
    
    Dim iWidthSave&
        
    If strText <> m.strCaption Then
        m.strCaption = strText
        If m.nTextOnly = 1 Then
            iWidthSave = m.nWidth
            m.nWidth = BtnBestWidth
            If m.nWidth <> iWidthSave Then
                If Not m.frm Is Nothing Then
                    If IsFrmChartMDI(m.frm) Then
                        ToolbarResize2 frmMain, frmMain.pbTbBack, frmMain.imgTbBack, frmMain.TbButtonsArray(kTbGeneral), frmMain.ToolBarWrapGet(kTbGeneral)
                    Else
                        ToolbarResize2 m.frm, m.frm.pbTbBack, m.frm.imgTbBack, m.frm.TbButtonsArray(kTbGeneral), m.frm.ToolBarWrapGet(kTbGeneral)
                    End If
                End If
            End If
        End If
    End If

End Property

Public Property Get BtnToolTip() As String
    BtnToolTip = m.strTip
End Property

Public Property Let BtnToolTip(ByVal strTip$)
    m.strTip = strTip
End Property

Public Property Get BtnStyle() As eBtnStyle
    BtnStyle = m.eStyle
End Property

Public Property Get BtnCategory() As eBtnCategory
    BtnCategory = m.eCategory
End Property

Public Property Get BtnGroup() As eBtnGroup
    BtnGroup = m.eGroup
End Property

Public Property Get BtnIconID() As String
    BtnIconID = m.strIconID
End Property

Public Property Get BtnStatus() As String
    BtnStatus = m.strStatus
End Property

Public Property Let BtnStatus(ByVal kStatus$)
    m.strStatus = kStatus
End Property

Public Property Get BtnState() As eBtnState
    
    If m.strStatus = kBtnStateDown Then
        BtnState = eBtnState_Selected
    ElseIf m.strStatus = kKeyMouseInProg Then
        BtnState = eBtnState_InProg
    ElseIf m.strStatus = kBtnStateDisabled Then
        BtnState = eBtnState_Disabled
    Else
        BtnState = eBtnState_Neutral
    End If

End Property

Public Property Let BtnState(ByVal eState As eBtnState)
    
    If eState = eBtnState_InProg Then
        m.strStatus = kKeyMouseInProg
    ElseIf eState = eBtnState_Selected Then
        m.strStatus = kBtnStateDown
    ElseIf eState = eBtnState_Disabled Then
        m.strStatus = kBtnStateDisabled
    Else
        m.strStatus = kKeyMouseNone
    End If
    
End Property

Public Property Get BtnWidth() As Long
    'BtnWidth = m.nWidth     'pixel
    BtnWidth = BtnBestWidth()
End Property

Public Property Get BtnHeight() As Long
    BtnHeight = m.nHeight
End Property

Public Property Get ToolBarName() As String
    ToolBarName = m.strToolbar
End Property

Public Property Get ToolBarPos() As Long
    ToolBarPos = m.nToolbarPos
End Property

Public Property Get BtnIsTextOnly() As Long
    BtnIsTextOnly = m.nTextOnly
End Property

Public Property Get LastVisibleBtnIndex() As Long
    LastVisibleBtnIndex = m.nLastVisibleBtnIdx
End Property

Public Property Let LastVisibleBtnIndex(ByVal idx&)
    m.nLastVisibleBtnIdx = idx
End Property

Public Property Get PicboxIndex() As Long
    PicboxIndex = m.nPicboxIndex
End Property

Public Property Get BtnIndex() As Long
    BtnIndex = m.nIndex
End Property

Public Property Get BtnDrawIndex() As Long
    BtnDrawIndex = m.nDrawIndex
End Property

Public Sub BtnDrawIndexReset()
    m.nDrawIndex = -1
End Sub

Public Sub Draw(pB As PictureBox, aButtons As cGdArray, _
    Optional ByVal iPicboxIdx& = -1, Optional ByVal strIcon$ = "", _
    Optional ByRef oBtnPrevious As cPicBoxButton = Nothing)
On Error Resume Next:

    Dim imgBk As ListImage
    Dim imgIcon As ListImage
    Dim oBtnPrev As cPicBoxButton
    
    Dim ptrBtnPrev As picbox_struct
    Dim ptrBtn As picbox_struct
    
    Dim i&, idx&, strText$
    Dim X&, Y&
    
    Dim bLocatePrevBtn As Boolean
    
    m.nTextColor = -1
    
    If pB Is Nothing Or aButtons Is Nothing Then Exit Sub
    If m.nLastVisibleBtnIdx = -1 Then Exit Sub          'button was never drawn, not visible
    If m.strID <> kMoreBtnID And m.nIndex > m.nLastVisibleBtnIdx Then Exit Sub
       
    If iPicboxIdx >= 0 Then m.nPicboxIndex = iPicboxIdx
        
    If g.vbeTbAlignDraw = vbAlignBottom Then
        If TypeOf pB.Parent Is frmTbMoreButtons Then
            'do nothing
        ElseIf m.strToolbar = kTbDraw And m.nIndex = 0 Then
            bLocatePrevBtn = True
        End If
    End If
    
    Set imgBk = ToolBtnBkground()
    Set imgIcon = ToolBtnIcon()
    
    If m.nWidth >= 59 Then
        strText = m.strCaption
        If strText = "Component" Then strText = " Component"
    Else
        If m.nTextOnly = 1 Then
            strText = m.strCaption
        Else
            strText = ""
        End If
    End If
    
    If m.strID = "ID_TradeTracker" Then
        'If m.nWidth <> kBtnSmallIcoTextWd Then strText = ""
        If g.nTbIncludeText = 0 Then strText = ""
        If DockState(frmTTSummary) = eHidden Then
            strText = strText & "~D"
        Else
            strText = strText & "~U"
        End If
    End If
    
    If m.strID = kMoreBtnID Then
        idx = m.nLastVisibleBtnIdx + 1
        Set oBtnPrev = aButtons(m.nLastVisibleBtnIdx)
    ElseIf Not oBtnPrevious Is Nothing Then
        Set oBtnPrev = oBtnPrevious
        If oBtnPrev.ToolBarName = m.strToolbar Then idx = oBtnPrev.BtnDrawIndex + 1
    ElseIf bLocatePrevBtn Then
        Dim aTemp As cGdArray
        Dim frm As Form
        
        Set frm = pB.Parent
        If Not frm Is Nothing Then
            If frm.pbTbBack(0).Visible Then
                Set aTemp = frm.TbButtonsArray("")
                If Not aTemp Is Nothing Then
                    Set oBtnPrev = aTemp(aTemp.Size - 1)
                    If oBtnPrev.BtnDrawIndex = -1 Then
                        Set oBtnPrev = aTemp(0)
                        If Not oBtnPrev Is Nothing Then
                            i = oBtnPrev.LastVisibleBtnIndex
                            If i >= 0 And i < aTemp.Size Then       'precautionary; theoretically should always be valid
                                Set oBtnPrev = aTemp(i)
                            Else
                                Set oBtnPrev = Nothing
                            End If
                        End If
                    End If
                End If
            End If
        End If
    ElseIf m.nPicboxIndex <= 0 Then
        idx = m.nIndex
        If m.nIndex > 0 Then Set oBtnPrev = aButtons(m.nIndex - 1)
    Else
        idx = -1
        For i = 0 To aButtons.Size - 1
            If aButtons(i).PicboxIndex = m.nPicboxIndex Then
                idx = idx + 1
                If i = m.nIndex Then
                    Exit For
                Else
                    Set oBtnPrev = aButtons(i)
                End If
            End If
        Next
    End If
    
    
    m.nDrawIndex = idx
        
    If oBtnPrev Is Nothing Then
        ptrBtnPrev.iLeft = 0
        ptrBtnPrev.iTop = 0
        ptrBtnPrev.iconHeight = 0
        ptrBtnPrev.iconWidth = 0
        ptrBtnPrev.pbWidth = 0
        ptrBtnPrev.pbHeight = 0
    Else
        oBtnPrev.BtnStructInfo ptrBtnPrev.iLeft, ptrBtnPrev.iTop, ptrBtnPrev.pbWidth, ptrBtnPrev.pbHeight, ptrBtnPrev.iconWidth
        ptrBtnPrev.iconWidth = ptrBtnPrev.iconHeight
    End If
    
    BtnStructInfo ptrBtn.iLeft, ptrBtn.iTop, ptrBtn.pbWidth, ptrBtn.pbHeight, ptrBtn.iconWidth
    ptrBtn.iconHeight = ptrBtn.iconWidth
    
    If m.strID = "ID_BarPeriod" Then
        m.nLeft = ptrBtnPrev.iLeft + ptrBtnPrev.pbWidth
        m.nTop = ptrBtnPrev.iTop
        
        frmMain.ToolBarBtnSizeGet kTbGeneral, X, Y
        
        If Y > kBtnSmallIco Then
            m.nHeight = Y
            m.nTop = m.nTop + m.nHeight
            m.nWidth = X * 2 + 1
            
            If m.nHeight = kBtnLargeIcoTextHt Then
                m.nTop = m.nTop - (m.frm.cboBarPeriod.Height * 1.3) / Screen.TwipsPerPixelY
            Else
                m.nTop = m.nTop - m.frm.cboBarPeriod.Height / Screen.TwipsPerPixelY
            End If
            
            ptrBtn.pbHeight = m.nHeight
            ptrBtn.pbWidth = m.nWidth
            
            If g.eTbSkin = eTbSkin_DarkFlat Then
                geOverlayICOAll pB.hDC, _
                                imgBk.ExtractIcon, _
                                0, _
                                m.nToolbarPos, _
                                RGB(238, 238, 238), _
                                strText, _
                                ptrBtnPrev, _
                                ptrBtn
            Else
                geOverlayICOAll pB.hDC, _
                                imgBk.ExtractIcon, _
                                0, _
                                m.nToolbarPos, _
                                vbBlack, _
                                strText, _
                                ptrBtnPrev, _
                                ptrBtn
            End If
            
            m.nLeft = ptrBtn.iLeft
            m.frm.cboBarPeriod.Width = (m.nWidth - 6) * Screen.TwipsPerPixelX
        Else
            m.nTop = m.nTop + 2 ' TLB: add 2 pixels to center vertically
            m.frm.cboBarPeriod.Width = kBtnSmallIco * Screen.TwipsPerPixelX * 4
            m.nWidth = kBtnSmallIco * 4 + 4
            m.nHeight = Y
        End If
        m.frm.cboBarPeriod.Move (m.nLeft + 3) * Screen.TwipsPerPixelX, m.nTop * Screen.TwipsPerPixelY
        m.frm.cboBarPeriod.ZOrder
    Else
        Dim textColor As Long
        
        If m.nTextColor = -1 Then
            textColor = RGB(238, 238, 238)
        Else
            textColor = m.nTextColor
        End If
        
        If imgIcon Is Nothing Then
            If g.eTbSkin = eTbSkin_DarkFlat Or g.eTbSkin = eTbSkin_LightFlat Then
                geOverlayICOAll pB.hDC, _
                                imgBk.ExtractIcon, _
                                0, _
                                m.nToolbarPos, _
                                textColor, _
                                strText, _
                                ptrBtnPrev, _
                                ptrBtn
            Else
                geOverlayICOAll pB.hDC, _
                                imgBk.ExtractIcon, _
                                0, _
                                m.nToolbarPos, _
                                vbBlack, _
                                strText, _
                                ptrBtnPrev, _
                                ptrBtn
            End If
        ElseIf g.eTbSkin = eTbSkin_DarkFlat Or g.eTbSkin = eTbSkin_LightFlat Then
            geOverlayICOAll pB.hDC, _
                            imgBk.ExtractIcon, _
                            imgIcon.ExtractIcon, _
                            m.nToolbarPos, _
                            textColor, _
                            strText, _
                            ptrBtnPrev, _
                            ptrBtn
        Else
            geOverlayICOAll pB.hDC, _
                            imgBk.ExtractIcon, _
                            imgIcon.ExtractIcon, _
                            m.nToolbarPos, _
                            vbBlack, _
                            strText, _
                            ptrBtnPrev, _
                            ptrBtn
        End If
        m.nLeft = ptrBtn.iLeft
        m.nTop = ptrBtn.iTop
    End If
    
    If m.strStatus = kBtnStateDown Then
        BtnGroupClear pB, aButtons
    ElseIf m.strID = "ID_RepeatDraw" Then
        If Len(g.strActiveDraw) > 0 Then
            If m.strStatus = kKeyMouseNone Then
                If Not ActiveChart Is Nothing Then ActiveChart.SyncDrawTools        '5665
            End If
        End If
    End If

    Set imgIcon = Nothing
            
End Sub

Public Function HitTest(ByVal iPicboxIdx&, ByVal X&, ByVal Y&) As Boolean
On Error Resume Next:

    Dim iPix&

    If m.nPicboxIndex = iPicboxIdx Then
        If m.nDrawIndex <> -1 Then
            If m.nToolbarPos = vbAlignTop Or m.nToolbarPos = vbAlignBottom Then
                iPix = X / Screen.TwipsPerPixelX
                If m.nLeft < iPix And m.nLeft + m.nWidth > iPix Then HitTest = True
            Else
                iPix = Y / Screen.TwipsPerPixelY
                If m.nTop < iPix And m.nTop + m.nHeight > iPix Then HitTest = True
            End If
        End If
    End If

End Function

Public Function CursorCheckClear(frm As Form, aButtons As cGdArray) As Boolean
On Error Resume Next:

    Dim i&, X&, Y&
    Dim pB As PictureBox
    Dim bClear As Boolean
        
    If m.strStatus <> kKeyMouseMove And m.strID <> "ID_BarPeriod" Then Exit Function
    
    i = m.nPicboxIndex
    
    If m.strToolbar = kTbDraw Then
        If g.vbeTbAlignDraw = vbAlignBottom And frm.pbTbBack(0).Visible Then
            If i >= 0 And i <= frm.pbTbBack.UBound Then Set pB = frm.pbTbBack(i)
        ElseIf i >= 0 And i <= frm.pbTbBackDraw.UBound Then
            Set pB = frm.pbTbBackDraw(i)
        End If
    ElseIf i >= 0 And i <= frm.pbTbBack.UBound Then
        Set pB = frm.pbTbBack(i)
    End If
        
    If Not pB Is Nothing Then
        If geIsCursorInWnd(pB.hWnd, X, Y) = 1 Then
            If m.nToolbarPos = vbAlignTop Or m.nToolbarPos = vbAlignBottom Then
                If X >= m.nLeft And X <= m.nLeft + m.nWidth Then
                    bClear = False
                Else
                    bClear = True
                End If
            ElseIf Y >= m.nTop And Y <= m.nTop + m.nHeight Then
                bClear = False
            Else
                bClear = True
            End If
        Else
            bClear = True
        End If
        
        If m.strID = "ID_BarPeriod" And bClear Then
            If frm.ActiveControl Is frm.cboBarPeriod Then
                If X >= m.nLeft And X <= m.nLeft + m.nWidth Then
                    If Y >= m.nTop And Y < m.nTop + m.nHeight + 120 Then
                        bClear = False      'cursor may be in dropdown portion, don't clear
                    End If
                End If
            End If
        End If
        
        If bClear Then
            If IsFrmChart(frm) And m.strID = "ID_BarPeriod" Then MoveFocus frm.pbChart
            BtnClearNow pB, aButtons
        End If
        
    End If
    
    CursorCheckClear = bClear
    
End Function

Public Function MouseMove(frm As Form, pB As PictureBox, aButtons As cGdArray, _
    oBtnMore As cPicBoxButton) As cPicBoxButton
On Error Resume Next:

    Dim oButton As cPicBoxButton
    
    If pB Is Nothing Then Exit Function
    If aButtons Is Nothing Then Exit Function
    If oBtnMore Is Nothing Then Exit Function
            
    If FormIsLoaded("frmTemplatePage") Then Unload frmTemplatePage
    
    If Not TypeOf frm Is frmTbMoreButtons Then
        If FormIsLoaded("frmTbMoreButtons") Then
            If Not m.strID = kMoreBtnID Then
                If m.nHeight > 22 Then
                    If oBtnMore.BtnDrawIndex - m.nDrawIndex >= 2 Then
                        Unload frmTbMoreButtons
                    End If
                ElseIf oBtnMore.ToolBarPos = vbAlignTop Then
                    If oBtnMore.BtnDrawIndex - m.nDrawIndex >= 4 Then
                        'don't unload form unless user is 4 buttons away
                        Unload frmTbMoreButtons
                    End If
                Else
                    If oBtnMore.BtnDrawIndex - m.nDrawIndex >= 6 Then
                        'don't unload form unless user is 4 buttons away
                        Unload frmTbMoreButtons
                    End If
                End If
            End If
        End If
    End If
    
    If m.nDrawIndex = -1 Then
        'do nothing
    ElseIf oBtnMore.BtnDrawIndex > 0 And m.nDrawIndex > oBtnMore.BtnDrawIndex Then
        'do nothing
    ElseIf m.nDrawIndex = oBtnMore.BtnDrawIndex Then
        Set oButton = oBtnMore
    Else
        Set oButton = Me
    End If
        
    If Not oButton Is Nothing Then
        If m.strStatus = kKeyMouseNone Then
            oButton.BtnStatus = kKeyMouseMove
            oButton.Draw pB, aButtons
            pB.Refresh
        End If
    End If
        
    Set MouseMove = oButton

End Function

Public Function MouseDown(frm As Form, pB As PictureBox, aButtons As cGdArray, _
    oBtnMore As cPicBoxButton) As cPicBoxButton
On Error Resume Next:

    Dim oButton As cPicBoxButton
    Dim i&, iTop&

    If pB Is Nothing Then Exit Function
    If aButtons Is Nothing Then Exit Function

'JM 02-23-2011 - original code, leave awhile then remove
'    If g.strActiveDraw = "ID_Pattern" Then
'        If m.strID = "ID_ChartMove" Or m.strID = "ID_ZoomIn" Or m.strID = "ID_Eraser" Then
'            Exit Function       '5332, 6000
'        End If
'    End If
    
    'JM 02-23-2011 - fix for 5332, 6000, 6183
    If m.strID = "ID_ChartMove" Or m.strID = "ID_ZoomIn" Or m.strID = "ID_Eraser" Then
        If g.strActiveDraw = "ID_Pattern" Or BtnState = eBtnState_Selected Then Exit Function
    End If
        
       
    If m.nDrawIndex = oBtnMore.BtnDrawIndex Then
        Set oButton = oBtnMore
        If FormIsLoaded("frmTbMoreButtons") Then
            Unload frmTbMoreButtons
        Else
            If oButton.ToolBarName = kTbDraw Then
                If oButton.ToolBarPos = vbAlignTop Then
                    For i = frm.pbTbBack.UBound To 0 Step -1
                        If frm.pbTbBack(i).Visible Then
                            iTop = frm.pbTbBack(i).Top + frmMain.pbTbBack(i).Height
                            Exit For
                        End If
                    Next
                End If
            End If
            frmTbMoreButtons.ShowMe frm, oButton.ToolBarName, oButton.BtnLeft + oButton.BtnWidth, oButton.BtnTop + oButton.BtnHeight + iTop / Screen.TwipsPerPixelY
            oButton.BtnClearNow pB, aButtons
        End If
        Exit Function
    Else
        Set oButton = Me
        If FormIsLoaded("frmTbMoreButtons") And Not TypeOf frm Is frmTbMoreButtons Then
            Unload frmTbMoreButtons
        End If
    End If
    
    If oButton Is Nothing Then Exit Function
        
    If IsDropdownBtn() Then
        If TypeOf frm Is frmTbMoreButtons Then
            'form is getting unloaded, no need to redraw or anything
            HandleDropDownBtns
        Else
            If m.strStatus = kKeyMouseNone Or m.strStatus = kKeyMouseMove Then
                m.strStatus = kKeyMouseDown
                Draw pB, aButtons
                HandleDropDownBtns
            Else
                BtnClearNow pB, aButtons
            End If
            pB.Refresh
        End If
    ElseIf m.strStatus = kKeyMouseNone Or m.strStatus = kKeyMouseMove Then
    
        If TypeOf frm Is frmMain Then       '5242
            'apply action of charting buttons to active chart that is either not detached or detached but does not have its own toolbar
            If oButton.BtnCategory = eBtnCat_Charting Then
                If Not ActiveChart Is Nothing Then
                    If ActiveChart.DetachStatus = eDetached Then
                        If ActiveChart.Chart.ShowToolbar Then
                            If Not g.ChartGlobals.frmActiveNonDetached Is Nothing Then
                                ActiveChartFormSet g.ChartGlobals.frmActiveNonDetached
                            End If
                        End If
                    End If
                End If
            End If
        End If
        
        If m.eStyle = eBtnStyle_State Then
            If m.strID = "ID_Chain" Then
                If frmMain.tbToolbar.Tools(m.strID).State = ssChecked Then
                    frmMain.tbToolbar.Tools(m.strID).State = ssUnchecked        '6344 - this is timing issue (traffic light is yellow)
                    Exit Function
                End If
            End If
        
            frmMain.tbToolbar.Tools(m.strID).State = ssChecked
            If oButton.BtnID = "ID_RealTime" Then
                'let realtime object update button state
            ElseIf Not TypeOf frm Is frmTbMoreButtons Then
                oButton.BtnStatus = kBtnStateDown
                oButton.Draw pB, aButtons
                pB.Refresh
            End If
        ElseIf m.eStyle = eBtnStyle_Push Then
            If m.strID = "ID_Tile" Then
                If Not g.ChartGlobals.frmActiveNonDetached Is Nothing Then
                    SendMessage g.ChartGlobals.frmActiveNonDetached.hWnd, WM_NCACTIVATE, 1, 0
                    SendMessage g.ChartGlobals.frmActiveNonDetached.hWnd, WM_MOUSEACTIVATE, 1, 0
                End If
            End If
            If TypeOf frm Is frmTbMoreButtons Then
                If m.strID = "ID_Symbol" Then
                    If Not ActiveChart Is Nothing Then
                        ActiveChart.ChangeSymWrapper
                    End If
                ElseIf m.strID = "ID_ShowEWI" Then
                    'do nothing - only process this button on mouseup
                Else
                    ToolBarClick frmMain.tbToolbar.Tools(m.strID), frmMain
                End If
            Else
                m.strStatus = kKeyMouseDown
                Draw pB, aButtons
                pB.Refresh
            End If
        End If
    ElseIf m.strStatus = kKeyMouseDown Or m.strStatus = kBtnStateDown Then
        If oButton.BtnID = "ID_RealTime" Then
        Else
            BtnClearNow pB, aButtons
        End If
        
        If m.eStyle = eBtnStyle_State Then
            frmMain.tbToolbar.Tools(m.strID).State = ssUnchecked
            If m.strID = "ID_ElliotLabels" Then
                Unload frmElliot
            ElseIf m.strID = "ID_Icon" Then
                Unload frmIconAnnot
            End If
        End If
    ElseIf m.strStatus = kKeyMouseInProg Then
        If oButton.BtnID = "ID_RealTime" Then
            If ProcessIsBusy(True) Then
                If frmStatus.Status = eStatus_Aborting Or frmStatus.Status = eStatus_Aborted Then
                    'do nothing - not done aborting
                    'the status dialog will say "Aborted", but is is not done, it needs to reload data etc.
                Else
                    frmStatus.Status = eStatus_Aborting
                    g.RealTime.RefreshSymbolList True
                End If
            Else
                frmMain.tbToolbar.Tools(m.strID).State = ssUnchecked
            End If
        End If
    End If
    
    If m.strID = "ID_ChartMove" Then
        g.ChartGlobals.eChartMode = eMode_Move
    ElseIf m.strID = "ID_ZoomIn" Then
        g.ChartGlobals.eChartMode = eMode_Zoom
    ElseIf m.strID = "ID_Eraser" Then
        g.ChartGlobals.eChartMode = eMode_Erase
    ElseIf m.strID = "ID_Chain" Then
        If HasModule("OPTNAV") And Len(OptNavIP) > 0 And FileExist(OptNavExeFile) Then
            BtnClearNow pB, aButtons
        End If
    End If
    
    BtnSyncAll
    
    Set MouseDown = oButton
    
End Function

Public Function MouseUp(frm As Form, pB As PictureBox, aButtons As cGdArray, _
    oBtnMore As cPicBoxButton, Optional ByVal bRightClick As Boolean = False) As cPicBoxButton
On Error Resume Next:

    Dim oButton As cPicBoxButton

    If pB Is Nothing Then Exit Function
    If aButtons Is Nothing Then Exit Function
    
    If TypeOf frm Is frmTbMoreButtons Then
        If Me.BtnID = "ID_ShowEWI" Then
            ToolBarClick frmMain.tbToolbar.Tools(m.strID), frmMain
        End If
        Exit Function
    End If
       
    If m.nDrawIndex = oBtnMore.BtnDrawIndex Then
        Set oButton = oBtnMore
    Else
        Set oButton = Me
    End If
    
    If oButton Is Nothing Then Exit Function
    
    If oButton.BtnID = "ID_ZoomOut" Then
        If m.strStatus = kKeyMouseDown Or m.strStatus = kBtnStateDown Then
            frmMain.tbToolbar.Tools(m.strID).State = ssUnchecked
            oButton.BtnClearNow pB, aButtons
            pB.Refresh
        End If
    ElseIf oButton.BtnStyle = eBtnStyle_Push Then
'JM 08-19-2010: development code; leave awhile then remove if all okay
'        If m.strID = "ID_PatternProfit" And bRightClick Then
'            ToolBarClick frmMain.tbToolbar.Tools(m.strID), frmMain, , eTbExtraInfo_PFPOldForm
'        Else
'            ToolBarClick frmMain.tbToolbar.Tools(m.strID), frmMain
'        End If
        
        ToolBarClick frmMain.tbToolbar.Tools(m.strID), frmMain
        oButton.BtnClearNow pB, aButtons
        pB.Refresh
    End If
    
    Set MouseUp = oButton

End Function


Public Sub BtnClearNow(pbIn As PictureBox, aButtons As cGdArray, Optional ChartIn As cChart = Nothing)
On Error Resume Next:

    Dim Chart As cChart
    Dim frm As Form
    Dim pB As PictureBox
    
    Set Chart = ChartIn

    If g.bStarting Or g.bUnloading Or g.bLoadingChartPage Then Exit Sub
    If pbIn Is Nothing Then Exit Sub
    
    Set frm = pbIn.Parent
    
    If frm Is Nothing Then
        Set pB = pbIn
    ElseIf TypeOf frm Is frmMain Or IsFrmChart(frm) Then
        If m.strToolbar = kTbDraw Then
            Set pB = pbIn
        ElseIf m.nPicboxIndex >= 0 And m.nPicboxIndex <= frm.pbTbBack.UBound Then
            'when called from private sub BtnGroupClear need to find picbox group button is on
            Set pB = frm.pbTbBack(m.nPicboxIndex)
        Else
            Set pB = pbIn
        End If
    Else
        Set pB = pbIn
    End If

    If pB Is Nothing Then Exit Sub      'precautionary, theoretically should never happen
    If aButtons Is Nothing Then Exit Sub

    If m.strID = "ID_ZoomOut" Then
        If Chart Is Nothing Then
            If TypeOf m.frm Is frmMain Then
                If Not g.ChartGlobals.frmActiveNonDetached Is Nothing Then
                    Set Chart = g.ChartGlobals.frmActiveNonDetached.Chart
                End If
            ElseIf IsFrmChart(m.frm) Then
                Set Chart = m.frm.Chart
            End If
        End If
        If Chart Is Nothing Then
            m.strStatus = kBtnStateDisabled
        ElseIf Chart.Zoomed Then
            m.strStatus = kKeyMouseNone
        Else
            m.strStatus = kBtnStateDisabled
        End If
        Draw pB, aButtons
        pB.Refresh
    ElseIf m.strID = "ID_TextIncrease" Or m.strID = "ID_TextDecrease" Then
        If ActiveForm Is Nothing Or BtnState = eBtnState_Disabled Then
            m.strStatus = kBtnStateDisabled
        Else
            m.strStatus = kKeyMouseNone
        End If
        Draw pB, aButtons
        pB.Refresh
    Else
        m.strStatus = kKeyMouseNone
        Draw pB, aButtons
        pB.Refresh
        
        If m.strID = "ID_Templates" Or m.strID = "ID_Pages" Then
            If FormIsLoaded("frmTemplatePage") Then Unload frmTemplatePage
        End If
        
    End If

End Sub

Private Sub BtnGroupClear(pB As PictureBox, aButtons As cGdArray)
On Error Resume Next:

    Dim i&, j&
    Dim frm As Form
    Dim oButton As cPicBoxButton
    
    If m.eGroup = eBtnGroup_None Or m.eGroup = eBtnGroup_Unknown Then Exit Sub
    
    For i = 0 To aButtons.Size - 1
        Set oButton = aButtons(i)
        If Not oButton Is Me Then
            If oButton.BtnGroup = m.eGroup Then
            
                If m.strID = "ID_ZoomOut" And oButton.BtnID = "ID_ZoomIn" Then
                    'do nothing - user is in zoom mode & just unzoomed; leave zoom-in btn depressed
                ElseIf oButton.BtnStatus <> kKeyMouseNone And oButton.BtnStatus <> kBtnStateDisabled Then
                    If oButton.PicboxIndex = m.nPicboxIndex Then
                        oButton.BtnClearNow pB, aButtons
                    Else
                        Set frm = pB.Parent
                        If Not frm Is Nothing Then
                            j = oButton.PicboxIndex
                            If g.vbeTbAlignDraw = vbAlignBottom And frm.pbTbBack(0).Visible Then
                                If j >= 0 And j <= frm.pbTbBack.UBound Then
                                    oButton.BtnClearNow frm.pbTbBack(oButton.PicboxIndex), aButtons
                                End If
                            ElseIf j >= 0 And j <= frm.pbTbBackDraw.UBound Then
                                oButton.BtnClearNow frm.pbTbBackDraw(oButton.PicboxIndex), aButtons
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next

End Sub

Private Function BtnBestWidth() As Long
On Error Resume Next:

    Dim i&
    
'Public Const kBtnLargeIcoTextWd = 68    '78
'Public Const kBtnLargeIcoTextHt = 52
'Public Const kBtnSmallIcoTextWd = 59
'Public Const kBtnSmallIcoTextHt = 35
'Public Const kBtnSmallIco = 22          'buttons with no text have same width/height
'Public Const kBtnLargeIco = 36
    
    If m.strID = kMoreBtnID Then
        i = m.nWidth
    ElseIf m.nHeight = kBtnLargeIco Then
        If m.strID = "ID_Templates" Or m.strID = "ID_Components" Or m.strID = "ID_TradeTracker" Then
            i = kBtnLargeIco * 2 + 1         '75
        ElseIf m.strID = "ID_Subsectors" Or m.strID = "ID_Sectors" Then
            i = kBtnLargeIco * 3 + 2         '113
        Else
            i = kBtnLargeIco
        End If
    ElseIf m.nHeight = kBtnSmallIcoTextHt Then
        If m.strID = "ID_BarPeriod" Or m.strID = "ID_Sectors" Or m.strID = "ID_Subsectors" Then
            i = kBtnSmallIcoTextWd * 2 + 1          '119
        Else
            i = kBtnSmallIcoTextWd      '59
        End If
    ElseIf m.nHeight = kBtnLargeIcoTextHt Then
        If m.strID = "ID_BarPeriod" Or m.strID = "ID_Sectors" Or m.strID = "ID_Subsectors" Then
            i = kBtnLargeIcoTextWd * 2 + 1          '157
        Else
            i = kBtnLargeIcoTextWd      '78
        End If
    Else
        Select Case m.strID
            Case "ID_Pages", "ID_TradeTracker"
                i = kBtnSmallIco * 2 + 1      '45
            Case "ID_Sectors", "ID_Subsectors"
                i = kBtnSmallIco * 5 + 4      '114
            Case "ID_Templates", "ID_Components"
                i = kBtnSmallIco * 3 + 2      '68
            Case "ID_BarPeriod"
                i = kBtnSmallIco * 4 + 4      '92
            Case Else
                i = kBtnSmallIco  '22
        End Select
    End If
    
    BtnBestWidth = i

End Function

Private Function ToolBtnBkground() As ListImage
On Error Resume Next:

    Dim strImgListKey$, iWidth&
    Dim strTradeTrackerKey$             'green if there are no open position(s)/order(s) else orange
    Dim imgToUse As ListImage
    
    If m.strID = "ID_TradeTracker" Then
        ' 03/17/2010 DAJ: Use this call now to see if there are any visible working orders or
        ' open positions...
        If g.ConsoleForms.NumVisible(eGDConsoleForm_Summary) > 0 Then
            If g.eTbSkin = eTbSkin_ALuminumBlue Or g.eTbSkin = eTbSkin_AluminumSilver Then
                strTradeTrackerKey = "MouseOpenTradeAluminum"
            Else
                strTradeTrackerKey = "MouseOpenTradeCon"
            End If
        ElseIf g.eTbSkin = eTbSkin_ALuminumBlue Or g.eTbSkin = eTbSkin_AluminumSilver Then
            strTradeTrackerKey = "MouseNoneTradeAluminum"
        Else
            strTradeTrackerKey = "MouseNoneTradeCon"
        End If
    End If
    
    If m.strID = "ID_ShowEWI" Then
        If m.strStatus = kKeyMouseNone Or m.strStatus = kBtnStateDown Then
            
            Dim Chart As cChart
            
            If Not m.frm Is Nothing Then
                If IsFrmChart(m.frm) Then
                    Set Chart = m.frm.Chart
                ElseIf Not ActiveChart Is Nothing Then
                    Set Chart = ActiveChart.Chart
                End If
            End If
            
            If Not Chart Is Nothing Then
                If 0 = Chart.HasHiddenAnnots(eANNOT_ElliotLabel) Then
                    m.strStatus = kBtnStateDown
                Else
                    m.strStatus = kKeyMouseNone
                End If
            End If
            
        End If
    End If
    
    If g.eTbSkin = eTbSkin_Blue Then
        Select Case m.strStatus
            Case kKeyMouseNone, kBtnStateDisabled, kMouseJustUp
                strImgListKey = "MouseNoneBlue"
            Case kKeyMouseMove
                strImgListKey = "MouseMove"
            Case Else
                If m.strStatus = kKeyMouseInProg Then
                    If g.RealTime.ConnectionStatus = eGDConnectionStatus_Connecting Or _
                       g.RealTime.ConnectionStatus = eGDConnectionStatus_Connected Then
                        strImgListKey = "MouseNoneBlue"
                    Else
                        strImgListKey = "MouseDownBlue"
                    End If
                Else
                    strImgListKey = "MouseDownBlue"
                End If
        End Select
    ElseIf g.eTbSkin = eTbSkin_ALuminumBlue Then
        Select Case m.strStatus
            Case kKeyMouseNone, kBtnStateDisabled, kMouseJustUp
                strImgListKey = "MouseNoneAluminumBlue"
            Case kKeyMouseMove
                strImgListKey = "MouseMoveAluminum"
            Case Else
                If m.strStatus = kKeyMouseInProg Then
                    If g.RealTime.ConnectionStatus = eGDConnectionStatus_Connecting Or _
                       g.RealTime.ConnectionStatus = eGDConnectionStatus_Connected Then
                        strImgListKey = "MouseNoneAluminumBlue"
                    Else
                        strImgListKey = "MouseDownAluminumBlue"
                    End If
                Else
                    strImgListKey = "MouseDownAluminumBlue"
                End If
        End Select
    ElseIf g.eTbSkin = eTbSkin_AluminumSilver Then
        Select Case m.strStatus
            Case kKeyMouseNone, kBtnStateDisabled, kMouseJustUp
                strImgListKey = "MouseNoneAluminumSilver"
            Case kKeyMouseMove
                strImgListKey = "MouseMoveAluminum"
            Case Else
                If m.strStatus = kKeyMouseInProg Then
                    If g.RealTime.ConnectionStatus = eGDConnectionStatus_Connecting Or _
                       g.RealTime.ConnectionStatus = eGDConnectionStatus_Connected Then
                        strImgListKey = "MouseNoneAluminumSilver"
                    Else
                        strImgListKey = "MouseDownAluminumSilver"
                    End If
                Else
                    strImgListKey = "MouseDownAluminumSilver"
                End If
        End Select
    ElseIf g.eTbSkin = eTbSkin_DarkFlat Then
        Select Case m.strStatus
            Case kKeyMouseNone, kBtnStateDisabled, kMouseJustUp
                strImgListKey = "kMouseNoneDark"
            Case kKeyMouseMove
                strImgListKey = "kMouseMoveFlat"
            Case Else
                If m.strStatus = kKeyMouseInProg Then
                    If g.RealTime.ConnectionStatus = eGDConnectionStatus_Disconnected Then
                        strImgListKey = "kMouseNoneDark"
                    Else
                        strImgListKey = "kMouseMoveFlat"
                    End If
                Else
                    strImgListKey = "kMouseMoveFlat"
                End If
        End Select
    ElseIf g.eTbSkin = eTbSkin_LightFlat Then
        Select Case m.strStatus
            Case kKeyMouseNone, kBtnStateDisabled, kMouseJustUp
                strImgListKey = "kMouseNoneLight"
            Case kKeyMouseMove
                strImgListKey = "kMouseMoveFlat"
            Case Else
                If m.strStatus = kKeyMouseInProg Then
                    If g.RealTime.ConnectionStatus = eGDConnectionStatus_Disconnected Then
                        strImgListKey = "kMouseNoneLight"
                    Else
                        strImgListKey = "kMouseMoveFlat"
                    End If
                Else
                    strImgListKey = "kMouseMoveFlat"
                End If
        End Select
    Else
        Select Case m.strStatus
            Case kKeyMouseNone, kBtnStateDisabled, kMouseJustUp
                strImgListKey = "MouseNoneSilver"
            Case kKeyMouseMove
                strImgListKey = "MouseMove"
            Case Else
                If m.strStatus = kKeyMouseInProg Then
                    If ProcessIsBusy(True) And frmStatus.Status = eStatus_Aborted Or frmStatus.Status = eStatus_Aborting Then
                        strImgListKey = "MouseDownSilver"
                    Else
                        If g.RealTime.ConnectionStatus = eGDConnectionStatus_Connecting Or _
                           g.RealTime.ConnectionStatus = eGDConnectionStatus_Connected Then
                            strImgListKey = "MouseDownSilver"
                        Else
                            strImgListKey = "MouseNoneSilver"
                        End If
                    End If
                Else
                    strImgListKey = "MouseDownSilver"
                End If
        End Select
    End If
    
    If m.nWidth < 59 Then
        iWidth = m.nWidth
    Else
        If m.nHeight = 22 Then
            iWidth = 110
        Else
            iWidth = 118
        End If
    End If

    Select Case iWidth
        Case 59
            Select Case m.strStatus
                Case kKeyMouseNone, kBtnStateDisabled, kMouseJustUp
                    If m.strID = "ID_TradeTracker" Then
                        Set imgToUse = g.CoreBridge.ImageListMouse59.ListImages(strTradeTrackerKey)
                    Else
                        Set imgToUse = g.CoreBridge.ImageListMouse59.ListImages(strImgListKey)
                    End If
                Case kKeyMouseMove
                    Set imgToUse = g.CoreBridge.ImageListMouse59.ListImages(strImgListKey)
                Case Else
                    Set imgToUse = g.CoreBridge.ImageListMouse59.ListImages(strImgListKey)
            End Select
        Case 22
            Select Case m.strStatus
                Case kKeyMouseNone, kBtnStateDisabled, kMouseJustUp
                    If m.strID = "ID_TradeTracker" Then
                        If g.eTbSkin = eTbSkin_ALuminumBlue Or g.eTbSkin = eTbSkin_AluminumSilver Then
                            Set imgToUse = g.CoreBridge.ImageListMouse45.ListImages(strTradeTrackerKey)
                        Else
                            Set imgToUse = g.CoreBridge.ImageListMouse22.ListImages(strTradeTrackerKey)
                        End If
                    Else
                        Set imgToUse = g.CoreBridge.ImageListMouse22.ListImages(strImgListKey)
                    End If
                Case kKeyMouseMove
                    Set imgToUse = g.CoreBridge.ImageListMouse22.ListImages(strImgListKey)
                Case Else
                    Set imgToUse = g.CoreBridge.ImageListMouse22.ListImages(strImgListKey)
            End Select
        Case kBtnLargeIco
            Select Case m.strStatus
                Case kKeyMouseNone, kBtnStateDisabled, kMouseJustUp
                    If m.strID = "ID_TradeTracker" Then
                        Set imgToUse = g.CoreBridge.ImageListMouse59.ListImages(strTradeTrackerKey)
                    Else
                        Set imgToUse = g.CoreBridge.ImageListMouse35.ListImages(strImgListKey)
                    End If
                Case kKeyMouseMove
                    Set imgToUse = g.CoreBridge.ImageListMouse35.ListImages(strImgListKey)
                Case Else
                    Set imgToUse = g.CoreBridge.ImageListMouse35.ListImages(strImgListKey)
            End Select
        Case 45
            Select Case m.strStatus
                Case kKeyMouseNone, kBtnStateDisabled, kMouseJustUp
                    If m.strID = "ID_TradeTracker" Then
                        If g.eTbSkin = eTbSkin_ALuminumBlue Or g.eTbSkin = eTbSkin_AluminumSilver Then
                            Set imgToUse = g.CoreBridge.ImageListMouse45.ListImages(strTradeTrackerKey)
                        Else
                            Set imgToUse = g.CoreBridge.ImageListMouse22.ListImages(strTradeTrackerKey)
                        End If
                    Else
                        Set imgToUse = g.CoreBridge.ImageListMouse45.ListImages(strImgListKey)
                    End If
                Case kKeyMouseMove
                    Set imgToUse = g.CoreBridge.ImageListMouse45.ListImages(strImgListKey)
                Case Else
                    Set imgToUse = g.CoreBridge.ImageListMouse45.ListImages(strImgListKey)
            End Select
        Case 110
            Select Case m.strStatus
                Case kKeyMouseNone, kBtnStateDisabled, kMouseJustUp
                    If m.strID = "ID_TradeTracker" Then
                        Set imgToUse = g.CoreBridge.ImageListMouse22.ListImages(strTradeTrackerKey)
                    Else
                        Set imgToUse = g.CoreBridge.ImageListMouse110.ListImages(strImgListKey)
                    End If
                Case kKeyMouseMove
                    Set imgToUse = g.CoreBridge.ImageListMouse110.ListImages(strImgListKey)
                Case Else
                    Set imgToUse = g.CoreBridge.ImageListMouse110.ListImages(strImgListKey)
            End Select
        Case 118
            Select Case m.strStatus
                Case kKeyMouseNone, kBtnStateDisabled, kMouseJustUp
                    If m.strID = "ID_TradeTracker" Then
                        If g.eTbSkin = eTbSkin_ALuminumBlue Or g.eTbSkin = eTbSkin_AluminumSilver Then
                            Set imgToUse = g.CoreBridge.ImageListMouse118.ListImages(strTradeTrackerKey)
                        Else
                            Set imgToUse = g.CoreBridge.ImageListMouse22.ListImages(strTradeTrackerKey)
                        End If
                    Else
                        Set imgToUse = g.CoreBridge.ImageListMouse118.ListImages(strImgListKey)
                    End If
                Case kKeyMouseMove
                    Set imgToUse = g.CoreBridge.ImageListMouse118.ListImages(strImgListKey)
                Case Else
                    Set imgToUse = g.CoreBridge.ImageListMouse118.ListImages(strImgListKey)
            End Select
    End Select
    
    Set ToolBtnBkground = imgToUse

End Function

Private Function ToolBtnIcon() As ListImage
On Error Resume Next:

    Dim strIcon$, strTbName$, iSize&
    
    Dim Chart As cChart
    Dim imgToUse As ListImage
    
    If m.nTextOnly = 1 Then
        If g.eTbSkin = eTbSkin_DarkFlat Then
            If m.strStatus = kBtnStateDown Or m.strStatus = kKeyMouseDown Then
                m.nTextColor = 0
            Else
                m.nTextColor = RGB(238, 238, 238)
            End If
        ElseIf g.eTbSkin = eTbSkin_LightFlat Then
            If m.strStatus = kBtnStateDown Or m.strStatus = kKeyMouseDown Then
                m.nTextColor = RGB(238, 238, 238)
            Else
                m.nTextColor = 0
            End If
        End If
        Exit Function
    End If

    If Not m.frm Is Nothing Then
        If IsFrmChart(m.frm) Then
            Set Chart = m.frm.Chart
        ElseIf Not ActiveChart Is Nothing Then
            Set Chart = ActiveChart.Chart
        End If
    End If
    
    Select Case m.eCategory
        Case eBtnCat_General
            strTbName = "General"
        Case eBtnCat_Window
            strTbName = "Windows"
        Case eBtnCat_Charting
            strTbName = "Charting"
        Case eBtnCat_Drawing
            strTbName = "Draw"
    End Select
    
    If g.nTbLargeIcons = 1 Then
        iSize = 32
    Else
        iSize = 16
    End If
    
    Select Case m.strID
        Case "ID_ZoomOut"
            If m.strStatus = kBtnStateDisabled Then
                strIcon = "kZoomOutDisabled"
            Else
                strIcon = "kZoomOut"
            End If
        Case "ID_ChartMove"
            If Chart Is Nothing Then
                strIcon = "kChartMoveHz"
            ElseIf Chart.AutoScale Then
                strIcon = "kChartMoveHz"
            Else
                strIcon = "kChartMove"
            End If
        Case "ID_UndoDraw"
            If Chart Is Nothing Then
                strIcon = "kUndoDisabled"
            ElseIf Chart.LastEditedAnnot Is Nothing Then
                strIcon = "kUndoDisabled"
            Else
                strIcon = "kUndoEnabled"
            End If
        Case "ID_ShowEWI"
            strIcon = "kElliotLabelsOff"
            If Not Chart Is Nothing Then
                If -1 <> Chart.HasHiddenAnnots(eANNOT_ElliotLabel) Then strIcon = "kElliotLabelsOn"
            End If
        Case "ID_ChartOnOff"
            If BtnState = eBtnState_Neutral Then
                strIcon = "kChartOnOff_Show"
            Else
                strIcon = "kChartOnOff_Hide"
            End If
        Case "ID_RealTime"
            Select Case m.strStatus
                Case kKeyMouseDown, kBtnStateDown
                    If g.nReplaySession = 0 Then
                        strIcon = "kGreenLight"
                    Else
                        strIcon = "kGrayLight"
                    End If
                Case kKeyMouseInProg
                    If g.nReplaySession = 0 Then
                        strIcon = "kYellowLight"
                    Else
                        strIcon = "kGrayLight"
                    End If
                Case Else
                    strIcon = "kRedLight"
            End Select
        Case "ID_TextIncrease"
            If m.strStatus = kBtnStateDisabled Then
                strIcon = "kTextIncreaseDisabled"
            Else
                strIcon = "kTextIncrease"
            End If
        Case "ID_TextDecrease"
            If m.strStatus = kBtnStateDisabled Then
                strIcon = "kTextDecreaseDisabled"
            Else
                strIcon = "kTextDecrease"
            End If
    
        Case Else
            strIcon = ToolbarIcon(m.strID)
    End Select
    
'JM 10-21-2015: new code for dark, light color scheme
    If g.nTbIconStyle = 0 Then
        Set imgToUse = g.CoreBridge.ImgListToolbarExt("Classic", strIcon, strTbName, iSize)
    ElseIf m.strID = "ID_TradeTracker" Then
        Set imgToUse = g.CoreBridge.ImgListToolbarExt("Dark", strIcon, strTbName, iSize)
        m.nTextColor = 0
    ElseIf g.eTbSkin = eTbSkin_DarkFlat Then
            Set imgToUse = g.CoreBridge.ImgListToolbarExt("Light", strIcon, strTbName, iSize)
            m.nTextColor = RGB(238, 238, 238)
    Else
        Set imgToUse = g.CoreBridge.ImgListToolbarExt("Dark", strIcon, strTbName, iSize)
        m.nTextColor = 0
    End If
    
    Set ToolBtnIcon = imgToUse

End Function

Private Sub BtnSyncEach(pbBk As Variant, aButtons As cGdArray)
On Error Resume Next:

    Dim i&
    Dim bCheckGroup As Boolean
    
    Dim pB As PictureBox
    Dim oButton As cPicBoxButton

    If pbBk Is Nothing Then Exit Sub
    If aButtons Is Nothing Then Exit Sub
    
    If m.eGroup <> eBtnGroup_None And m.eGroup <> eBtnGroup_Unknown Then bCheckGroup = True

    For i = 0 To aButtons.Size - 2
        Set pB = Nothing
        Set oButton = aButtons(i)
        
        If Not oButton Is Nothing Then
            If oButton.PicboxIndex >= 0 And oButton.PicboxIndex <= pbBk.UBound Then
                Set pB = pbBk(oButton.PicboxIndex)
            End If
        End If
        
        If Not pB Is Nothing Then
            If oButton.BtnID = m.strID Then
                If oButton.BtnGroup <> eBtnGroup_None Then
                    If oButton.BtnStatus <> m.strStatus Then
                        oButton.BtnStatus = m.strStatus
                        oButton.Draw pB, aButtons
                        pB.Refresh
                    End If
                End If
                If Not bCheckGroup Then Exit For
            ElseIf bCheckGroup Then
                If oButton.BtnGroup = m.eGroup Then
                    oButton.BtnClearNow pB, aButtons
                End If
            End If
        End If
    Next

End Sub

Private Sub BtnSyncAll()
On Error Resume Next:

    Dim i&
    
    Dim frm As Form
    Dim Chart As cChart
    Dim pbBk As Variant
    Dim aButtons As cGdArray
       
    If m.strToolbar <> kTbDraw Then Exit Sub
    
    If Not TypeOf m.frm Is frmMain Then
        Set aButtons = frmMain.TbButtonsArray(m.strToolbar)
        If m.strToolbar = kTbDraw And g.vbeTbAlignDraw = vbAlignBottom And frmMain.pbTbBack(0).Visible Then
            Set pbBk = frmMain.pbTbBack
        Else
            Set pbBk = frmMain.pbTbBackDraw
        End If
                        
        If Not pbBk Is Nothing Then BtnSyncEach pbBk, aButtons
    End If
    
    For i = 0 To Forms.Count - 1
        Set pbBk = Nothing
        Set Chart = Nothing
        Set frm = Forms(i)
        If IsFrmChart(frm) Then
            If Not frm Is m.frm Then
                If frm.DetachStatus = eDetached Then
                    Set Chart = frm.Chart
                    If Not Chart Is Nothing Then
                        If Chart.ShowToolbar Then
                            Set aButtons = frm.TbButtonsArray(m.strToolbar)
                            If m.strToolbar = kTbDraw And g.vbeTbAlignDraw = vbAlignBottom And frm.pbTbBack(0).Visible Then
                                Set pbBk = frm.pbTbBack
                            Else
                                Set pbBk = frm.pbTbBackDraw
                            End If
                            If Not pbBk Is Nothing Then BtnSyncEach pbBk, aButtons
                        End If
                    End If
                End If
            End If
        End If
    Next

End Sub

Public Sub BtnStructInfo(iLeft&, iTop&, iWidth&, iHeight&, iIconSize&)
On Error Resume Next:

    iLeft = m.nLeft
    iTop = m.nTop
    iWidth = m.nWidth
    iHeight = m.nHeight
    
    If g.nTbLargeIcons Then
        iIconSize = 32
    ElseIf g.nTbIncludeText Then
        iIconSize = 18
    Else
        iIconSize = 16
    End If

End Sub

Private Function IsDropdownBtn() As Boolean
On Error Resume Next:

    Dim bIsDropdown As Boolean
    
    Select Case m.strID
        Case "ID_Templates", "ID_Pages", "ID_Sectors", "ID_Subsectors", "ID_Components"
            bIsDropdown = True
    End Select

    IsDropdownBtn = bIsDropdown

End Function

Private Sub HandleDropDownBtns()
On Error Resume Next:

    Dim strID$
    Dim oButton As cPicBoxButton
    Dim eFormMode As eTemplateFormMode
    
    strID = m.strID

    If strID = "ID_Templates" Or strID = "ID_Pages" Or _
        strID = "ID_Sectors" Or strID = "ID_Subsectors" Or strID = "ID_Components" Then
                
        Dim nX&, nY&
                
        If IsFrmChart(m.frm) Then
            'FixFocusChart frm
        ElseIf TypeOf m.frm Is frmMain Then
            If Not ActiveChart Is Nothing Then
                'user clicked a tool button on main app toolbar when detached chart has focus
                If ActiveChart.DetachStatus = eDetached Then
                    If ActiveChart.Chart.ShowToolbar Then
                        If Not g.ChartGlobals.frmActiveNonDetached Is Nothing Then
                            ActiveChart.SkipFocusFix = True
                            SendMessage ActiveChart.hWnd, WM_ACTIVATE, 0, 0
                            SendMessage ActiveChart.hWnd, WM_NCACTIVATE, 0, 0
                            SendMessage g.ChartGlobals.frmActiveNonDetached.hWnd, WM_MOUSEACTIVATE, 1, 0
                            SendMessage g.ChartGlobals.frmActiveNonDetached.hWnd, WM_NCACTIVATE, 1, 0
                        End If
                    Else
                        ActiveChart.SkipFocusFix = True
                    End If
                End If
            End If
        End If
        
        If strID = "ID_Templates" Then
            eFormMode = eMode_Templates
        ElseIf strID = "ID_Pages" Then
            eFormMode = eMode_Pages
        Else
            eFormMode = eMode_SecSubCom
        End If
        
        If TypeOf m.frm Is frmTbMoreButtons Then
            nX = (m.frm.Left + m.frm.ScaleWidth) / Screen.TwipsPerPixelX
            Set oButton = ButtonByID(frmTbMoreButtons.FormSource, kMoreBtnID, kTbGeneral)
            If oButton Is Nothing Then
                nY = m.nHeight * (m.nPicboxIndex + 1)
            Else
                nY = oButton.BtnTop
            End If
        Else
            nX = m.nLeft
            nY = m.nHeight * (m.nPicboxIndex + 1)
        End If

        frmTemplatePage.ShowMe m.frm, eFormMode, nX, nY, m.strID, m.strCaption
    End If

End Sub

Public Property Get BtnLeft() As Long
    BtnLeft = m.nLeft
End Property

Public Property Get BtnTop() As Long
    BtnTop = m.nTop
End Property

Public Sub BtnToolTipShow(ByRef frm As Form, ByRef vCtrl As Control)
On Error Resume Next:

    Dim l&, t&, r&, b&

    l = m.nLeft
    t = m.nTop
    r = l + m.nWidth
    b = m.nTop + m.nHeight
    
    Set m.oToolTip = New cToolTip
        
    With m.oToolTip
        .Create frm
        
        .DelayTime(ttDelayInitial) = 0
        .DelayTime(ttDelayReshow) = 0
        .DelayTime(ttDelayShow) = 3000
        .MaxTipWidth = 1
        
        .AddRect vCtrl, l, t, r, b, m.strTip, False
    End With
    
End Sub
