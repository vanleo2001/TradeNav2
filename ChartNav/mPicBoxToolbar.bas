Attribute VB_Name = "mPicBoxToolbar"
Option Explicit

'toolbar property name & section
Public Const kTbTemplate = "ToolbarTemplate"
Public Const kTbTemplateDate = "ToolbarTemplateDate"
Public Const kTbIniSection = "Toolbars"
'toolbar names
Public Const kTbGeneral = "General"
Public Const kTbWindows = "Windows"
Public Const kTbChartSettings = "Chart Settings"
Public Const kTbDraw = "Drawing Tools"
'visible properties
Public Const kTbGeneralVisible = "TB_General_Visible"
Public Const kTbWindowsVisible = "TB_Windows_Visible"
Public Const kTbChartSettingsVisible = "TB_ChartSetting_Visible"
Public Const kTbDrawVisible = "TB_Draw_Visible"

'drawing tool background when vertical
Public Const kSkinDrawToolBlue = 15121294
Public Const kSkinDrawToolSilver = 14208712
'toolbar button imagelist keys
Public Const kKeyMouseDown = "_MouseDown"
Public Const kKeyMouseMove = "_MouseMove"
Public Const kKeyMouseNone = "_MouseNone"
Public Const kKeyMouseInProg = "_MouseInProg"
'identifies 'just happened' MouseUp event for toolbar button
Public Const kMouseJustUp = "_MouseUp"
Public Const kBtnStateDown = "_StateDown"
Public Const kBtnStateDisabled = "_StateDisabled"
'dimension of buttons for background picbox - must be larger than ICO sizes
Public Const kBtnLargeIcoTextWd = 68    '78
Public Const kBtnLargeIcoTextHt = 52
Public Const kBtnSmallIcoTextWd = 59
Public Const kBtnSmallIcoTextHt = 35
Public Const kBtnSmallIco = 22          'buttons with no text have same width/height
Public Const kBtnLargeIco = 36
'ID of More button
Public Const kMoreBtnID = "ID_MoreButtons"

Public Enum eBtnStyle
    eBtnStyle_Unknown = -1
    eBtnStyle_Push
    eBtnStyle_State
    eBtnStyle_Dropdown
    eBtnStyle_Combobox
End Enum

Public Enum eBtnState
    eBtnState_Unknown = -1
    eBtnState_Neutral               'enabled, mouse-up picture
    eBtnState_Selected              'enabled, mouse-down picture
    eBtnState_Disabled              'disabled
    eBtnState_InProg
End Enum

Public Enum eTbSkin
    eTbSkin_Unknown = -1
    eTbSkin_Silver = 0
    eTbSkin_Blue
    eTbSkin_AluminumSilver
    eTbSkin_ALuminumBlue
    eTbSkin_DarkFlat
    eTbSkin_LightFlat
End Enum

Public Enum eBtnGroup
    eBtnGroup_Unknown = -1
    eBtnGroup_None
    eBtnGroup_BarDisplayType
    eBtnGroup_BarPeriods
    eBtnGroup_Cursor
    eBtnGroup_Drawing
    eBtnGroup_ChartDropDown
End Enum

Public Enum eBtnCategory
    eBtnCat_Unknown = -1
    eBtnCat_None
    eBtnCat_General
    eBtnCat_Window
    eBtnCat_Charting
    eBtnCat_Drawing
End Enum

Public Enum eBtnCfgFields
    eBtnCfgField_ID             'string read in from config file
    eBtnCfgField_Name           'string read in from config file
    eBtnCfgField_Group          'string read in from config file
    eBtnCfgField_Style          'string read in from config file
    eBtnCfgField_Category       'string read in from config file
    eBtnCfgField_Toolbar        'string read in from config file
    eBtnCfgField_Tooltip        'string read in from config file
    
    eBtnCfgField_TextOnly       'string read in from config file (0 or 1 flag)
    
    eBtnCfgField_eStyle         'string converted to enumerated type
    eBtnCfgField_eGroup         'string converted to enumerated type
    eBtnCfgField_eCategory      'string converted to enumerated type
End Enum

Public Enum eBtnProperties
    eBtnProp_ID = 1
    eBtnProp_Status
    eBtnProp_Style
    eBtnProp_Group
    eBtnProp_Category
    eBtnProp_TextOnlyFlag
    eBtnProp_Caption
    eBtnProp_LastButtonIdx
    eBtnProp_Tooltip
    eBtnProp_ToolbarName       'string of toolbar name passed in at initialization
    eBtnProp_ToolbarPos        'enum of toolbar position passed in at initialization
End Enum

Private Type mPrivate
    tblButtonConfig As cGdTable     'table containing button style, properties etc. info (read in from config file)
    aTblButtonIdx As cGdArray       'index for button config table
    oBtnLastProcessed As cPicBoxButton

    frmTextIncDec As Form           'form that the +T / -T buttons apply to
    bFuncOverride As Boolean
End Type

Private m As mPrivate

Public Function BtnStyleToEnum(ByVal strStyle$) As eBtnStyle
On Error GoTo ErrSection:

    Dim eStyle As eBtnStyle
    
    Select Case UCase(strStyle)
        Case "PUSH"
            eStyle = eBtnStyle_Push
        Case "STATE"
            eStyle = eBtnStyle_State
        Case "DROPDOWN"
            eStyle = eBtnStyle_Dropdown
        Case "COMBOBOX"
            eStyle = eBtnStyle_Combobox
        Case Else
            eStyle = eBtnStyle_Unknown
    End Select
    
    BtnStyleToEnum = eStyle

ErrExit:
    Exit Function

ErrSection:
    RaiseError "mPicBoxToolbar.BtnStyleToEnum"

End Function

Public Function BtnGroupToEnum(ByVal strGroup$) As eBtnGroup
On Error GoTo ErrSection:

    Dim eGroup As eBtnGroup
    
    If Len(strGroup) = 0 Then
        eGroup = eBtnGroup_None
    Else
        Select Case UCase(strGroup)
            Case "BARDISPLAYTYPE"
                eGroup = eBtnGroup_BarDisplayType
            Case "BARPERIODS"
                eGroup = eBtnGroup_BarPeriods
            Case "CURSOR"
                eGroup = eBtnGroup_Cursor
            Case "DRAWING"
                eGroup = eBtnGroup_Drawing
            Case "CHARTDROPDOWN"
                eGroup = eBtnGroup_ChartDropDown
            Case Else
                eGroup = eBtnGroup_Unknown
        End Select
    End If
        
    BtnGroupToEnum = eGroup
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mPicBoxToolbar.BtnGroupToEnum"

End Function

Public Function BtnCategoryToEnum(ByVal strCategory$) As eBtnCategory
On Error GoTo ErrSection:

    Dim eCategory As eBtnCategory
    
    If Len(strCategory) = 0 Then
        eCategory = eBtnCat_None
    Else
        Select Case UCase(strCategory)
            Case "GENERAL"
                eCategory = eBtnCat_General
            Case "CHARTING"
                eCategory = eBtnCat_Charting
            Case "WINDOWS"
                eCategory = eBtnCat_Window
            Case Else
                eCategory = eBtnCat_Unknown
        End Select
    End If
    
    BtnCategoryToEnum = eCategory

ErrExit:
    Exit Function

ErrSection:
    RaiseError "mPicBoxToolbar.BtnCategoryToEnum"

End Function

Public Function BtnConfigLoad(ByVal strFile$) As Long
On Error GoTo ErrSection:

    Dim aFile As cGdArray
    Dim strStyle$, strGroup$, strCategory$
    Dim i&, j&
        
    If m.tblButtonConfig Is Nothing Then
        Set m.tblButtonConfig = New cGdTable
        m.tblButtonConfig.CreateField eGDARRAY_Strings, eBtnCfgField_ID, "ButtonID"
        m.tblButtonConfig.CreateField eGDARRAY_Strings, eBtnCfgField_Name, "ButtonName"
        m.tblButtonConfig.CreateField eGDARRAY_Strings, eBtnCfgField_Group, "ButtonGroup"
        m.tblButtonConfig.CreateField eGDARRAY_Strings, eBtnCfgField_Style, "ButtonStyle"
        m.tblButtonConfig.CreateField eGDARRAY_Strings, eBtnCfgField_Category, "ButtonCategory"
        m.tblButtonConfig.CreateField eGDARRAY_Strings, eBtnCfgField_Toolbar, "ButtonToolbar"
        m.tblButtonConfig.CreateField eGDARRAY_Strings, eBtnCfgField_TextOnly, "ButtonTextOnly"
        
        m.tblButtonConfig.CreateField eGDARRAY_Strings, eBtnCfgField_Tooltip, "ButtonToolTip"
        
        m.tblButtonConfig.CreateField eGDARRAY_Longs, eBtnCfgField_eStyle, -1
        m.tblButtonConfig.CreateField eGDARRAY_Longs, eBtnCfgField_eGroup, -1
        m.tblButtonConfig.CreateField eGDARRAY_Longs, eBtnCfgField_eCategory, -1
    End If
    
    m.tblButtonConfig.NumRecords = 0
    
    If FileExist(strFile) Then
        Set aFile = New cGdArray
        aFile.FromFile strFile
        aFile.Sort eGdSort_Default Or eGdSort_DeleteDuplicates
        For i = 1 To aFile.Size - 1
            strStyle = Parse(aFile(i), vbTab, eBtnCfgField_Style + 1)
            strGroup = Parse(aFile(i), vbTab, eBtnCfgField_Group + 1)
            strCategory = Parse(aFile(i), vbTab, eBtnCfgField_Category + 1)
            
            m.tblButtonConfig.AddRecord aFile(i)
            j = m.tblButtonConfig.NumRecords - 1
                                    
            m.tblButtonConfig(eBtnCfgField_eStyle, j) = BtnStyleToEnum(strStyle)
            m.tblButtonConfig(eBtnCfgField_eGroup, j) = BtnGroupToEnum(strGroup)
            If UCase(strGroup) = "DRAWING" Then
                m.tblButtonConfig(eBtnCfgField_eCategory, j) = eBtnCat_Drawing
            Else
                m.tblButtonConfig(eBtnCfgField_eCategory, j) = BtnCategoryToEnum(strCategory)
            End If
        Next
        Set m.aTblButtonIdx = m.tblButtonConfig.CreateIndex(eBtnCfgField_ID)
    End If
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "mPicBoxToolbar.BtnConfigLoad"
    
End Function

Public Function ButtonsFromMainTB(frmSource As Form, ByVal strTbFilter$) As cGdArray
On Error GoTo ErrSection:

    Dim i&, strItem$, strToolbar$, strText$

    Dim aItems As cGdArray
    Dim aButtons As New cGdArray
    Dim aMainTbArray As cGdArray
    
    Dim bShow As Boolean
    Dim bTbGeneral As Boolean
    Dim bTbWindows As Boolean
    Dim bTbDraw As Boolean
    
    Dim bTbChart As Boolean
    Dim bFrmChart As Boolean
    
    Dim oButton As cPicBoxButton
    
    Set aItems = frmToolbar.aItems
    If aItems Is Nothing Then GoTo ErrExit
    
    If TypeOf frmSource Is frmMain Then
        bTbGeneral = GetIniFileProperty(kTbGeneralVisible, True, "Toolbars", g.strIniFile)
        bTbWindows = GetIniFileProperty(kTbWindowsVisible, True, "Toolbars", g.strIniFile)
        bTbChart = GetIniFileProperty(kTbChartSettingsVisible, True, "Toolbars", g.strIniFile)
        bTbDraw = GetIniFileProperty(kTbDrawVisible, True, "Toolbars", g.strIniFile)
        
        If g.bStarting Then
            If Not bTbGeneral And Not bTbWindows And Not bTbChart Then
                bTbGeneral = True           '5846: work-around fix until can figure out what real problem is
                If frmMain.tbToolbar.ToolBars(kTbGeneral).Tools.Count = 0 Then
                    If Len(strTbFilter) = 0 Then aButtons.Add "ID_Download"
                End If
            End If
        End If
        
        If Not bTbGeneral And Not bTbWindows And Not bTbChart Then
            If frmMain.pbTbBack(0).Visible Then frmMain.pbTbBack(0).Visible = False
            If frmMain.imgTbBack(0).Visible Then frmMain.imgTbBack(0).Visible = False
        Else
            If Not frmMain.pbTbBack(0).Visible Then frmMain.pbTbBack(0).Visible = True
            If Not frmMain.imgTbBack(0).Visible Then frmMain.imgTbBack(0).Visible = True
        End If
        If frmMain.pbTbBackDraw(0).Visible <> bTbDraw Then frmMain.pbTbBackDraw(0).Visible = bTbDraw
        If frmMain.imgTbBackDraw(0).Visible <> bTbDraw Then frmMain.imgTbBackDraw(0).Visible = bTbDraw
        
        If Not bTbDraw Then frmMain.TbButtonsArray(kTbDraw).Size = 0
        
    ElseIf IsFrmChart(frmSource) Then
        bTbChart = GetIniFileProperty(kTbChartSettingsVisible, True, "Toolbars", g.strIniFile)
        bTbDraw = GetIniFileProperty(kTbDrawVisible, True, "Toolbars", g.strIniFile)
        bFrmChart = True
        
        If frmSource.pbTbBack(0).Visible <> bTbChart Then frmSource.pbTbBack(0).Visible = bTbChart
        If frmSource.imgTbBack(0).Visible <> bTbChart Then frmSource.imgTbBack(0).Visible = bTbChart
        
        If frmSource.pbTbBackDraw(0).Visible <> bTbDraw Then frmSource.pbTbBackDraw(0).Visible = bTbDraw
        If frmSource.imgTbBackDraw(0).Visible <> bTbDraw Then frmSource.imgTbBackDraw(0).Visible = bTbDraw
        
        If strTbFilter = kTbDraw Then
            If bTbDraw Then
                Set aMainTbArray = frmMain.TbButtonsArray(kTbDraw)
                If Not aMainTbArray Is Nothing Then
                    For i = 0 To aMainTbArray.Size - 1
                        aButtons.Add aMainTbArray(i).BtnID
                    Next
                End If
            Else
                frmSource.TbButtonsArray(kTbDraw).Size = 0
            End If
        ElseIf bTbChart Then
            Set aMainTbArray = frmMain.TbButtonsArray(kTbGeneral)
            If Not aMainTbArray Is Nothing Then
                For i = 0 To aMainTbArray.Size - 1
                    Set oButton = aMainTbArray(i)
                    If (oButton.BtnCategory = eBtnCat_Charting And oButton.BtnID <> "ID_Tile") Or _
                       oButton.BtnID = "ID_Symbol" Or oButton.BtnID = "ID_Chart" Then
                        aButtons.Add aMainTbArray(i).BtnID
                    End If
                Next
            End If
            frmSource.ToolBarWrapSet kTbGeneral, frmMain.ToolBarWrapGet(kTbGeneral)
        End If
        
        Set ButtonsFromMainTB = aButtons
        
        GoTo ErrExit
        
    End If

    If Not bTbGeneral And Not bTbWindows And Not bTbChart And Not bTbDraw Then Exit Function
    
    If Len(strTbFilter) = 0 Then
        'JM 03-10-2009: this loop mimics the loop in frmToolbar.ShowMe
        For i = 0 To aItems.Size - 1
            strItem = aItems(i)
            If Left(strItem, 1) = "=" Then
                strToolbar = Mid(strItem, 2)
            ElseIf frmMain.tbToolbar.Tools(strItem).Visible Then

                If bFrmChart Then
                    If strItem = "ID_Tile" Or strToolbar = kTbGeneral Or strToolbar = kTbWindows Then
                        'do not add
                    ElseIf strToolbar = kTbChartSettings And bTbChart Then
                        On Error Resume Next
                        bShow = False
                        bShow = frmMain.tbToolbar.ToolBars(strToolbar).Tools.Item(strItem).Visible
                        If bShow Then aButtons.Add strItem
                    End If
                ElseIf (strToolbar = kTbGeneral And bTbGeneral) Or _
                       (strToolbar = kTbWindows And bTbWindows) Or _
                       (strToolbar = kTbChartSettings And bTbChart) Then

                    On Error Resume Next
                    bShow = False
                    bShow = frmMain.tbToolbar.ToolBars(strToolbar).Tools.Item(strItem).Visible
                    If bShow Then aButtons.Add strItem
                End If

            End If
        Next
    ElseIf strTbFilter = kTbDraw And bTbDraw Then
        For i = 0 To aItems.Size - 1
            strItem = aItems(i)
            If Left(strItem, 1) = "=" Then
                strToolbar = Mid(strItem, 2)
            ElseIf frmMain.tbToolbar.Tools(strItem).Visible Then
                If strToolbar = strTbFilter Then
                    On Error Resume Next
                    bShow = False
                    bShow = frmMain.tbToolbar.ToolBars(strToolbar).Tools.Item(strItem).Visible
                    If bShow Then aButtons.Add strItem
                End If
            End If
        Next
    End If
    
    Set ButtonsFromMainTB = aButtons
    
ErrExit:
    Set aItems = Nothing
    Set aButtons = Nothing
    Exit Function

ErrSection:
    Set aItems = Nothing
    Set aButtons = Nothing
    RaiseError "mPicBoxToolbar.ButtonsFromMainTB"

End Function

Public Sub ToolbarInit2(frm As Form, aButtons As cGdArray, _
    Optional aButtonsShow As cGdArray = Nothing, _
    Optional ByVal strToolbar$ = "", _
    Optional ByVal strFile$ = "", _
    Optional ByVal vbeToolbarAlign As Long = vbAlignTop, _
    Optional ByVal bUseMainTbButtons As Boolean = True, _
    Optional ByVal bShowMoreButtonsPB As Boolean = True)
On Error GoTo ErrSection:

    Dim i&, strID$
    
    Dim aShow As cGdArray
    
    Dim oButton As cPicBoxButton
    Dim oBtnSector As cPicBoxButton
    Dim oBtnSubsector As cPicBoxButton
    Dim oBtnComp As cPicBoxButton

    Dim bSyncSector As Boolean
    Dim bLocked As Boolean
    
    If g.bUnloading Then Exit Sub
    
    If frm Is Nothing Then Exit Sub
    If aButtons Is Nothing Then Exit Sub
    
    If m.tblButtonConfig Is Nothing Then BtnConfigLoad App.Path & "\Provided\Toolbuttons.cfg"

    If Not aButtonsShow Is Nothing Then
        If aButtonsShow.Size > 0 Then
            Set aShow = aButtonsShow
        End If
    ElseIf Len(strFile) > 0 Then
        If FileExist(strFile) Then
            Set aShow = New cGdArray
            aShow.FromFile strFile
        Else
            Exit Sub
        End If
    ElseIf bUseMainTbButtons Then
        bLocked = LockWindowUpdate(GetDesktopWindow())          '5855, 6255
        Set aShow = ButtonsFromMainTB(frm, strToolbar)
        If bLocked Then LockWindowUpdate 0
    End If

    If aShow Is Nothing Then Exit Sub
    If aShow.Size <= 0 Then Exit Sub            '4955
    aShow.Add "ID_MoreButtons"   'always make this the last button

    If IsFrmChart(frm) Then     'set toolbar skin on detached chart to same as toolbar skin on main app form
        If strToolbar = kTbDraw Then
            If frm.imgTbBackDraw(0).Picture <> frmMain.imgTbBackDraw(0).Picture Then
                frm.imgTbBackDraw(0).Picture = frmMain.imgTbBackDraw(0).Picture
            End If
            If frm.pbTbBackDraw(0).BackColor <> frmMain.pbTbBackDraw(0).BackColor Then
                frm.pbTbBackDraw(0).BackColor = frmMain.pbTbBackDraw(0).BackColor
            End If
        ElseIf frm.imgTbBack(0).Picture <> frmMain.imgTbBack(0).Picture Then
            frm.imgTbBack(0).Picture = frmMain.imgTbBack(0).Picture
        End If
    End If

    aButtons.Clear
    For i = 0 To aShow.Size - 1
        strID = aShow(i)
        Set oButton = New cPicBoxButton
        If oButton.PropertiesSet(frm, m.tblButtonConfig, m.aTblButtonIdx, strID, strToolbar, vbeToolbarAlign, aButtons.Size) Then
            Select Case strID
                Case "ID_ChartMove"
                    If g.ChartGlobals.eChartMode = eMode_Move Then oButton.BtnStatus = kBtnStateDown
                Case "ID_Eraser"
                    If g.ChartGlobals.eChartMode = eMode_Erase Then oButton.BtnStatus = kBtnStateDown
                Case "ID_Magnet"
                    If g.ChartGlobals.nMagnetValue > 0 Then oButton.BtnStatus = kBtnStateDown
                Case "ID_ZoomIn"
                    If g.ChartGlobals.eChartMode <> eMode_Move And g.ChartGlobals.eChartMode <> eMode_Erase Then
                        oButton.BtnStatus = kBtnStateDown
                    End If
                Case "ID_ZoomOut"
                    oButton.BtnStatus = kBtnStateDisabled
                Case "ID_TextIncrease", "ID_TextDecrease"
                    If TypeOf frm Is frmChart2 Then
                        oButton.BtnStatus = kKeyMouseNone       '+T/-T should never be disabled on detached chart's toolbar
                    Else
                        oButton.BtnStatus = kBtnStateDisabled
                    End If
            End Select
            
            aButtons.Add oButton
            
            If g.RealTime.Active Then
                If Not oButton Is Nothing Then
                    If oButton.BtnID = "ID_RealTime" Then
                        If g.RealTime.ConnectionStatus = eGDConnectionStatus_Connected Then
                            oButton.BtnState = eBtnState_Selected           '6254
                        Else
                            oButton.BtnState = eBtnState_InProg
                        End If
                    End If
                End If
            End If
        
            If TypeOf frm Is frmTbMoreButtons Then
                Select Case strID
                    Case "ID_Sectors"
                        Set oBtnSector = oButton
                        bSyncSector = True
                    Case "ID_Subsectors"
                        Set oBtnSubsector = oButton
                        bSyncSector = True
                    Case "ID_Components"
                        Set oBtnComp = oButton
                        bSyncSector = True
                End Select
            End If
        End If
    Next
    
    Dim Chart As cChart

    If Not g.bStarting And Not g.bLoadingChartPage Then     '5455
        If TypeOf frm Is frmMain Then
            If Not ActiveChart Is Nothing Then
                If ActiveChart.DetachStatus = eDetached Then
                    If ActiveChart.Chart.ShowToolbar = 0 Then
                        Set Chart = ActiveChart.Chart
                    ElseIf Not g.ChartGlobals.frmActiveNonDetached Is Nothing Then
                        Set Chart = g.ChartGlobals.frmActiveNonDetached.Chart
                    End If
                Else
                    Set Chart = ActiveChart.Chart
                End If
            End If
        ElseIf TypeOf frm Is frmChart2 Then
            Set Chart = frm.Chart
        End If
        
        If Not Chart Is Nothing Then
            If FormIsLoaded("frmToolbar") Then
                If frmToolbar.Visible Then
                    'do nothing - issue 6254
                Else
                    Chart.SyncToolbar True
                    SyncFormBtns
                End If
            Else
                Chart.SyncToolbar True
                SyncFormBtns
            End If
        End If
    End If
    
'JM 04-14-2011: do not do this, causes issue 6254
'    If TypeOf frm Is frmMain Then
'        g.RealTime.SetRTbutton 0        'call with zero to simulate a "reset" of RT button
'        g.RealTime.SetRTbutton -1       'call with -1 to sync button with realtime streaming status
'    End If

ErrExit:
    Set aShow = Nothing
    Set oButton = Nothing
    Exit Sub

ErrSection:
    Set aShow = Nothing
    Set oButton = Nothing
    RaiseError "mPicBoxToolbar.BtnStyleToEnum"

End Sub

Private Function ToolbarResizeHorz(frm As Form, pbTbBack As Variant, imgBack As Variant, aButtons As cGdArray, _
    ByVal bWrap As Boolean, ByVal xStart&, ByVal yStart&) As Long
On Error GoTo ErrSection:

    Dim i&, iRight&, iTop&
    Dim iLastVisibleIdx&, iPicboxIdx&
    
    Dim bBarPeriodShow As Boolean
    Dim bDrawMoreButton As Boolean
    
    Dim lpRect As Rect
    
    Dim oButton As cPicBoxButton
    Dim oButtonPrev As cPicBoxButton
    Dim oButtonNext As cPicBoxButton
    
    Dim pB As PictureBox
    Dim pBGeneral As Variant
    
    Dim imgToUse As ListImage

    If frm Is Nothing Then Exit Function
    If pbTbBack Is Nothing Then Exit Function
    If imgBack Is Nothing Then Exit Function
    If aButtons Is Nothing Then Exit Function
    If aButtons.Size < 1 Then Exit Function
    
    GetClientRect frm.hWnd, lpRect
                            
    aButtons(aButtons.Size - 1).BtnDrawIndexReset
    
    'select background skin
    Select Case g.eTbSkin
        Case eTbSkin_Unknown, eTbSkin_Silver, eTbSkin_AluminumSilver
            Set imgToUse = g.CoreBridge.ImageListBackgroundSkin.ListImages("kBkSilver")
        Case eTbSkin_Blue
            Set imgToUse = g.CoreBridge.ImageListBackgroundSkin.ListImages("kBkBlue")
        Case eTbSkin_ALuminumBlue
            Set imgToUse = g.CoreBridge.ImageListBackgroundSkin.ListImages("kBkAluminumBlue")
        Case eTbSkin_DarkFlat
            Set imgToUse = g.CoreBridge.ImageListBackgroundSkin.ListImages("kBkDarkFlat")
        Case eTbSkin_LightFlat
            Set imgToUse = g.CoreBridge.ImageListBackgroundSkin.ListImages("kBkLightFlat")
    End Select
    
    Set oButton = aButtons(0)
    If oButton.ToolBarName = kTbDraw Then
        'make sure drawing toolbar is drawn below all general toolbar (ie in case of wrap)
        Set pBGeneral = frm.pbTbBack
        For i = 0 To pBGeneral.UBound           '6547
            If pBGeneral(i).Visible = False Then Exit For
        Next
        If i <= 0 Then
            i = pBGeneral.UBound
        Else
            i = i - 1
        End If
        pbTbBack(0).AutoRedraw = True
        pbTbBack(0).Move 0, pBGeneral(i).Top + pBGeneral(i).Height, frm.Width, (oButton.BtnHeight + 2) * Screen.TwipsPerPixelY
        imgBack(0).Picture = Nothing
        imgBack(0).Move 0, 0, pbTbBack(0).Width, pbTbBack(0).Height
    Else
        pbTbBack(0).AutoRedraw = True
        pbTbBack(0).Move 0, 0, frm.Width, (oButton.BtnHeight + 2) * Screen.TwipsPerPixelY
        imgBack(0).Picture = Nothing
        imgBack(0).Move 0, 0, pbTbBack(0).Width, pbTbBack(0).Height
    End If

    geOverlayICO pbTbBack(0).hDC, imgToUse.ExtractIcon, 0, pbTbBack(0).Width / Screen.TwipsPerPixelX, oButton.BtnHeight + 3, 0, 0, vbBlack, ""
        
    iPicboxIdx = 0
    Set pB = pbTbBack(0)
    
    'draw first button
    Set oButton = aButtons(0)
    oButton.LastVisibleBtnIndex = 0
    oButton.Draw pB, aButtons, 0
    If oButton.BtnID = "ID_BarPeriod" Then
        bBarPeriodShow = True
        Set frm.cboBarPeriod.Container = pbTbBack(0)
        frm.cboBarPeriod.ZOrder
    End If
    
    Set oButtonPrev = oButton
    
    For i = 1 To aButtons.Size - 2
        Set oButton = aButtons(i)
        Set oButtonNext = aButtons(i + 1)
        
        If bWrap Then
            'get pixel of right edge of previous button
            iRight = oButtonPrev.BtnLeft + oButtonPrev.BtnWidth
        Else
            'get pixel of where right edge of this button will be
            iRight = oButtonPrev.BtnLeft + oButtonPrev.BtnWidth + oButton.BtnWidth + 1
        End If
        
        If oButton.BtnID = kMoreBtnID Then
            Exit For
        ElseIf oButtonNext.BtnID = kMoreBtnID And Not bWrap Then
            If iRight < lpRect.Right Then
                'button after this one is the More button, so draw this one
                oButton.LastVisibleBtnIndex = i
                iLastVisibleIdx = i
                oButton.Draw pB, aButtons, iPicboxIdx
                If oButton.BtnID = "ID_BarPeriod" Then
                    bBarPeriodShow = True
                    Set frm.cboBarPeriod.Container = pbTbBack(iPicboxIdx)
                    frm.cboBarPeriod.ZOrder
                End If
            Else
                bDrawMoreButton = True
                Exit For
            End If
        ElseIf (lpRect.Right - iRight) > oButtonNext.BtnWidth Then
            oButton.LastVisibleBtnIndex = i
            iLastVisibleIdx = i
            oButton.Draw pB, aButtons, iPicboxIdx
            If oButton.BtnID = "ID_BarPeriod" Then
                bBarPeriodShow = True
                Set frm.cboBarPeriod.Container = pbTbBack(iPicboxIdx)
                frm.cboBarPeriod.ZOrder
            End If
            Set oButtonPrev = oButton
        ElseIf bWrap Then
            iPicboxIdx = iPicboxIdx + 1
            If pbTbBack.UBound < iPicboxIdx Then Load pbTbBack(iPicboxIdx)
            If imgBack.UBound < iPicboxIdx Then Load imgBack(iPicboxIdx)
            Set imgBack(iPicboxIdx).Picture = Nothing
            pbTbBack(iPicboxIdx).AutoRedraw = True
            Set imgBack(iPicboxIdx).Container = pbTbBack(iPicboxIdx)
            iTop = pbTbBack(iPicboxIdx - 1).Top + pbTbBack(iPicboxIdx - 1).Height
            pbTbBack(iPicboxIdx).Move 0, iTop, pbTbBack(iPicboxIdx - 1).Width, pbTbBack(iPicboxIdx - 1).Height
            imgBack(iPicboxIdx).Stretch = True
            imgBack(iPicboxIdx).Move 0, 0, pbTbBack(iPicboxIdx).Width, pbTbBack(iPicboxIdx).Height
            imgBack(iPicboxIdx).Picture = Nothing
            If Not pbTbBack(iPicboxIdx).Visible Then pbTbBack(iPicboxIdx).Visible = True
            If Not imgBack(iPicboxIdx).Visible Then imgBack(iPicboxIdx).Visible = True
            geOverlayICO pbTbBack(iPicboxIdx).hDC, imgToUse.ExtractIcon, 0, pbTbBack(iPicboxIdx).Width / Screen.TwipsPerPixelX, oButton.BtnHeight + 2, 0, 0, 0, ""
            
            Set pB = pbTbBack(iPicboxIdx)
            
            oButton.LastVisibleBtnIndex = i
            iLastVisibleIdx = i
            oButton.Draw pB, aButtons, iPicboxIdx
            If oButton.BtnID = "ID_BarPeriod" Then
                bBarPeriodShow = True
                Set frm.cboBarPeriod.Container = pbTbBack(iPicboxIdx)
                frm.cboBarPeriod.ZOrder
            End If
            Set oButtonPrev = oButton
        Else
            bDrawMoreButton = True
            Exit For
        End If
    Next
        
    If bDrawMoreButton Then
        Set oButtonNext = aButtons(aButtons.Size - 1)   'set next button to the More button
        
        'note: we might have gotten here because the remaining width could not accommodate
        '   a regular sized button AND a double or triple width button, but it may be able
        '   to accomodate 2 regular sized buttons (i.e. the More and a regular size button)
        If (lpRect.Right - iRight) > oButtonNext.BtnWidth Then
            'this button & the more button both fit within remaining width so draw both
            iLastVisibleIdx = oButton.BtnIndex
            oButton.LastVisibleBtnIndex = oButton.BtnIndex
            oButton.Draw pB, aButtons, iPicboxIdx
            If oButton.BtnID = "ID_BarPeriod" Then
                bBarPeriodShow = True
                Set frm.cboBarPeriod.Container = pbTbBack(iPicboxIdx)
                frm.cboBarPeriod.ZOrder
            End If
            Set oButtonPrev = oButton
        Else
            'only more button will fit within remaining width, don't draw current button
            oButton.LastVisibleBtnIndex = -1
        End If
        
        oButtonNext.LastVisibleBtnIndex = iLastVisibleIdx
        oButtonNext.Draw pB, aButtons, iPicboxIdx
    End If
    
    'reset draw info indexes
    For i = 0 To iLastVisibleIdx
        aButtons(i).LastVisibleBtnIndex = iLastVisibleIdx
    Next
    
    For i = iLastVisibleIdx + 1 To aButtons.Size - 2
        aButtons(i).LastVisibleBtnIndex = -1
        aButtons(i).BtnDrawIndexReset
    Next
   
''''''''''''''''''''''''''''''''''''''''''''''''''''''
'finish miscellaneous tasks
''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For i = 0 To pbTbBack.UBound
        If i > iPicboxIdx Then
            If pbTbBack(i).Visible Then pbTbBack(i).Visible = False
            If imgBack(i).Visible Then imgBack(i).Visible = False
        End If
    Next
    
    If aButtons(0).ToolBarName <> kTbDraw Then
        If bBarPeriodShow And frm.cboBarPeriod.ListCount = 0 Then BarPeriodCboInit frm.cboBarPeriod
        frm.cboBarPeriod.Visible = bBarPeriodShow
        frm.cboBarPeriod.Enabled = bBarPeriodShow
    End If
        
    pbTbBack(0).Refresh
    
'    ToolbarResizeHorz = 0      - JM 7/20/2009: not using this for now


ErrExit:
    Exit Function

ErrSection:
    RaiseError "mPicBoxToolbar.ToolbarResizeHorz"

End Function

Private Function ToolbarResizeVert(frm As Form, pbTbBack As Variant, imgBack As Variant, aButtons As cGdArray, _
    ByVal bForceWrap As Boolean, ByVal xStart&, ByVal yStart&) As Long
On Error Resume Next:
    
    Dim i&, iBottom&, iLastVisibleIdx&
    Dim bDrawMoreButton As Boolean
    Dim lpRect As Rect
    
    Dim oButton As cPicBoxButton
    Dim oButtonPrev As cPicBoxButton
    Dim oButtonNext As cPicBoxButton
    
    Dim pB As PictureBox

    If frm Is Nothing Then Exit Function
    If pbTbBack Is Nothing Then Exit Function
    If imgBack Is Nothing Then Exit Function
    If aButtons Is Nothing Then Exit Function

    imgBack(0).Picture = Nothing
    pbTbBack(0).AutoRedraw = True
    
    Select Case g.eTbSkin
        Case eTbSkin_Unknown, eTbSkin_Silver, eTbSkin_AluminumSilver
            pbTbBack(0).BackColor = kSkinDrawToolSilver
        Case eTbSkin_Blue, eTbSkin_ALuminumBlue
            pbTbBack(0).BackColor = kSkinDrawToolBlue
        Case eTbSkin_DarkFlat
            pbTbBack(0).BackColor = 0
        Case eTbSkin_LightFlat
            pbTbBack(0).BackColor = RGB(238, 238, 238)      'this is btn neutral skin color
    End Select
    
    Set oButton = aButtons(0)
    
    If g.vbeTbAlignDraw = vbAlignLeft Then
        pbTbBack(0).align = vbAlignLeft
        pbTbBack(0).Left = 0
    Else
        pbTbBack(0).align = vbAlignRight
        pbTbBack(0).Left = frmMain.Width - (oButton.BtnWidth * Screen.TwipsPerPixelX)
    End If
    pbTbBack(0).Top = 0
    pbTbBack(0).Width = oButton.BtnWidth * Screen.TwipsPerPixelX
    pbTbBack(0).Height = frmMain.ScaleHeight
        
    imgBack(0).Move 0, 0, pbTbBack(0).Width, pbTbBack(0).Height

    GetClientRect pbTbBack(0).hWnd, lpRect
    
    aButtons(aButtons.Size - 1).BtnDrawIndexReset       '5458
    
    Set pB = pbTbBack(0)
    'draw first button
    Set oButton = aButtons(0)
    oButton.LastVisibleBtnIndex = 0
    oButton.Draw pB, aButtons, 0
    Set oButtonPrev = oButton
    
    lpRect.Bottom = lpRect.Bottom - oButton.BtnHeight
    For i = 1 To aButtons.Size - 2
        Set oButton = aButtons(i)
        Set oButtonNext = aButtons(i + 1)
        iBottom = oButtonPrev.BtnTop + oButtonPrev.BtnHeight + oButton.BtnHeight + 1
        
        If oButton.BtnID = kMoreBtnID Then
            Exit For
        ElseIf oButtonNext.BtnID = kMoreBtnID Then
            If iBottom < lpRect.Bottom Then
                iLastVisibleIdx = i
                oButton.LastVisibleBtnIndex = i
                oButton.Draw pB, aButtons, 0
                Set oButtonPrev = oButton
            Else
                bDrawMoreButton = True
            End If
            Exit For
        ElseIf iBottom < lpRect.Bottom Then
            iLastVisibleIdx = i
            oButton.LastVisibleBtnIndex = i
            oButton.Draw pB, aButtons, 0
            Set oButtonPrev = oButton
        Else
            bDrawMoreButton = True
            Exit For
        End If
    Next

    If bDrawMoreButton Then
        oButton.LastVisibleBtnIndex = -1
        
        Set oButton = aButtons(aButtons.Size - 1)   'set button to the More button
        oButton.LastVisibleBtnIndex = iLastVisibleIdx
        oButton.Draw pB, aButtons, 0
    End If

    'reset draw info indexes
    For i = 0 To iLastVisibleIdx
        aButtons(i).LastVisibleBtnIndex = iLastVisibleIdx
    Next
    
    For i = iLastVisibleIdx + 1 To aButtons.Size - 2
        aButtons(i).LastVisibleBtnIndex = -1
        aButtons(i).BtnDrawIndexReset
    Next
   
''''''''''''''''''''''''''''''''''''''''''''''''''''''
'finish miscellaneous tasks
''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Not pbTbBack(0).Visible Then pbTbBack(0).Visible = True
    If Not imgBack(0).Visible Then imgBack(0).Visible = True
    pbTbBack(0).Refresh
    
    ToolbarResizeVert = 1           'this is buttons per row

End Function

Public Function ToolbarResize2(frm As Form, pbInUse As Variant, imgInUse As Variant, aButtons As cGdArray, _
    ByVal bWrap As Boolean, Optional ByVal Left& = -1, Optional ByVal Top& = -1) As Long
On Error Resume Next:

    Static bInProgress As Boolean
        
    Dim bLocked As Boolean

    Dim i&, X&, Y&, iAlign&
    Dim oButton As cPicBoxButton
    
    If bInProgress Then Exit Function
    
    If frm Is Nothing Then Exit Function
    If pbInUse Is Nothing Then Exit Function
    If imgInUse Is Nothing Then Exit Function
    If aButtons Is Nothing Then Exit Function
    If frmMain.WindowState = vbMinimized Then Exit Function     '5185
    
    If Not TypeOf frm Is frmTbMoreButtons Then
        If FormIsLoaded("frmTbMoreButtons") Then Unload frmTbMoreButtons
'        bLocked = LockWindowUpdate(GetDesktopWindow())          '6255
        
        'JM 10-06-2011 - original fix above, caused position confirm to not draw. Tim changed the showform
        '   function to always unlock windows for modal forms to take care of position confirm issue.
        '   Changed this code anyways as it is not necessary to lock entire desktop.
        bLocked = LockWindowUpdate(frmMain.hWnd)
    End If
    
    Set oButton = aButtons(0)
    
    bInProgress = True
    
    If oButton Is Nothing Or (pbInUse(0).Visible = False And Not TypeOf frm Is frmTbMoreButtons) Then
        For i = 0 To pbInUse.UBound
            If pbInUse(i).Visible Then pbInUse(i).Visible = False     '5153
            If imgInUse(i).Visible Then imgInUse(i).Visible = False
        Next
        GoTo ErrExit
    ElseIf frm.pbTbBack(0).Visible = False Then
        If g.vbeTbAlignDraw = vbAlignBottom And (oButton.ToolBarName = kTbDraw Or TypeOf frm Is frmTbMoreButtons) Then
            'user selected side-by-side, but only drawing toolbar is visible
            X = 2 * Screen.TwipsPerPixelX
            Y = 1 * Screen.TwipsPerPixelY
            pbInUse.align = vbAlignTop
            i = ToolbarResizeHorz(frm, pbInUse, imgInUse, aButtons, bWrap, X, Y)
            GoTo ErrExit
        End If
    End If
    
    If g.vbeTbAlignDraw = vbAlignBottom Then
        iAlign = vbAlignBottom
    Else
        iAlign = oButton.ToolBarPos
    End If
        
    If Left >= 0 Then
        X = Left
    ElseIf iAlign = vbAlignRight Or iAlign = vbAlignLeft Then
        X = 1 * Screen.TwipsPerPixelX
    Else
        X = 2 * Screen.TwipsPerPixelX
    End If
    
    If Top >= 0 Then
        Y = Top
    ElseIf oButton.ToolBarName = kTbDraw And iAlign = vbAlignTop Then
        'position drawing toolbar
        If frmMain.pbTbBack(0).Visible Then
            'draw tool bar needs to be below general toolbar when general toolbar is visible
            For i = 1 To frmMain.pbTbBack.UBound
                If frmMain.pbTbBack(i).Visible Then
                    Y = frmMain.pbTbBack(i).Top + frmMain.pbTbBack(i).Height
                Else
                    Exit For
                End If
            Next
        Else
            Y = 1 * Screen.TwipsPerPixelY
        End If
    Else
        'general toolbar always starts flush top
        Y = 1 * Screen.TwipsPerPixelY
    End If
    
    If iAlign = vbAlignTop Then
        pbInUse.align = vbAlignTop
        If oButton.ToolBarName = kTbDraw And frm.pbTbBack(0).Visible Then
            'both general & drawing toolbars are set at top, don't wrap drawing toolbar
            i = ToolbarResizeHorz(frm, pbInUse, imgInUse, aButtons, False, X, Y)
        Else
            i = ToolbarResizeHorz(frm, pbInUse, imgInUse, aButtons, bWrap, X, Y)
        End If
    ElseIf iAlign = vbAlignBottom Then
        frm.pbTbBackDraw(0).align = vbAlignTop
        frm.pbTbBackDraw(0).Visible = False
        frm.imgBackDraw(0).Visible = False
        
        If oButton.ToolBarName <> kTbDraw Then
            pbInUse.align = vbAlignTop
        End If
        
        i = ToolbarSideBySide(frm, pbInUse, imgInUse, bWrap)        '6437 - need to always redraw
    Else
        i = ToolbarResizeVert(frm, pbInUse, imgInUse, aButtons, False, X, Y)
    End If
       
ErrExit:
    ToolbarResize2 = i
    
    If bLocked Then LockWindowUpdate 0
        
    If IsFrmChart(frm) Then
        If Not frm.Chart Is Nothing Then frm.Chart.SyncToolbar True, True
    ElseIf Not g.ChartGlobals.frmActiveNonDetached Is Nothing Then
        If Not g.ChartGlobals.frmActiveNonDetached.Chart Is Nothing Then
            g.ChartGlobals.frmActiveNonDetached.Chart.SyncToolbar True, True
        End If
    End If

    bInProgress = False
    Exit Function
    
ErrSection:
    If bLocked Then LockWindowUpdate 0
    bInProgress = False
    Set oButton = Nothing

End Function

Public Function ToolbarHeight(pB As Variant) As Long
On Error GoTo ErrSection:

    Dim i&, iHeight&

    If pB Is Nothing Then Exit Function
    
    For i = 0 To pB.UBound
        If pB(i).Visible Then iHeight = iHeight + pB(i).Height
    Next

    ToolbarHeight = iHeight

ErrExit:
    Exit Function

ErrSection:
    RaiseError "mPicBoxToolbar.ToolbarHeight"

End Function

Private Sub SyncStudyBtns(frm As Form, Chart As cChart, pbDraw As Variant, aBtnDraw As cGdArray, _
    ByVal bResyncStudyBtns As Boolean)
On Error Resume Next

    Static frmPrevious As Form
    Static iPrevToolbarShow As Long

    Dim i&, idxPricePane&
    
    Dim Tree As cGdTree
    Dim Pane As cPane
    Dim Ind As cIndicator
    
    Dim strCodedName$, strID$
    
    Dim eBtnNewState As eBtnState
    
    'JP indicators
    Dim bJPPivotD As Boolean
    Dim bJPResistD As Boolean
    Dim bJPSupportD As Boolean
    Dim bJPAvgD As Boolean
    
    Dim bJPPivotW As Boolean
    Dim bJPResistW As Boolean
    Dim bJPSupportW As Boolean
    Dim bJPAvgW As Boolean
    
    Dim bJPPivotM As Boolean
    Dim bJPResistM As Boolean
    Dim bJPSupportM As Boolean
    Dim bJPAvgM As Boolean
    
    Dim bJPPivotQ As Boolean
    Dim bJPResistQ As Boolean
    Dim bJPSupportQ As Boolean
    Dim bJPAvgQ As Boolean
    
    Dim bJPPivotE As Boolean
    Dim bJPResistE As Boolean
    Dim bJPSupportE As Boolean
    Dim bJPAvgE As Boolean
    
    'Dinapoli indicators
    Dim bIndOscPredUpper As Boolean
    Dim bIndOscPredLower As Boolean
    Dim bIndMacdPredictor As Boolean
    Dim bIndDinapMA As Boolean
    
    'Dinapoli panes
    Dim bCheckPanes As Boolean
    Dim PaneMacD As cPane
    Dim PaneDetrendOsc As cPane
    Dim PanePrefStoch As cPane
    
    Dim aIndicators As cGdArray
    
    If g.bUnloading Or g.bLoadingChartPage Then Exit Sub
    
    bCheckPanes = HasModule("CTP")
    
    If frm Is Nothing Then Exit Sub
    If Chart Is Nothing Then Exit Sub
    
    If Not bResyncStudyBtns Then
        If Chart.Form Is frmPrevious Then
            If Chart.RedoMode < eRedo5_RecalcInd And iPrevToolbarShow = Chart.ShowToolbar Then
                Exit Sub
            End If
        End If
    End If
    
    Set frmPrevious = Chart.Form
    iPrevToolbarShow = Chart.ShowToolbar
    
    If TypeOf frm Is frmMain Then
        If Chart.Form Is ActiveChart Then
            If Chart.Form.DetachStatus = eDetached And Chart.ShowToolbar <> 0 Then
                Exit Sub
            End If
        ElseIf Not Chart.Form Is g.ChartGlobals.frmActiveNonDetached Then
            Exit Sub
        End If
    End If
    
    Set Tree = Chart.Tree
    If Tree Is Nothing Then Exit Sub
    
    If bCheckPanes Then
        Set aIndicators = New cGdArray
        aIndicators.Create eGDARRAY_Objects
    End If
    
    idxPricePane = Tree.RelativeIndex("PRICE", eTREE_Parent)

    For i = 1 To Tree.Count
        If Tree.NodeLevel(i) = 0 Then
            Set Pane = Tree(i)
            If bCheckPanes Then
                Pane.CheckDinapoliStudies PaneMacD, PaneDetrendOsc, PanePrefStoch, aIndicators
                aIndicators.Size = 0
            End If
        ElseIf Tree.AncestorIndex(i, 0) = idxPricePane Then
            Set Ind = Tree(i)
            If Not Ind Is Nothing Then
                strCodedName = UCase(Ind.CodedName)
                If InStr(strCodedName, "JPPIVOTPOINT") <> 0 Then
                    If InStr(strCodedName, "DAILY") <> 0 Then
                        bJPPivotD = Ind.Display
                    ElseIf InStr(strCodedName, "WEEKLY") <> 0 Then
                        bJPPivotW = Ind.Display
                    ElseIf InStr(strCodedName, "MONTHLY") <> 0 Then
                        bJPPivotM = Ind.Display
                    ElseIf InStr(strCodedName, "QUARTERLY") <> 0 Then
                        bJPPivotQ = Ind.Display
                    ElseIf InStr(strCodedName, "EXP") <> 0 Then
                        bJPPivotE = Ind.Display
                    End If
                ElseIf InStr(strCodedName, "JPPIVOTRESISTANCE") <> 0 Then
                    If InStr(strCodedName, "DAILY") <> 0 Then
                        bJPResistD = Ind.Display
                    ElseIf InStr(strCodedName, "WEEKLY") <> 0 Then
                        bJPResistW = Ind.Display
                    ElseIf InStr(strCodedName, "MONTHLY") <> 0 Then
                        bJPResistM = Ind.Display
                    ElseIf InStr(strCodedName, "QUARTERLY") <> 0 Then
                        bJPResistQ = Ind.Display
                    ElseIf InStr(strCodedName, "EXP") <> 0 Then
                        bJPResistE = Ind.Display
                    End If
                ElseIf InStr(strCodedName, "JPPIVOTSUPPORT") <> 0 Then
                    If InStr(strCodedName, "DAILY") <> 0 Then
                        bJPSupportD = Ind.Display
                    ElseIf InStr(strCodedName, "WEEKLY") <> 0 Then
                        bJPSupportW = Ind.Display
                    ElseIf InStr(strCodedName, "MONTHLY") <> 0 Then
                        bJPSupportM = Ind.Display
                    ElseIf InStr(strCodedName, "QUARTERLY") <> 0 Then
                        bJPSupportQ = Ind.Display
                    ElseIf InStr(strCodedName, "EXP") <> 0 Then
                        bJPSupportE = Ind.Display
                    End If
                ElseIf InStr(strCodedName, "JPAVERAGEPIVOT") <> 0 Then
                    If InStr(strCodedName, "DAILY") <> 0 Then
                        bJPAvgD = Ind.Display
                    ElseIf InStr(strCodedName, "WEEKLY") <> 0 Then
                        bJPAvgW = Ind.Display
                    ElseIf InStr(strCodedName, "MONTHLY") <> 0 Then
                        bJPAvgM = Ind.Display
                    ElseIf InStr(strCodedName, "QUARTERLY") <> 0 Then
                        bJPAvgQ = Ind.Display
                    ElseIf InStr(strCodedName, "EXP") <> 0 Then
                        bJPAvgE = Ind.Display
                    End If
                ElseIf InStr(strCodedName, "UPPEROSCPRED") <> 0 Then
                    bIndOscPredUpper = Ind.Display
                ElseIf InStr(strCodedName, "LOWEROSCPRED") <> 0 Then
                    bIndOscPredLower = Ind.Display
                ElseIf InStr(strCodedName, "DINAPOLIMACDPREDICTOR") <> 0 Then
                    bIndMacdPredictor = Ind.Display
                ElseIf Ind.IsDinapoliDMA Then
                    bIndDinapMA = Ind.Display
                End If
            End If
        ElseIf bCheckPanes Then
            aIndicators.Add Ind
        End If
    Next
    
    'check the last pane
    If bCheckPanes Then
        If Not Pane Is Nothing Then
            Pane.CheckDinapoliStudies PaneMacD, PaneDetrendOsc, PanePrefStoch, aIndicators
        End If
    End If
    
    If Not Chart.tbToolbar Is Nothing Then
        With Chart.tbToolbar
            i = .Redraw
            .Redraw = False
            
            If bJPPivotD And bJPResistD And bJPSupportD And bJPAvgD Then
                .Tools("ID_JPDaily").State = ssChecked
            Else
                .Tools("ID_JPDaily").State = ssUnchecked
            End If
            
            If bJPPivotW And bJPResistW And bJPSupportW And bJPAvgW Then
                .Tools("ID_JPWeekly").State = ssChecked
            Else
                .Tools("ID_JPWeekly").State = ssUnchecked
            End If
            
            If bJPPivotM And bJPResistM And bJPSupportM And bJPAvgM Then
                .Tools("ID_JPMonthly").State = ssChecked
            Else
                .Tools("ID_JPMonthly").State = ssUnchecked
            End If
            
            If bJPPivotQ And bJPResistQ And bJPSupportQ And bJPAvgQ Then
                .Tools("ID_JPQuarterly").State = ssChecked
            Else
                .Tools("ID_JPQuarterly").State = ssUnchecked
            End If
            
            If bJPPivotE And bJPResistE And bJPSupportE And bJPAvgE Then
                .Tools("ID_JPExpiration").State = ssChecked
            Else
                .Tools("ID_JPExpiration").State = ssUnchecked
            End If
            
            'Dinap buttons
            If bIndOscPredUpper And bIndOscPredLower Then
                .Tools("ID_OscPredictor").State = ssChecked
            Else
                .Tools("ID_OscPredictor").State = ssUnchecked
            End If
            
            If bIndMacdPredictor Then
                .Tools("ID_MacdPredictor").State = ssChecked
            Else
                .Tools("ID_MacdPredictor").State = ssUnchecked
            End If
            
            If bIndDinapMA Then
                .Tools("ID_DisplacedMA").State = ssChecked
            Else
                .Tools("ID_DisplacedMA").State = ssUnchecked
            End If
            
            'dinapoli panes
            If PaneMacD Is Nothing Then
                .Tools("ID_DiNapoliMACD").State = ssUnchecked
            ElseIf PaneMacD.Display Then
                .Tools("ID_DiNapoliMACD").State = ssChecked
            Else
                .Tools("ID_DiNapoliMACD").State = ssUnchecked
            End If
            
            If PanePrefStoch Is Nothing Then
                .Tools("ID_PrefStoch").State = ssUnchecked
            ElseIf PanePrefStoch.Display Then
                .Tools("ID_PrefStoch").State = ssChecked
            Else
                .Tools("ID_PrefStoch").State = ssUnchecked
            End If
            
            If PaneDetrendOsc Is Nothing Then
                .Tools("ID_DetrendOsc").State = ssUnchecked
            ElseIf PaneDetrendOsc.Display Then
                .Tools("ID_DetrendOsc").State = ssChecked
            Else
                .Tools("ID_DetrendOsc").State = ssUnchecked
            End If
            
            .Redraw = i
        End With
    End If
    
    Dim oBtn As cPicBoxButton
    Dim pB As PictureBox
    
    If Not aBtnDraw Is Nothing And Not pbDraw Is Nothing Then
        For i = 0 To aBtnDraw.Size
            Set oBtn = aBtnDraw(i)
            If Not oBtn Is Nothing Then
                eBtnNewState = eBtnState_Neutral
                strID = oBtn.BtnID
                Select Case strID
                    Case "ID_OscPredictor"
                        If bIndOscPredUpper And bIndOscPredLower Then eBtnNewState = eBtnState_Selected
                    Case "ID_MacdPredictor"
                        If bIndMacdPredictor Then eBtnNewState = eBtnState_Selected
                    Case "ID_DisplacedMA"
                        If bIndDinapMA Then eBtnNewState = eBtnState_Selected
                    Case "ID_DiNapoliMACD"
                        If Not PaneMacD Is Nothing Then
                            If PaneMacD.Display Then eBtnNewState = eBtnState_Selected
                        End If
                    Case "ID_PrefStoch"
                        If Not PanePrefStoch Is Nothing Then
                            If PanePrefStoch.Display Then eBtnNewState = eBtnState_Selected
                        End If
                    Case "ID_DetrendOsc"
                        If Not PaneDetrendOsc Is Nothing Then
                            If PaneDetrendOsc.Display Then eBtnNewState = eBtnState_Selected
                        End If
                    Case Else
                        Set oBtn = Nothing
                End Select
            End If
            
            If Not oBtn Is Nothing Then
                If oBtn.BtnState <> eBtnNewState Then
                    oBtn.BtnState = eBtnNewState
                    If oBtn.BtnDrawIndex >= 0 Then      'draw index <0 means button is not drawn/visible
                        i = oBtn.PicboxIndex
                        If i >= 0 And i <= pbDraw.UBound Then
                            Set pB = pbDraw(i)
                            If Not pB Is Nothing Then
                                oBtn.Draw pB, aBtnDraw, i
                                pB.Refresh
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End If

    Dim aBtnArray As cGdArray
    Dim pbContainer As Variant

    Set aBtnArray = frm.TbButtonsArray(kTbChartSettings)
    Set pbContainer = frm.pbTbBack

    If Not aBtnArray Is Nothing And Not pbContainer Is Nothing Then
        For i = 0 To aBtnArray.Size
            Set oBtn = aBtnArray(i)
            If Not oBtn Is Nothing Then
                eBtnNewState = eBtnState_Neutral
                strID = oBtn.BtnID
                Select Case strID
                    Case "ID_JPDaily"
                        If bJPPivotD And bJPResistD And bJPSupportD And bJPAvgD Then eBtnNewState = eBtnState_Selected
                    Case "ID_JPWeekly"
                        If bJPPivotW And bJPResistW And bJPSupportW And bJPAvgW Then eBtnNewState = eBtnState_Selected
                    Case "ID_JPMonthly"
                        If bJPPivotM And bJPResistM And bJPSupportM And bJPAvgM Then eBtnNewState = eBtnState_Selected
                    Case "ID_JPQuarterly"
                        If bJPPivotQ And bJPResistQ And bJPSupportQ And bJPAvgQ Then eBtnNewState = eBtnState_Selected
                    Case "ID_JPExpiration"
                        If bJPPivotE And bJPResistE And bJPSupportE And bJPAvgE Then eBtnNewState = eBtnState_Selected
                    Case Else
                        Set oBtn = Nothing
                End Select
            End If
            
            If Not oBtn Is Nothing Then
                If oBtn.BtnState <> eBtnNewState Then
                    oBtn.BtnState = eBtnNewState
                    If oBtn.BtnDrawIndex >= 0 Then      'draw index <0 means button is not drawn/visible
                        i = oBtn.PicboxIndex
                        If i >= 0 And i <= pbContainer.UBound Then
                            Set pB = pbContainer(i)
                            If Not pB Is Nothing Then
                                oBtn.Draw pB, aBtnArray, i
                                pB.Refresh
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End If

End Sub

Private Sub SyncZoomMoveBtns(frm As Form, Chart As cChart, ByVal bResyncStudyBtns As Boolean)
On Error Resume Next:

    Dim aBtnArray As cGdArray
    Dim oBtn As cPicBoxButton, i&
    
    Dim pB As PictureBox
    Dim pbContainer As Variant
    Dim Tool As ActiveToolBars.SSTool

    Set aBtnArray = frm.TbButtonsArray(kTbDraw)
    If aBtnArray Is Nothing Then Exit Sub
    If aBtnArray.Size = 0 Then Exit Sub
    
    If g.vbeTbAlignDraw = vbAlignBottom And frm.pbTbBack(0).Visible Then
        Set pbContainer = frm.pbTbBack
    Else
        Set pbContainer = frm.pbTbBackDraw
    End If
    If pbContainer Is Nothing Then Exit Sub
    
    If g.ChartGlobals.eChartMode = eMode_Zoom Then
        Set oBtn = ButtonByID(frm, "ID_ZoomOut", kTbDraw)
        If Not oBtn Is Nothing Then
            If oBtn.BtnDrawIndex >= 0 Then
                i = oBtn.PicboxIndex
                If i >= 0 And i <= pbContainer.UBound Then
                    Set pB = pbContainer(i)
                    If Not pB Is Nothing Then
                        oBtn.BtnClearNow pB, aBtnArray
                        pB.Refresh
                    End If
                End If
            End If
        End If
    End If

    Set oBtn = ButtonByID(frm, "ID_ChartMove", kTbDraw)
    If Not oBtn Is Nothing Then
        If oBtn.BtnDrawIndex >= 0 Then      'draw index <0 means button is not drawn/visible
            i = oBtn.PicboxIndex
            If i >= 0 And i <= pbContainer.UBound Then
                Set pB = pbContainer(i)
                If Not pB Is Nothing Then
                    oBtn.Draw pB, aBtnArray, i
                    pB.Refresh
                End If
            End If
        End If
    End If
    
    Set oBtn = ButtonByID(frm, "ID_UndoDraw", kTbDraw)
    If Not oBtn Is Nothing Then
        If oBtn.BtnDrawIndex >= 0 Then      'draw index <0 means button is not drawn/visible
            i = oBtn.PicboxIndex
            If i >= 0 And i <= pbContainer.UBound Then
                Set pB = pbContainer(i)
                If Not pB Is Nothing Then
                    oBtn.Draw pB, aBtnArray, i
                    pB.Refresh
                End If
            End If
        End If
    End If
    
    Set oBtn = ButtonByID(frm, "ID_ShowEWI", kTbDraw)
    If Not oBtn Is Nothing Then
        If oBtn.BtnDrawIndex >= 0 Then      'draw index <0 means button is not drawn/visible
            i = oBtn.PicboxIndex
            If i >= 0 And i <= pbContainer.UBound Then
                Set pB = pbContainer(i)
                If Not pB Is Nothing Then
                    oBtn.Draw pB, aBtnArray, i
                    pB.Refresh
                End If
            End If
        End If
    End If
    
    'sync the menu item
    If Not ActiveChart Is Nothing Then
        If Not ActiveChart.Chart Is Nothing Then
            Set Tool = frmMain.tbToolbar.Tools("ID_UndoDraw")
            If Not Tool Is Nothing Then
                If ActiveChart.Chart.LastEditedAnnot Is Nothing Then
                    Tool.Enabled = False
                Else
                    Tool.Enabled = True
                End If
            End If
        End If
    End If
    
    SyncStudyBtns frm, Chart, pbContainer, aBtnArray, bResyncStudyBtns

End Sub

Public Sub SyncChartingBtns(frm As Form, ByVal strSec$, ByVal strSecDesc$, ByVal strSub$, ByVal strSubDesc$, _
    ByVal strCmp$, ByVal strCmpDesc$, ByVal strPeriodCbo$, ByVal strPeriod$, ByVal bResyncStudyBtns As Boolean, _
    Optional ByVal bDrawUndoOnly As Boolean = False)
On Error Resume Next:

    Static bInProgress As Boolean
    
    If g.bStarting Or g.bUnloading Or g.bLoadingChartPage Then Exit Sub
    
    If bInProgress Then Exit Sub

bInProgress = True
    
    Dim i&, idx&, strID$
    
    Dim Chart As cChart
    Dim Ind As cIndicator
    
    Dim pB As PictureBox
    Dim frmToolbar As Form
    
    Dim oBtn As cPicBoxButton
    Dim aBtnArray As cGdArray
    
    Dim oBtnIndType As cPicBoxButton
    Dim aIndTypeBtns As cGdArray
    
    Dim eBtnCurrState As eBtnState
        
    If g.bStarting Or g.bUnloading Or g.bLoadingChartPage Then GoTo ErrExit

    If frm Is Nothing Then GoTo ErrExit
    If Not IsFrmChart(frm) Then GoTo ErrExit
    
    Set Chart = frm.Chart
    If Chart Is Nothing Then GoTo ErrExit
    
    If IsFrmChartMDI(frm) Or Chart.ShowToolbar <> 1 Then
        'chart is not detached or detached chart does not have toolbar
        Set frmToolbar = frmMain
    Else
        Set frmToolbar = frm
    End If
    
    SyncZoomMoveBtns frmToolbar, Chart, bResyncStudyBtns
    
    'this flag is true when called from cChart to just sync the undoDraw button
    If bDrawUndoOnly Then GoTo ErrExit
    
    Set aBtnArray = frmToolbar.TbButtonsArray(kTbGeneral)
    If aBtnArray Is Nothing Then GoTo ErrExit
    
    Set Ind = Chart.Tree("PRICE")

    For i = 0 To aBtnArray.Size - 1
        Set oBtn = aBtnArray(i)
        If Not oBtn Is Nothing Then
            strID = aBtnArray(i).BtnID
            eBtnCurrState = oBtn.BtnState
            Select Case strID
                Case "ID_WhatIf"
                    If Chart.IsInWhatIfMode And eBtnCurrState <> eBtnState_Selected Then
                        oBtn.BtnState = eBtnState_Selected
                    ElseIf Not Chart.IsInWhatIfMode And eBtnCurrState = eBtnState_Selected Then
                        oBtn.BtnState = eBtnState_Neutral
                    Else
                        Set oBtn = Nothing      'clear object so draw code below won't get called
                    End If
                Case "ID_ChartOrderbar"
                    If Chart.ShowTrades = 2 And eBtnCurrState <> eBtnState_Selected Then
                        oBtn.BtnState = eBtnState_Selected
                    ElseIf Chart.ShowTrades <> 2 And eBtnCurrState = eBtnState_Selected Then
                        oBtn.BtnState = eBtnState_Neutral
                    Else
                        Set oBtn = Nothing      'clear object so draw code below won't get called
                    End If
                Case "ID_AutoScale"
                    If Chart.AutoScale And eBtnCurrState <> eBtnState_Selected Then
                        oBtn.BtnState = eBtnState_Selected
                    ElseIf Not Chart.AutoScale And eBtnCurrState = eBtnState_Selected Then
                        oBtn.BtnState = eBtnState_Neutral
                    Else
                        Set oBtn = Nothing
                    End If
                Case "ID_Sectors"
                    If g.nTbIncludeText = 1 And g.nTbLargeIcons = 1 Then
                        oBtn.BtnCaption = " " & Left(strSecDesc, 28)    '5779
                    Else
                        oBtn.BtnCaption = " " & Left(strSecDesc, 21)    '5779
                    End If
                    oBtn.BtnToolTip = strSec & ": " & strSecDesc
                Case "ID_Subsectors"
                    If g.nTbIncludeText = 1 And g.nTbLargeIcons = 1 Then
                        oBtn.BtnCaption = " " & Left(strSubDesc, 28)
                    Else
                        oBtn.BtnCaption = " " & Left(strSubDesc, 21)
                    End If
                    oBtn.BtnToolTip = strSub & ": " & strSubDesc
                Case "ID_Components"
                    oBtn.BtnCaption = strCmp
                    oBtn.BtnToolTip = strCmpDesc
                Case "ID_BarPeriod"
                    If Len(strPeriodCbo) > 0 Then
                        frmToolbar.cboBarPeriod.Text = strPeriodCbo              '5139
                    End If
                Case strPeriod
                    If eBtnCurrState <> eBtnState_Selected Then
                        oBtn.BtnState = eBtnState_Selected
                    Else
                        Set oBtn = Nothing
                    End If
                Case Else
                    If oBtn.BtnGroup = eBtnGroup_BarPeriods And oBtn.BtnID <> strPeriod Then
                        If eBtnCurrState = eBtnState_Selected Then
                            oBtn.BtnState = eBtnState_Neutral
                        Else
                            Set oBtn = Nothing
                        End If
                    ElseIf oBtn.BtnGroup = eBtnGroup_BarDisplayType And oBtnIndType Is Nothing Then
                        If aIndTypeBtns Is Nothing Then Set aIndTypeBtns = New cGdArray         '5242
                        Set oBtnIndType = PriceBarBtn(Ind, aIndTypeBtns, oBtn)
                        Set oBtn = Nothing
                    Else
                        Set oBtn = Nothing
                    End If
            End Select
        End If
        
        If Not oBtn Is Nothing Then
            If oBtn.BtnDrawIndex >= 0 Then
                idx = oBtn.PicboxIndex
                If idx >= 0 And idx <= frmToolbar.pbTbBack.UBound Then
                    Set pB = frmToolbar.pbTbBack(idx)
                    If Not pB Is Nothing And Not aBtnArray Is Nothing Then
                        oBtn.Draw pB, aBtnArray, idx
                    End If
                End If
            End If
        End If
    Next
    
    If oBtnIndType Is Nothing Then
        For i = 0 To aIndTypeBtns.Size - 1
            Set oBtn = aIndTypeBtns(i)
            If oBtn.BtnState <> eBtnState_Neutral Then
                idx = oBtn.PicboxIndex
                oBtn.BtnState = eBtnState_Neutral
                Set pB = frmToolbar.pbTbBack(idx)
                If Not pB Is Nothing Then oBtn.Draw pB, aBtnArray, idx
            End If
        Next
    ElseIf oBtnIndType.BtnState <> eBtnState_Selected Then
        oBtnIndType.BtnState = eBtnState_Selected
        idx = oBtnIndType.PicboxIndex
        Set pB = frmToolbar.pbTbBack(idx)
        If Not pB Is Nothing Then oBtnIndType.Draw pB, aBtnArray, idx
    End If
    
    'need to loop through picture box controls because a group button, eg bar period, may have gotten set
    'causing another in the group to be cleared, but the 2 buttons are on different picture box controls
    For i = 0 To frmToolbar.pbTbBack.UBound
        If frmToolbar.pbTbBack(i).Visible Then
            frmToolbar.pbTbBack(i).Refresh
        End If
    Next
    
ErrExit:
    Set pB = Nothing
    Set oBtn = Nothing
    Set aBtnArray = Nothing
    
    Set Chart = Nothing
    
    bInProgress = False

End Sub

Public Sub BarPeriodClick(frm As Form, oBtnMouseLast As cPicBoxButton, ByVal bDropdown As Boolean)
On Error GoTo ErrSection:

    Dim i&
    
    Dim pB As PictureBox
    Dim aButtons As cGdArray
    
    Dim oButton As cPicBoxButton
        
    If frm Is Nothing Then Exit Sub
    
    Set aButtons = frm.TbButtonsArray(kTbGeneral)
    If aButtons Is Nothing Then Exit Sub
    If aButtons.Size <= 0 Then Exit Sub
        
    For i = 0 To aButtons.Size - 2
        Set oButton = aButtons(i)
        If Not oButton Is Nothing Then
            If oButton.BtnID = "ID_BarPeriod" Then
                If oButton.PicboxIndex >= 0 And oButton.PicboxIndex <= frm.pbTbBack.UBound Then
                    Set pB = frm.pbTbBack(oButton.PicboxIndex)
                    If bDropdown Then
                        oButton.MouseDown frm, pB, aButtons, aButtons(aButtons.Size - 1)
                        Set oBtnMouseLast = oButton
                    Else
                        oButton.BtnClearNow pB, aButtons
                        Set oBtnMouseLast = Nothing
                    End If
                End If
                Exit For
            End If
        End If
    Next
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mPicBoxToolbar.BarPeriodClick"

End Sub

Private Sub SyncDrawBtns(ByVal strID$, ByVal eNewState As eBtnState)
On Error Resume Next

    Dim i&
    
    Dim aToolArray As cGdArray
    Dim oButton As cPicBoxButton
    Dim pB As PictureBox
  
'look for  button on main app window
    Set oButton = ButtonByID(frmMain, strID, kTbDraw)
    If oButton Is Nothing Then GoTo ErrExit
    
    Set aToolArray = frmMain.TbButtonsArray(kTbDraw)
    If aToolArray Is Nothing Then GoTo ErrExit
    
    If Not oButton Is Nothing And Not aToolArray Is Nothing Then
        i = oButton.PicboxIndex
        If i >= 0 Then
            If g.vbeTbAlignDraw = vbAlignBottom And frmMain.pbTbBack(0).Visible Then
                If i <= frmMain.pbTbBack.UBound Then Set pB = frmMain.pbTbBack(i)
            ElseIf i <= frmMain.pbTbBackDraw.UBound Then
                Set pB = frmMain.pbTbBackDraw(i)
            End If
            If Not pB Is Nothing Then
                oButton.MouseDown frmMain, pB, aToolArray, aToolArray(aToolArray.Size - 1)
                GoTo ErrExit
            End If
        End If
    End If

'look for button on detached charts
    Dim frm As Form
    Dim Chart As cChart
    
    For i = 0 To Forms.Count - 1
        Set oButton = Nothing
        Set aToolArray = Nothing
        
        Set pB = Nothing
        Set Chart = Nothing
        
        Set frm = Forms(i)
        If IsFrmChart(frm) Then
            If frm.DetachStatus = eDetached Then
                Set Chart = frm.Chart
                If Not Chart Is Nothing Then
                    If Chart.ShowToolbar Then
                        Set oButton = ButtonByID(frm, strID, kTbDraw)
                        Set aToolArray = frm.TbButtonsArray(kTbDraw)
                        
                        If Not oButton Is Nothing And Not aToolArray Is Nothing Then
                            If g.vbeTbAlignDraw = vbAlignBottom And frm.pbTbBack(0).Visible Then
                                Set pB = frm.pbTbBack(oButton.PicboxIndex)
                            Else
                                Set pB = frm.pbTbBackDraw(oButton.PicboxIndex)
                            End If
                            If Not pB Is Nothing Then
                                oButton.MouseDown frm, pB, aToolArray, aToolArray(aToolArray.Size - 1)
                                Exit For
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next
    

ErrExit:
    Set aToolArray = Nothing
    Set oButton = Nothing
    Set frm = Nothing
    Set Chart = Nothing

End Sub

Public Sub SyncStateButton(ByVal strID$, ByVal strGroup$, ByVal strCategory$, ByVal eNewState As eBtnState)

    Static bInProgress As Boolean
    
    Dim Chart As cChart, i&
    
    If g.bStarting Or g.bUnloading Or g.bLoadingChartPage Then Exit Sub

If bInProgress Then Exit Sub

bInProgress = True
    
    If Not m.oBtnLastProcessed Is Nothing Then
        If m.oBtnLastProcessed.BtnID = strID Then
            If m.oBtnLastProcessed.BtnState = eNewState Then GoTo ErrExit
        End If
    End If
    
    If UCase(strGroup) = "DRAWING" Then
        If Not ActiveChart Is Nothing Then
            Set Chart = ActiveChart.Chart
            If Not Chart Is Nothing Then
                If Chart.TypeOfChart = eTypeChart_Seasonal Then
                    Select Case strID
                        Case "ID_Trendline", "ID_TrendChannel", "ID_SRLine", _
                             "ID_HorzLine", "ID_VertLine", "ID_ArrowLine", _
                             "ID_Text", "ID_Icon", "ID_ElliotLabels", "ID_ElliotEndUser", _
                             "ID_Bracket", "ID_Ellipse", "ID_Rectangle"
                             
                             'okay to sync button
                             i = 0
                        
                        Case Else
                            GoTo ErrExit        'drawing tool not allowed, just exit - 6357
                    End Select
                End If
            End If
        End If
        SyncDrawBtns strID, eNewState
    ElseIf UCase(strCategory) = "GENERAL" Then
        SyncFormBtns
    End If
    
ErrExit:
Set m.oBtnLastProcessed = Nothing
bInProgress = False

End Sub

Private Sub SyncFormBtns()
On Error Resume Next:

    Dim i&
    Dim oBtn As cPicBoxButton
    
    Dim pB As PictureBox
    Dim aButtons As cGdArray
    
    If g.bStarting Or g.bUnloading Or g.bLoadingChartPage Then Exit Sub
    
    Set aButtons = frmMain.TbButtonsArray(kTbGeneral)
    
    If aButtons Is Nothing Then Exit Sub
    If aButtons.Size <= 0 Then Exit Sub
    
    For i = 0 To aButtons.Size - 1
        Set oBtn = aButtons(i)
        If oBtn Is Nothing Then
            Exit For
        Else
            Select Case oBtn.BtnID
                Case "ID_Quote"
                    If frmMain.DockPro.State("frmQuotes") = DPHidden Then
                        oBtn.BtnState = eBtnState_Neutral
                    Else
                        oBtn.BtnState = eBtnState_Selected
                    End If
                    Set pB = frmMain.pbTbBack(oBtn.PicboxIndex)
                    If Not pB Is Nothing Then oBtn.Draw pB, aButtons
                Case "ID_SymbolGrid"
                    If frmMain.DockPro.State("frmSymbolGrid") = DPHidden Then
                        oBtn.BtnState = eBtnState_Neutral
                    Else
                        oBtn.BtnState = eBtnState_Selected
                    End If
                    Set pB = frmMain.pbTbBack(oBtn.PicboxIndex)
                    If Not pB Is Nothing Then oBtn.Draw pB, aButtons
                Case "ID_SectorBrowser"
                    If frmSectorTree.Visible Then
                        oBtn.BtnState = eBtnState_Selected
                    Else
                        oBtn.BtnState = eBtnState_Neutral
                    End If
                    Set pB = frmMain.pbTbBack(oBtn.PicboxIndex)
                    If Not pB Is Nothing Then oBtn.Draw pB, aButtons
                Case "ID_Snapshot"
                    If frmMain.DockPro.State("frmSnapshot") = DPHidden Then
                        oBtn.BtnState = eBtnState_Neutral
                    Else
                        oBtn.BtnState = eBtnState_Selected
                    End If
                    Set pB = frmMain.pbTbBack(oBtn.PicboxIndex)
                    If Not pB Is Nothing Then oBtn.Draw pB, aButtons
                Case "ID_ChartData"
                    If frmMain.DockPro.State("frmChartData") = DPHidden Then
                        oBtn.BtnState = eBtnState_Neutral
                    Else
                        oBtn.BtnState = eBtnState_Selected
                    End If
                    Set pB = frmMain.pbTbBack(oBtn.PicboxIndex)
                    If Not pB Is Nothing Then oBtn.Draw pB, aButtons
                Case "ID_ChartOnOff"
                    If frmMain.DockPro.State("frmChartOnOff") = DPHidden Then
                        oBtn.BtnState = eBtnState_Neutral
                    Else
                        oBtn.BtnState = eBtnState_Selected
                    End If
                    Set pB = frmMain.pbTbBack(oBtn.PicboxIndex)
                    If Not pB Is Nothing Then oBtn.Draw pB, aButtons
                Case "ID_PlanetData"
                    If frmMain.DockPro.State("frmPlanetData") = DPHidden Then
                        oBtn.BtnState = eBtnState_Neutral
                    Else
                        oBtn.BtnState = eBtnState_Selected
                    End If
                    Set pB = frmMain.pbTbBack(oBtn.PicboxIndex)
                    If Not pB Is Nothing Then oBtn.Draw pB, aButtons
'                Case "ID_TradeTracker"
'                    If frmMain.DockPro.State("frmTTSummary") = DPHidden Then
'                        oBtn.BtnState = eBtnState_Neutral
'                    Else
'                        oBtn.BtnState = eBtnState_Selected
'                    End If
'                    Set pB = frmMain.pbTbBack(oBtn.PicboxIndex)
                    If Not pB Is Nothing Then oBtn.Draw pB, aButtons
                Case "ID_Chain"
                    If frmMain.DockPro.State("frmOptionChain") = DPHidden Then
                        oBtn.BtnState = eBtnState_Neutral
                    Else
                        oBtn.BtnState = eBtnState_Selected
                    End If
                    Set pB = frmMain.pbTbBack(oBtn.PicboxIndex)
                    If Not pB Is Nothing Then oBtn.Draw pB, aButtons
                Case "ID_Orders"
                    If frmMain.DockPro.State("frmOrderTracker") = DPHidden Then
                        oBtn.BtnState = eBtnState_Neutral
                    Else
                        oBtn.BtnState = eBtnState_Selected
                    End If
                    Set pB = frmMain.pbTbBack(oBtn.PicboxIndex)
                    If Not pB Is Nothing Then oBtn.Draw pB, aButtons
            End Select
        End If
    Next
    
    If Not pB Is Nothing Then pB.Refresh


End Sub

Private Function PriceBarBtn(IndPrice As cIndicator, aIndTypeBtns As cGdArray, oBtn As cPicBoxButton) As cPicBoxButton
On Error Resume Next:

    Dim strID As String
    Dim oPriceBtn As cPicBoxButton

    If IndPrice Is Nothing Or aIndTypeBtns Is Nothing Or oBtn Is Nothing Then Exit Function

    strID = oBtn.BtnID
    
    Select Case IndPrice.DisplayType
        Case eINDIC_OHLC, eINDIC_HLC, eINDIC_HL
            If strID = "ID_OHLCBars" Then
                Set oPriceBtn = oBtn
            Else
                aIndTypeBtns.Add oBtn
            End If
        Case eINDIC_BollingerBar
            If strID = "ID_BollingerBars" Then
                Set oPriceBtn = oBtn
            Else
                aIndTypeBtns.Add oBtn
            End If
        Case eINDIC_Candlestick
            If strID = "ID_Candlesticks" Then
                Set oPriceBtn = oBtn
            Else
                aIndTypeBtns.Add oBtn
            End If
        Case eINDIC_Line
            If strID = "ID_CloseLine" Then
                Set oPriceBtn = oBtn
            Else
                aIndTypeBtns.Add oBtn
            End If
        Case eINDIC_Area
            If strID = "ID_Mountain" Then
                Set oPriceBtn = oBtn
            Else
                aIndTypeBtns.Add oBtn
            End If
        Case eINDIC_Renko
            If strID = "ID_Renko" Then
                Set oPriceBtn = oBtn
            Else
                aIndTypeBtns.Add oBtn
            End If
        Case eINDIC_Kagi
            If strID = "ID_Kagi" Then
                Set oPriceBtn = oBtn
            Else
                aIndTypeBtns.Add oBtn
            End If
        Case eINDIC_PNF
            If strID = "ID_PointFigure" Then
                Set oPriceBtn = oBtn
            Else
                aIndTypeBtns.Add oBtn
            End If
    End Select
    
    Set PriceBarBtn = oPriceBtn
    
End Function

Public Function ButtonByID(frm As Form, ByVal strID$, Optional ByVal strToolbar$ = "") As cPicBoxButton
On Error GoTo ErrSection:

    Dim i&
    Dim aToolArray As cGdArray
    Dim oButton As cPicBoxButton
    
    If frm Is Nothing Then Exit Function
    If Len(strID) = 0 Then Exit Function
    
    If Len(strToolbar) = 0 Then
        Set aToolArray = frm.TbButtonsArray(kTbGeneral)
        If Not aToolArray Is Nothing Then
            For i = 0 To aToolArray.Size - 1
                If aToolArray(i).BtnID = strID Then
                    Set oButton = aToolArray(i)
                    Exit For
                End If
            Next
        End If
        
        If oButton Is Nothing Then
            Set aToolArray = frm.TbButtonsArray(kTbDraw)
            If Not aToolArray Is Nothing Then
                For i = 0 To aToolArray.Size - 1
                    If aToolArray(i).BtnID = strID Then
                        Set oButton = aToolArray(i)
                        Exit For
                    End If
                Next
            End If
        End If
    Else
        Set aToolArray = frm.TbButtonsArray(strToolbar)
        If Not aToolArray Is Nothing Then
            For i = 0 To aToolArray.Size - 1
                If aToolArray(i).BtnID = strID Then
                    Set oButton = aToolArray(i)
                    Exit For
                End If
            Next
        End If
    End If
    
    Set ButtonByID = oButton
    
ErrExit:
    Set oButton = Nothing
    Set aToolArray = Nothing
    Exit Function

ErrSection:
    Set oButton = Nothing
    Set aToolArray = Nothing
    RaiseError "mPicBoxToolbar.ButtonByID"
    
End Function

Private Function ButtonByCoord(frm As Form, ByVal Index As Long, ByVal X As Single, ByVal Y As Single, _
    ByVal bDrawToolOnly As Boolean) As cPicBoxButton
On Error Resume Next:       'this routine called from mouse move events

    Dim i&
    
    Dim oButton As cPicBoxButton
    Dim aBtnsArray As cGdArray
    
    Dim bFound As Boolean
    
    If bDrawToolOnly Then
        Set aBtnsArray = frm.TbButtonsArray(kTbDraw)
    Else
        Set aBtnsArray = frm.TbButtonsArray("")
    End If

    For i = 0 To aBtnsArray.Size - 1
        Set oButton = aBtnsArray(i)
        If oButton.HitTest(Index, X, Y) Then
            bFound = True
            Exit For
        End If
    Next

    If Not bFound And Not bDrawToolOnly Then
        If g.vbeTbAlignDraw = vbAlignBottom Then
            Set aBtnsArray = frm.TbButtonsArray(kTbDraw)
            For i = 0 To aBtnsArray.Size - 1
                Set oButton = aBtnsArray(i)
                If oButton.HitTest(Index, X, Y) Then
                    bFound = True
                    Exit For
                End If
            Next
        End If
    End If
    
    If bFound Then Set ButtonByCoord = oButton

End Function

Private Function ToolbarSideBySide(frm As Form, pbTbBack As Variant, imgBack As Variant, _
    ByVal bWrap As Boolean) As Long
On Error GoTo ErrSection:

    Dim lpRect As Rect
    
    Dim i&, iEnd&
    Dim iRight&, iTop&, iLastVisibleIdx&
    Dim iPicboxIdx&, iBarPeriodPicBox&
    
    Dim bBarPeriodShow As Boolean
    Dim bDrawMoreButton As Boolean

    Dim aBtns As cGdArray
    Dim aBtnsDraw As cGdArray
    
    Dim oBtn As cPicBoxButton
    Dim oBtnPrev As cPicBoxButton
    Dim oBtnNext As cPicBoxButton
    
    Dim pB As PictureBox
    Dim imgToUse As ListImage
    
    
    Set aBtns = frm.TbButtonsArray("")
    If aBtns Is Nothing Then Exit Function
    
    Set aBtnsDraw = frm.TbButtonsArray(kTbDraw)
    If aBtnsDraw Is Nothing Then Exit Function
    
    If aBtnsDraw.Size <= 0 Then
        ToolbarResizeHorz frm, pbTbBack, imgBack, aBtns, bWrap, -1, -1
        Exit Function
    ElseIf aBtns.Size <= 0 Then
        Exit Function           'precautionary, theoretically should never happen
    End If
    
    Set oBtn = aBtns(0)
    Set pB = pbTbBack(0)
    'select & draw background skin
    Select Case g.eTbSkin
        Case eTbSkin_Unknown, eTbSkin_Silver, eTbSkin_AluminumSilver
            Set imgToUse = g.CoreBridge.ImageListBackgroundSkin.ListImages("kBkSilver")
        Case eTbSkin_Blue
            Set imgToUse = g.CoreBridge.ImageListBackgroundSkin.ListImages("kBkBlue")
        Case eTbSkin_ALuminumBlue
            Set imgToUse = g.CoreBridge.ImageListBackgroundSkin.ListImages("kBkAluminumBlue")
        Case eTbSkin_DarkFlat
            Set imgToUse = g.CoreBridge.ImageListBackgroundSkin.ListImages("kBkDarkFlat")
        Case eTbSkin_LightFlat
            Set imgToUse = g.CoreBridge.ImageListBackgroundSkin.ListImages("kBkLightFlat")
    End Select
        
    pB.AutoRedraw = True
    pB.Move 0, 0, frm.Width, (oBtn.BtnHeight + 2) * Screen.TwipsPerPixelY
    imgBack(0).Picture = Nothing
    imgBack(0).Move 0, 0, pB.Width, pB.Height
    
    geOverlayICO pB.hDC, imgToUse.ExtractIcon, 0, pB.Width / Screen.TwipsPerPixelX, oBtn.BtnHeight + 3, 0, 0, vbBlack, ""
        
    GetClientRect frm.hWnd, lpRect
    
    'reserve space for minimum of 5 drawing tool buttons
    If bWrap Then
        iEnd = lpRect.Right
    Else
        If aBtnsDraw.Size < 5 Then
            iEnd = lpRect.Right - ((aBtnsDraw(0).BtnWidth + 1) * aBtnsDraw.Size)
        Else
            iEnd = lpRect.Right - ((aBtnsDraw(0).BtnWidth + 1) * 5)
        End If
    End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''
'draw buttons from general toolbar
''''''''''''''''''''''''''''''''''''''''''''''''''''''
    iLastVisibleIdx = 0
    aBtns(aBtns.Size - 1).BtnDrawIndexReset
    
    Set oBtn = aBtns(0)
    oBtn.LastVisibleBtnIndex = 0
    oBtn.Draw pB, aBtns, 0
    If oBtn.BtnID = "ID_BarPeriod" Then bBarPeriodShow = True
       
    Set oBtnPrev = oBtn

    For i = 1 To aBtns.Size - 2
        Set oBtn = aBtns(i)
        Set oBtnNext = aBtns(i + 1)
        
        'get pixel of where right edge of this button will be
        iRight = oBtnPrev.BtnLeft + oBtnPrev.BtnWidth + oBtn.BtnWidth + 1
    
        If oBtn.BtnID = kMoreBtnID Then
            Exit For
        ElseIf iRight + oBtnNext.BtnWidth < iEnd Then
            oBtn.LastVisibleBtnIndex = i
            iLastVisibleIdx = i
            oBtn.Draw pB, aBtns, iPicboxIdx, , oBtnPrev
            If oBtn.BtnID = "ID_BarPeriod" Then
                bBarPeriodShow = True
                iBarPeriodPicBox = iPicboxIdx
            End If
            Set oBtnPrev = oBtn
        ElseIf bWrap Then
            iPicboxIdx = iPicboxIdx + 1
            If pbTbBack.UBound < iPicboxIdx Then Load pbTbBack(iPicboxIdx)
            If imgBack.UBound < iPicboxIdx Then Load imgBack(iPicboxIdx)
            Set imgBack(iPicboxIdx).Picture = Nothing
            pbTbBack(iPicboxIdx).AutoRedraw = True
            Set imgBack(iPicboxIdx).Container = pbTbBack(iPicboxIdx)
            iTop = pbTbBack(iPicboxIdx - 1).Top + pbTbBack(iPicboxIdx - 1).Height
            pbTbBack(iPicboxIdx).Move 0, iTop, pbTbBack(iPicboxIdx - 1).Width, pbTbBack(iPicboxIdx - 1).Height
            imgBack(iPicboxIdx).Stretch = True
            imgBack(iPicboxIdx).Move 0, 0, pbTbBack(iPicboxIdx).Width, pbTbBack(iPicboxIdx).Height
            imgBack(iPicboxIdx).Picture = Nothing
            If Not pbTbBack(iPicboxIdx).Visible Then pbTbBack(iPicboxIdx).Visible = True
            If Not imgBack(iPicboxIdx).Visible Then imgBack(iPicboxIdx).Visible = True
            
            Set pB = pbTbBack(iPicboxIdx)
            geOverlayICO pB.hDC, imgToUse.ExtractIcon, 0, pB.Width / Screen.TwipsPerPixelX, oBtn.BtnHeight + 2, 0, 0, vbBlack, ""
            
            oBtn.LastVisibleBtnIndex = i
            iLastVisibleIdx = i
            oBtn.Draw pB, aBtns, iPicboxIdx
            If oBtn.BtnID = "ID_BarPeriod" Then
                bBarPeriodShow = True
                iBarPeriodPicBox = iPicboxIdx
            End If
            Set oBtnPrev = oBtn
        Else
            bDrawMoreButton = True
            Exit For
        End If
    Next
    
    'reset draw info indexes
    For i = 0 To iLastVisibleIdx
        aBtns(i).LastVisibleBtnIndex = iLastVisibleIdx
    Next
    
    For i = iLastVisibleIdx + 1 To aBtns.Size - 1
        aBtns(i).LastVisibleBtnIndex = -1
        aBtns(i).BtnDrawIndexReset
    Next
    
    'draw the 'More' button if necessary
    If bDrawMoreButton Then
        Set oBtn = aBtns(aBtns.Size - 1)
        oBtn.LastVisibleBtnIndex = iLastVisibleIdx
        oBtn.Draw pB, aBtns, 0
        Set oBtnPrev = oBtn
    End If
    

''''''''''''''''''''''''''''''''''''''''''''''''''''''
'draw buttons from drawing toolbar
''''''''''''''''''''''''''''''''''''''''''''''''''''''
    iLastVisibleIdx = 0
    aBtnsDraw(aBtnsDraw.Size - 1).BtnDrawIndexReset
        
    'draw first button from drawing toolbar
    Set oBtn = aBtnsDraw(0)
    oBtn.LastVisibleBtnIndex = 0
    oBtn.Draw pB, aBtnsDraw, iPicboxIdx             '6254
    Set oBtnPrev = oBtn
        
    For i = 1 To aBtnsDraw.Size - 2
        Set oBtn = aBtnsDraw(i)
        Set oBtnNext = aBtnsDraw(i + 1)
        
        'get pixel of where right edge of this button will be
        iRight = oBtnPrev.BtnLeft + oBtnPrev.BtnWidth + oBtn.BtnWidth + 1
    
        oBtn.BtnDrawIndexReset
        If oBtn.BtnID = kMoreBtnID Then
            Exit For
        ElseIf (lpRect.Right - iRight) > oBtnNext.BtnWidth Then
            oBtn.LastVisibleBtnIndex = i
            iLastVisibleIdx = i
            oBtn.Draw pB, aBtnsDraw, iPicboxIdx, , oBtnPrev
            Set oBtnPrev = oBtn
        ElseIf bWrap Then
            iPicboxIdx = iPicboxIdx + 1
            If pbTbBack.UBound < iPicboxIdx Then Load pbTbBack(iPicboxIdx)
            If imgBack.UBound < iPicboxIdx Then Load imgBack(iPicboxIdx)
            Set imgBack(iPicboxIdx).Picture = Nothing
            pbTbBack(iPicboxIdx).AutoRedraw = True
            Set imgBack(iPicboxIdx).Container = pbTbBack(iPicboxIdx)
            
            iTop = pbTbBack(iPicboxIdx - 1).Top + pbTbBack(iPicboxIdx - 1).Height
            pbTbBack(iPicboxIdx).Move 0, iTop, pbTbBack(iPicboxIdx - 1).Width, pbTbBack(iPicboxIdx - 1).Height
            imgBack(iPicboxIdx).Stretch = True
            imgBack(iPicboxIdx).Move 0, 0, pbTbBack(iPicboxIdx).Width, pbTbBack(iPicboxIdx).Height
            imgBack(iPicboxIdx).Picture = Nothing
            
            If Not pbTbBack(iPicboxIdx).Visible Then pbTbBack(iPicboxIdx).Visible = True
            If Not imgBack(iPicboxIdx).Visible Then imgBack(iPicboxIdx).Visible = True
            
            Set pB = pbTbBack(iPicboxIdx)
            geOverlayICO pB.hDC, imgToUse.ExtractIcon, 0, pB.Width / Screen.TwipsPerPixelX, oBtn.BtnHeight + 2, 0, 0, vbBlack, ""
            
            oBtn.LastVisibleBtnIndex = i
            iLastVisibleIdx = i
            oBtn.Draw pB, aBtnsDraw, iPicboxIdx
            If oBtn.BtnID = "ID_BarPeriod" Then
                bBarPeriodShow = True
                iBarPeriodPicBox = iPicboxIdx
            End If
            Set oBtnPrev = oBtn
        Else
            bDrawMoreButton = True
            Exit For
        End If
    Next
    
    'reset draw info indexes
    For i = 0 To iLastVisibleIdx
        aBtnsDraw(i).LastVisibleBtnIndex = iLastVisibleIdx
    Next
    
    For i = iLastVisibleIdx + 1 To aBtnsDraw.Size - 1
        aBtnsDraw(i).LastVisibleBtnIndex = -1
        aBtnsDraw(i).BtnDrawIndexReset
    Next
    
    'draw the 'More' button if necessary
    If bDrawMoreButton Then
        Set oBtn = aBtnsDraw(aBtnsDraw.Size - 1)
        oBtn.LastVisibleBtnIndex = iLastVisibleIdx
        oBtn.Draw pB, aBtnsDraw, iPicboxIdx, , oBtnPrev
    End If
    
ErrExit:
''''''''''''''''''''''''''''''''''''''''''''''''''''''
'finish miscellaneous tasks
''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For i = iPicboxIdx + 1 To pbTbBack.UBound
        pbTbBack(i).Visible = False
        imgBack(i).Visible = False
    Next
    
    If bBarPeriodShow Then
        Set frm.cboBarPeriod.Container = pbTbBack(iBarPeriodPicBox)
        If frm.cboBarPeriod.ListCount = 0 Then BarPeriodCboInit frm.cboBarPeriod
        frm.cboBarPeriod.Enabled = True
        frm.cboBarPeriod.Visible = True
        frm.cboBarPeriod.ZOrder
    Else
        frm.cboBarPeriod.Enabled = False
        frm.cboBarPeriod.Visible = False
    End If
    
    For i = 0 To iPicboxIdx
        pbTbBack(i).Refresh
    Next

    Exit Function

ErrSection:
    RaiseError "mPicBoxToolbar.ToolbarSideBySide"

End Function

Public Function ToolbarMouseEvent(frm As Form, oLastMouseBtn As cPicBoxButton, ByVal MouseEvent&, ByVal Button&, _
    ByVal Index&, ByVal X As Single, ByVal Y As Single, ByVal bDrawToolOnly As Boolean, _
    Optional ByRef ButtonIn As cPicBoxButton = Nothing) As cPicBoxButton
On Error Resume Next:           'called from mouse events

    
    Dim strToolbar$
    
    Dim aBtnArray As cGdArray
    Dim oButton As cPicBoxButton
    Dim Chart As cChart
    
    Dim bExit As Boolean
    
    If frm Is Nothing Then Exit Function
    
    If ButtonIn Is Nothing Then
        Set oButton = ButtonByCoord(frm, Index, X, Y, bDrawToolOnly)
    Else
        Set oButton = ButtonIn
    End If
    
    If oButton Is Nothing Then Exit Function
    
    If oButton Is oLastMouseBtn And MouseEvent = WM_MOUSEMOVE Then
        Set ToolbarMouseEvent = oButton
        Exit Function
    End If
    
    strToolbar = oButton.ToolBarName
    
    If TypeOf frm Is frmMain Then
        If Not ActiveChart Is Nothing Then
            Set Chart = ActiveChart.Chart
            If ActiveChart.DetachStatus = eDetached Then
                If ActiveChart.Chart.ShowToolbar = 0 Then
                    If Not g.ChartGlobals.frmActiveNonDetached Is Nothing Then
                        SendMessage g.ChartGlobals.frmActiveNonDetached.hWnd, WM_NCACTIVATE, 0, 0
                    End If
                    SendMessage ActiveChart.hWnd, WM_NCACTIVATE, 1, 0
                ElseIf Not g.ChartGlobals.frmActiveNonDetached Is Nothing Then
                    SendMessage frmMain.hWnd, WM_NCACTIVATE, 1, 0
                    SendMessage frmMain.hWnd, WM_ACTIVATE, 1, 0
                    
                    ActiveChartFormSet g.ChartGlobals.frmActiveNonDetached
                    SendMessage g.ChartGlobals.frmActiveNonDetached.hWnd, WM_NCACTIVATE, 1, 0
                    SendMessage g.ChartGlobals.frmActiveNonDetached.hWnd, WM_MOUSEACTIVATE, 1, 0
                End If
            End If
        End If
    ElseIf TypeOf frm Is frmChart2 And Not frm Is ActiveChart Then
        ActiveChartFormSet frm
        SendMessage frm.hWnd, WM_NCACTIVATE, 1, 0
        SendMessage frm.hWnd, WM_MOUSEACTIVATE, 1, 0
        MoveFocus frm.pbChart
        DoEvents
    End If
    
    If MouseEvent = WM_LBUTTONDOWN Then
    
        If oButton.BtnID <> "ID_TextIncrease" And oButton.BtnID <> "ID_TextDecrease" Then
            TextIncDecUnregisterForm Nothing
        End If
    
        If Chart Is Nothing Then
            If IsFrmChart(frm) Then
                Set Chart = frm.Chart
            ElseIf Not ActiveChart Is Nothing Then
                Set Chart = ActiveChart.Chart
            End If
        End If
        
        If Not Chart Is Nothing Then
            If strToolbar = kTbDraw Then
                Select Case oButton.BtnID
                    Case "ID_ChartMove", "ID_ZoomIn", "ID_ZoomOut", "ID_Eraser", "ID_Magnet", "ID_RepeatDraw", _
                         "ID_DragModeY", "ID_Trendline", "ID_TrendChannel", "ID_SRLine", _
                         "ID_HorzLine", "ID_VertLine", "ID_ArrowLine", "ID_Text", "ID_Icon", "ID_ElliotLabels", "ID_ElliotEndUser", _
                         "ID_Bracket", "ID_Rectangle", "ID_Ellipse", "ID_MoreButtons"
                         
                         'these are ok, do nothing
                    
                    Case Else
                        If Chart.SeasonalCycleTypeEnum <> eCycleType_Undefined Then bExit = True
                End Select
            Else
                Select Case oButton.BtnID
                    
                    Case "ID_ChartOnOff", "ID_IndAnalyst", "ID_ChartOrderbar", "ID_Templates"
                        If Chart.SeasonalCycleTypeEnum <> eCycleType_Undefined Then bExit = True
                    
                    Case "ID_BarPeriod", "ID_Yearly", "ID_Quarterly", "ID_Monthly", "ID_Weekly", "ID_Daily", _
                         "ID_360minute", "ID_240minute", "ID_180minute", "ID_120minute", "ID_90minute", "ID_60minute", _
                         "ID_30minute", "ID_15minute", "ID_10minute", "ID_5minute", "ID_3minute", "ID_1minute", _
                         "ID_CustomMinute", "ID_CustomPeriod"
                        If Chart.SeasonalCycleTypeEnum <> eCycleType_Undefined Then bExit = True
    
                    Case "ID_OHLCBars", "ID_Candlesticks", "ID_BollingerBars", "ID_CloseLine", "ID_Mountain", _
                         "ID_PointFigure", "ID_Kagi", "ID_Renko", "ID_JPDaily", "ID_JPWeekly", "ID_JPMonthly", _
                         "ID_JPQuarterly", "ID_JPExpiration", "ID_Templates", "ID_WhatIf"
                        If Chart.SeasonalCycleTypeEnum <> eCycleType_Undefined Then bExit = True
                
                End Select
            End If
        End If
        
        If bExit Then
            InfBox kSeasonalUnavail, "I", "Ok", "Seasonal chart"
            Set ToolbarMouseEvent = oButton
            Exit Function
        End If
        
    End If
    
    If strToolbar = kTbDraw And Button = vbRightButton Then
        Dim Annot As cAnnotation
        
        If MouseEvent = WM_LBUTTONUP Then
            If TypeOf frm Is frmMain Then
                If Not ActiveChart Is Nothing Then
                    Set Chart = ActiveChart.Chart
                    If Not Chart Is Nothing Then
                        If ActiveChart.DetachStatus = eDetached And Chart.ShowToolbar Then
                            Set Chart = Nothing
                        End If
                    End If
                End If
            ElseIf IsFrmChart(frm) Then
                Set Chart = frm.Chart
            End If
        
            If Not Chart Is Nothing Then
                If oButton.BtnID = "ID_Hawkeye" Then
                    Chart.HandleHawkeyeButton -1
                Else
                    Set Annot = New cAnnotation
                    Chart.RemoveAnnots True, Annot.AnnotTypeFromToolID(oButton.BtnID)     '5045
                    Chart.SyncGlobalAnnots Nothing, True            '6335
                    Set Annot = Nothing
                End If
                Chart.GenerateChart eRedo1_Scrolled
                Set Chart = Nothing
            End If
        End If
    Else
        Set aBtnArray = frm.TbButtonsArray(strToolbar)
        
        If Not aBtnArray Is Nothing Then
            If TypeOf frm Is frmTbMoreButtons Then
                Select Case MouseEvent
                    Case WM_LBUTTONDOWN
                        Set ToolbarMouseEvent = oButton.MouseDown(frm, frm.pbTbBack(Index), aBtnArray, aBtnArray(aBtnArray.Size - 1))
                    Case WM_LBUTTONUP
                        Set ToolbarMouseEvent = oButton.MouseUp(frm, frm.pbTbBack(Index), aBtnArray, aBtnArray(aBtnArray.Size - 1))
                    Case WM_MOUSEMOVE
                        Set ToolbarMouseEvent = oButton.MouseMove(frm, frm.pbTbBack(Index), aBtnArray, aBtnArray(aBtnArray.Size - 1))
                End Select
            ElseIf strToolbar = kTbDraw And (g.vbeTbAlignDraw <> vbAlignBottom Or Not frm.pbTbBack(0).Visible) Then
                Select Case MouseEvent
                    Case WM_LBUTTONDOWN
                        Set ToolbarMouseEvent = oButton.MouseDown(frm, frm.pbTbBackDraw(Index), aBtnArray, aBtnArray(aBtnArray.Size - 1))
                    Case WM_LBUTTONUP
                        Set ToolbarMouseEvent = oButton.MouseUp(frm, frm.pbTbBackDraw(Index), aBtnArray, aBtnArray(aBtnArray.Size - 1))
                    Case WM_MOUSEMOVE
                        Set ToolbarMouseEvent = oButton.MouseMove(frm, frm.pbTbBackDraw(Index), aBtnArray, aBtnArray(aBtnArray.Size - 1))
                End Select
            Else
                Select Case MouseEvent
                    Case WM_LBUTTONDOWN
                        Set ToolbarMouseEvent = oButton.MouseDown(frm, frm.pbTbBack(Index), aBtnArray, aBtnArray(aBtnArray.Size - 1))
                    Case WM_LBUTTONUP
                        If oButton.BtnID = "ID_PatternProfit" And Button = vbRightButton Then
                            Set ToolbarMouseEvent = oButton.MouseUp(frm, frm.pbTbBack(Index), aBtnArray, aBtnArray(aBtnArray.Size - 1), True)
                        Else
                            Set ToolbarMouseEvent = oButton.MouseUp(frm, frm.pbTbBack(Index), aBtnArray, aBtnArray(aBtnArray.Size - 1))
                        End If
                    Case WM_MOUSEMOVE
                        Set ToolbarMouseEvent = oButton.MouseMove(frm, frm.pbTbBack(Index), aBtnArray, aBtnArray(aBtnArray.Size - 1))
                End Select
            End If
        End If
    
    End If
    
    Set m.oBtnLastProcessed = oButton       'save last button that received explicit mouse event from user's action on any toolbar

End Function

Public Function ClearLastMouseButton(frm As Form, oButton As cPicBoxButton) As Boolean
On Error Resume Next:       'this is called from mousemove event

    Dim i&, X&, Y&
    Dim pB As PictureBox
    Dim aBtnsArray As cGdArray
    
    Dim bCleared As Boolean
    
    If frm Is Nothing Or oButton Is Nothing Then Exit Function
    
    Set aBtnsArray = frm.TbButtonsArray(oButton.ToolBarName)
    If aBtnsArray Is Nothing Then Exit Function
    
    If oButton.BtnID = "ID_CustomizeToolbar" Then
        frmMain.ToolBarBtnSizeGet oButton.ToolBarName, X, Y
        If oButton.BtnWidth <> X Then       '5858
            Set oButton = Nothing
            Exit Function
        End If
    End If

    If oButton.BtnStatus <> kBtnStateDown And oButton.BtnState <> eBtnState_InProg Then         '5171
        If TypeOf frm Is frmTbMoreButtons Then
            Set pB = frm.pbTbBack(oButton.PicboxIndex)
            oButton.BtnClearNow pB, aBtnsArray
            Set oButton = Nothing
            bCleared = True
        ElseIf g.vbeTbAlignDraw = vbAlignBottom And frm.pbTbBack(0).Visible Then
            i = oButton.PicboxIndex
            If i >= 0 And i <= frm.pbTbBack.UBound Then
                Set pB = frm.pbTbBack(i)
                oButton.BtnClearNow pB, aBtnsArray
                Set oButton = Nothing
                bCleared = True
            End If
        ElseIf oButton.ToolBarName = kTbDraw Then
            i = oButton.PicboxIndex
            If i >= 0 And i <= frm.pbTbBackDraw.UBound Then
                Set pB = frm.pbTbBackDraw(i)
                oButton.BtnClearNow pB, aBtnsArray
                Set oButton = Nothing
                bCleared = True
            End If
        Else
            i = oButton.PicboxIndex
            If i >= 0 And i <= frm.pbTbBack.UBound Then
                Set pB = frm.pbTbBack(i)
                oButton.BtnClearNow pB, aBtnsArray
                Set oButton = Nothing
                bCleared = True
            End If
        End If
    End If
    
    ClearLastMouseButton = bCleared

End Function

Public Sub SyncTradeTrackerBtn(frm As Form)
On Error Resume Next:

    Dim aBtnArray As cGdArray
    Dim oBtn As cPicBoxButton, i&

    Dim pB As PictureBox
    Dim pbContainer As Variant

    Set aBtnArray = frm.TbButtonsArray("")
    If aBtnArray Is Nothing Then Exit Sub
    If aBtnArray.Size = 0 Then Exit Sub
    
    Set pbContainer = frm.pbTbBack
    If pbContainer Is Nothing Then Exit Sub
    
    Set oBtn = ButtonByID(frm, "ID_TradeTracker")
    If Not oBtn Is Nothing Then
        If oBtn.BtnDrawIndex >= 0 Then
            i = oBtn.PicboxIndex
            If i >= 0 And i <= pbContainer.UBound Then
                Set pB = pbContainer(i)
                If Not pB Is Nothing Then
                    oBtn.Draw pB, aBtnArray
                    pB.Refresh
                End If
            End If
        End If
    End If

End Sub

Public Sub TextIncDecUnregisterForm(frm As Form)
On Error Resume Next

    Dim oTextInc As cPicBoxButton
    Dim oTextDec As cPicBoxButton
    
    Dim eState As eBtnState
    
    Set oTextInc = ButtonByID(frmMain, "ID_TextIncrease")
    Set oTextDec = ButtonByID(frmMain, "ID_TextDecrease")
    
    If oTextInc Is Nothing And oTextInc Is Nothing Then
        Set m.frmTextIncDec = Nothing
        m.bFuncOverride = False
        Exit Sub
    End If
    
    If m.frmTextIncDec Is Nothing Then
        eState = eBtnState_Disabled
    ElseIf frm Is Nothing Then
        eState = eBtnState_Disabled
        Set m.frmTextIncDec = Nothing
        m.bFuncOverride = False
    ElseIf m.frmTextIncDec Is frm Then
        If TypeOf ActiveForm Is frmChart And Not TypeOf frm Is frmChart Then
            'do nothing ... the chart MDI child form will get activated when the main App toolbar is clicked
            '   do not want to switch out text inc/dec form here as user could be clicking the +T/-T multiple times
            '   the chart's form will explicitly register itself on a mouse down event
        Else
            Set m.frmTextIncDec = Nothing
            m.bFuncOverride = False
            eState = eBtnState_Disabled
        End If
    End If

    If eState = eBtnState_Disabled Then
        'no need to change button state unless form that +T/-T apply to has changed
        If Not oTextInc Is Nothing Then
            oTextInc.BtnState = eBtnState_Disabled
            oTextInc.BtnClearNow frmMain.pbTbBack(oTextInc.PicboxIndex), frmMain.TbButtonsArray(kTbGeneral)
        End If
            
        If Not oTextDec Is Nothing Then
            oTextDec.BtnState = eBtnState_Disabled
            oTextDec.BtnClearNow frmMain.pbTbBack(oTextInc.PicboxIndex), frmMain.TbButtonsArray(kTbGeneral)
        End If
    End If

End Sub

Public Sub TextIncDecRegisterForm(ByRef frm As Form, _
    Optional ByVal bFuncOverride As Boolean = False)
On Error Resume Next

    Dim oTextInc As cPicBoxButton
    Dim oTextDec As cPicBoxButton
    
    Dim eState As eBtnState
    
    Set m.frmTextIncDec = Nothing
    m.bFuncOverride = False
    
    Set oTextInc = ButtonByID(frmMain, "ID_TextIncrease")
    Set oTextDec = ButtonByID(frmMain, "ID_TextDecrease")
    
    If oTextInc Is Nothing And oTextInc Is Nothing Then
        Exit Sub
    ElseIf frm Is Nothing Then
        eState = eBtnState_Disabled
    Else
        eState = eBtnState_Neutral
        Set m.frmTextIncDec = frm
        m.bFuncOverride = bFuncOverride
    End If
    
    If Not oTextInc Is Nothing Then
        oTextInc.BtnState = eState
        oTextInc.BtnClearNow frmMain.pbTbBack(oTextInc.PicboxIndex), frmMain.TbButtonsArray(kTbGeneral)
    End If
    
    If Not oTextDec Is Nothing Then
        oTextDec.BtnState = eState
        oTextDec.BtnClearNow frmMain.pbTbBack(oTextDec.PicboxIndex), frmMain.TbButtonsArray(kTbGeneral)
    End If

End Sub

Public Sub GridTextIncrease()
On Error GoTo ErrSection:
    
    Dim i&
    Dim fg As VSFlexGrid
    
    If m.frmTextIncDec Is Nothing Then
        If TypeOf ActiveForm Is frmChart2 And ActiveForm Is ActiveChart Then
            ActiveChart.GridTextIncrease        'user is clicking +T on detached chart's toolbar
        End If
        Exit Sub
    End If
    
    If m.bFuncOverride Then
        m.frmTextIncDec.GridTextIncrease
    Else
        For i = 0 To m.frmTextIncDec.Controls.Count - 1
            If TypeOf m.frmTextIncDec.Controls(i) Is VSFlexGrid Then
                Set fg = m.frmTextIncDec.Controls(i)
                If fg.Visible Then fg.Font.Size = fg.Font.Size + 1
            End If
        Next
    End If
    
    Set fg = Nothing
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mPicBoxToolbar.GridTextIncrease"
    
End Sub

Public Sub GridTextDecrease()
On Error GoTo ErrSection:
    
    Dim i&
    Dim fg As VSFlexGrid
    
    If m.frmTextIncDec Is Nothing Then
        If TypeOf ActiveForm Is frmChart2 And ActiveForm Is ActiveChart Then
            ActiveChart.GridTextDecrease        'user is clicking -T on detached chart's toolbar
        End If
        Exit Sub
    End If
    
    If m.bFuncOverride Then
        m.frmTextIncDec.GridTextDecrease
    Else
        For i = 0 To m.frmTextIncDec.Controls.Count - 1
            If TypeOf m.frmTextIncDec.Controls(i) Is VSFlexGrid Then
                Set fg = m.frmTextIncDec.Controls(i)
                If fg.Visible Then fg.Font.Size = fg.Font.Size - 1
            End If
        Next
    End If
    
    Set fg = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mPicBoxToolbar.GridTextIncrease"
    
End Sub

