VERSION 5.00
Object = "{54F3CD43-5ADA-11D2-81EB-006008A2E49D}#1.0#0"; "Pego32a.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Object = "{CDFDF7C2-5B6A-11D2-81EB-006008A2E49D}#1.0#0"; "Pe3do32a.ocx"
Begin VB.Form frmOptChart 
   Caption         =   "Optimizer Chart"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   6660
   Begin PEGOALib.Pegoa Pegoa1 
      Height          =   1350
      Left            =   3240
      TabIndex        =   1
      Top             =   2205
      Width           =   2265
      _Version        =   65536
      _ExtentX        =   3995
      _ExtentY        =   2381
      _StockProps     =   96
      _AllProps       =   "frmOptChart.frx":0000
   End
   Begin PE3DOALib.Pe3doa Pe3doa1 
      Height          =   1500
      Left            =   705
      TabIndex        =   0
      Top             =   555
      Width           =   1905
      _Version        =   65536
      _ExtentX        =   3360
      _ExtentY        =   2646
      _StockProps     =   96
      _AllProps       =   "frmOptChart.frx":1B98
   End
   Begin ActiveToolBars.SSActiveToolBars tbToolbar 
      Left            =   150
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131083
      ToolBarsCount   =   1
      ToolsCount      =   5
      DisplayContextMenu=   0   'False
      Tools           =   "frmOptChart.frx":2E57
      ToolBars        =   "frmOptChart.frx":3C58
   End
End
Attribute VB_Name = "frmOptChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmOptChart.frm
'' Description: Form for showing an optimization chart
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 06/16/2011   DAJ         Added code for the Highlight Bar Reporter
'' 06/22/2011   DAJ         Changed "# Days" to "# Bars"
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kTeal = 12632064      'RGB(0,192,192)
Private Const kMaxLabels = 10
Private Const kDefaultIdxY = 17     'net profit

Private Type mPrivate
    fgData As VSFlexGrid
    obOptimizer As cOptimizer
    nIdxY As Long
    nIdxX As Long
    nIdxZ As Long
    nIdxVarStart As Long    'index of where variable "I-" or "R-" column begins
    nVarColCount As Long    'number of variable columns ("I-", "R-")
    nSubsets As Long
    'properties for 2D chart
    nPointsToGraph As Long
    n2dColor As Long
    n2dPlot As Long
    n2dDataShadow As Long
    'properties for 3D chart
    n3dPolyMode As Long     'chart type (bar or area)
    n3dPlotStyle As Long    'plot method (wire frame, surface, surface shade, contour)
    n3dViewHeight As Long
    n3dRotationDegree As Long
    nSingleColor As Long    '0/1 use only one color for 3d bar chart
    nLegLocation As Long    'legend location for 3d area chart
    nAllowArea As Long      '0/1 disallow 3d area chart for rules ("R-") variables
    nRotateDetail As Long   '0=wire,1=plot,2=full
    aColors As New cGdArray
    'common properties
    strTitle As String
    nFontSize As Long
    nGrid As Long
    'flags
    bAutoScaleData As Boolean
    bSettingsRead As Boolean
End Type
Private m As mPrivate

Public Sub ShowMe(fgSource As VSFlexGrid, objOpt As cOptimizer, strSystem As String, strSymbol As String)
On Error GoTo ErrSection:

    Set m.fgData = fgSource
    Set m.obOptimizer = objOpt
    
    SetVarColCount
    If m.nVarColCount > 2 Then
        MsgBox "Cannot chart more than 2 variable inputs."
        Unload Me
        Exit Sub
    End If
    
    SetTitle strSystem, strSymbol
    Me.Show

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChart.ShowMe"

End Sub

Private Sub SetTitle(ByVal strSystem$, ByVal strSymbol$)
On Error GoTo ErrSection:

    Dim strSym$
    
    strSym = Trim(Parse(strSymbol, ":", 2))
    m.strTitle = strSystem + " (" + strSym + ")"
    Me.Caption = strSym

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChart.SetTitle"

End Sub

Private Sub SetVarColCount()
On Error GoTo ErrSection:

    Dim i&
    
    With m.fgData
        m.nVarColCount = 0
        m.nIdxVarStart = -1
        For i = 0 To .Cols - 1
            If OptimizeColumn(i) Then
                If m.nIdxVarStart < 0 Then
                    m.nIdxVarStart = i   'aardvark 1606
                End If
                
                m.nVarColCount = m.nVarColCount + 1
            End If
        Next
    End With
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChart.SetVarColCount"

End Sub

Private Sub SetFirstVisibleCol()
On Error GoTo ErrSection:

    Dim i&
    
    'find first visible column on grid
    For i = 1 To m.fgData.Cols - 1
        If m.fgData.ColHidden(i) = False Then
            m.nIdxY = i
            Exit For
        End If
    Next
    
    'double check
    If m.nIdxY >= m.fgData.Cols Then
        m.nIdxY = 1
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChart.SetFirstVisibleCol"

End Sub

Private Sub InitPegoa1()
On Error GoTo ErrSection:

    Pegoa1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
        
    '** Set initial default values **'
    m.nPointsToGraph = -1
    m.n2dColor = kTeal
    m.n2dPlot = GPM_BAR
    m.n2dDataShadow = PEDS_NONE

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChart.InitPegoa1"

End Sub

Private Sub InitPe3doa1()
On Error GoTo ErrSection:

    Pe3doa1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
            
    '** Set initial default values **'
    m.n3dPolyMode = PEPM_3DBAR
    m.n3dPlotStyle = -1
    m.n3dViewHeight = -1
    m.n3dRotationDegree = -1
    m.nLegLocation = 0          'left
    m.nSingleColor = 0
    m.nRotateDetail = 2         'full detail

    '** Reference note from Gigasoft's sample project **'
    '** Subset colors work a little differently for 3d **'
    '** (0) will be wireframe color for wireframe plotting method **'
    '** (0) will be solid color if plotting method is solid
    '** (1) will be wire frame color for solid plotting methods
    
    '** contour/subset colors (from Gigasoft's sample project) **'
    m.aColors(0) = kTeal
    m.aColors(1) = RGB(128, 0, 0)    'rusty red
    m.aColors(2) = RGB(0, 0, 128)    'marine blue
    m.aColors(3) = RGB(128, 128, 0)  'olive green
    m.aColors(4) = RGB(0, 128, 128)  'dark teal
    m.aColors(5) = RGB(0, 128, 0)    'forest green
    m.aColors(6) = RGB(255, 0, 0)    'bright red
    m.aColors(7) = RGB(0, 0, 255)    'bright blue
    m.aColors(8) = RGB(255, 255, 0)  'bright yellow
    m.aColors(9) = RGB(0, 255, 255)  'bright light blue

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChart.InitPe3doa1"

End Sub

Private Sub ShowPegoa1()
On Error GoTo ErrSection:
    
    Dim strAxisY$, strAxisX$, strSubTitle$
    
    Pegoa1.Visible = True
    Pe3doa1.Visible = False
    tbToolbar.Tools("ID_Rotate").Visible = False

    If m.bSettingsRead = False Then ReadSettings True
    If m.nIdxY < 1 Then SetFirstVisibleCol
    m.nIdxX = m.nIdxVarStart
    
    'get text to use for subtitle and axes labels
    strSubTitle = SubTitle
    strAxisY = AxisLabelY
    strAxisX = AxisLabelHorz(m.nIdxX)
    
    'set properties not dependent on data information
    With Pegoa1
        .PEactions = 20
        '** properties not intended to be changed by user **'
        .PrepareImages = True
        .FocalRect = False
        .NullDataValue = kNullData
        .Subsets = 1
        .SubsetPointTypes(0) = PEPT_DOTSOLID
        .AllowPopup = False
        .AllowCustomization = False
        .AllowJpegOutput = True
        '** properties that can be changed by user **'
        .MainTitle = ""
        .SubTitle = ""
        .MultiSubTitles(0) = "|" + m.strTitle + "|"
        .MultiSubTitles(1) = "|" + strSubTitle + "|"
        .YAxisLabel = strAxisY
        .XAxisLabel = strAxisX
        .PlottingMethod = m.n2dPlot
        .GridLineControl = m.nGrid
        .SubsetColors(0) = m.n2dColor
        .FontSize = m.nFontSize
        .DataShadows = m.n2dDataShadow
    End With
        
    '** set data **'
    SetDataPegoa1

    'set properties that are dependent on data information
    With Pegoa1
        .AutoScaleData = m.bAutoScaleData
        If m.nPointsToGraph < 1 Then
            .PointsToGraph = Pegoa1.Points
        Else
            .PointsToGraph = m.nPointsToGraph
        End If
        '** set scroll position if necessary **'
            'scroll position at first point = 1
            'scroll position at last point = datapoints - pointsToGraph + 1;
        If .Points > .PointsToGraph Then .HorzScrollPos = .Points - .PointsToGraph + 1
        '** Always call PEactions = 0 at end **'
        .PEactions = 0
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChart.ShowPegoa1"

End Sub

'Reference Info from Gigasoft Help
'POLYMODE:
'PEPM_SURFACEPOLYGONS (1)    X,Y,Z data is processed to form a surface of polygons
'PEPM_3DBAR (2)  X,Y,Z data is processed to form multiple 3d bars via polygons
'PEPM_POLYGONDATA (3)    Polygon data is provided via PEP_structPOLYDATA
'PEPM_SCATTER (4)    X,Y,Z data is plotted as in a 2D scatter chart
'PLOTTING METHOD:
'When PEP_nPOLYMODE is PEPM_SURFACEPOLYGONS, PEPM_3DBAR, and PEPM_POLYGONDATA
'0 = WireFrame.
'1 = Surface.
'2 = Surface with Shading.
'3 = Surface with Pixels.
'4 = Surface with Contours.  Not available for PEPM_3DBAR and PEPM_POLYGONDATA
Private Sub ShowPe3doa1()
On Error GoTo ErrSection:
    
    Dim strAxisY$, strAxisX$, strAxisZ$, strSubTitle$
    Dim i&
         
    Pe3doa1.Visible = True
    Pegoa1.Visible = False
    tbToolbar.Tools("ID_Rotate").Visible = True
    
    If Not m.bSettingsRead Then ReadSettings False
    If m.nIdxY < 1 Then SetFirstVisibleCol
                
    If m.nIdxX < 1 Then
        m.nIdxX = m.nIdxVarStart
        m.nIdxZ = m.nIdxX + 1
    End If
        
    'get text for sub title & axes labels
    strSubTitle = SubTitle
    strAxisY = AxisLabelY
    strAxisX = AxisLabelHorz(m.nIdxX)
    strAxisZ = AxisLabelHorz(m.nIdxZ)
    
    'set properties not dependent on data information
    With Pe3doa1
        .PEactions = 20
        '** properties not intended to be changed by user **'
        .PrepareImages = True
        .FocalRect = False
        .NullDataValueX = kNullData
        .NullDataValueY = kNullData
        .NullDataValueZ = kNullData
        .AllowPopup = False
        .AllowCustomization = False
        .ShadingStyle = PESS_COLORSHADING
        .AllowJpegOutput = True
        .GridLineControl = m.nGrid
        '** properties that can be changed by user **'
        .MainTitle = ""
        .SubTitle = ""
        .MultiSubTitles(0) = "|" + m.strTitle + "|"
        .MultiSubTitles(1) = "|" + strSubTitle + "|"
        .YAxisLabel = strAxisY
        .XAxisLabel = strAxisX
        .ZAxisLabel = strAxisZ
        .PolyMode = m.n3dPolyMode
        .FontSize = m.nFontSize
        .GridLineControl = m.nGrid
        .RotationDetail = m.nRotateDetail
        '** legend location **'
        If m.nLegLocation = -1 Then
            .ShowContourLegends = False
        Else
            .ShowContourLegends = True
            .LegendLocation = m.nLegLocation
        End If
        '** plotting style **'
        If m.n3dPlotStyle < 0 Then
            If .PolyMode = PEPM_3DBAR Then
                .PlottingMethod = TDPM_2     'surface with shading
            ElseIf .PolyMode = PEPM_SURFACEPOLYGONS Then
                .PlottingMethod = TDPM_4     'surface with contour
            End If
            m.n3dPlotStyle = .PlottingMethod
        Else
            .PlottingMethod = m.n3dPlotStyle
        End If
        '** camera position **'
        If m.n3dViewHeight < 0 Then
            .ViewingHeight = 15
        Else
            .ViewingHeight = m.n3dViewHeight
        End If
        If m.n3dRotationDegree < 0 Then
            m.n3dRotationDegree = 1
            If .PolyMode = PEPM_3DBAR Then
                .DegreeOfRotation = 314
            ElseIf .PolyMode = PEPM_SURFACEPOLYGONS Then
                .DegreeOfRotation = 64
            End If
        Else
            .DegreeOfRotation = m.n3dRotationDegree
        End If
    End With
    
    '** set data **'
    SetDataPe3doa1
    
    'set properties dependent on data information
    With Pe3doa1
        .AutoScaleData = m.bAutoScaleData
        '** colors **'
        If .PolyMode = PEPM_3DBAR And m.nSingleColor = 1 Then
            For i = 0 To m.nSubsets
                .SubsetColors(i) = m.aColors(0)
            Next
        Else
            For i = 0 To m.nSubsets
                .SubsetColors(i) = m.aColors(i Mod 10)
            Next
        End If
        .PEactions = 0
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChart.ShowPe3doa1"

End Sub

Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strText$

    m.nIdxY = kDefaultIdxY
    m.nIdxX = -1
    m.nIdxZ = -1
        
    m.nFontSize = PEFS_MEDIUM
    m.nGrid = PEGLC_NONE
    
    InitPegoa1
    InitPe3doa1
                  
    If m.nVarColCount = 1 Then
        ShowPegoa1
    Else
        ShowPe3doa1
    End If
    
    With tbToolbar
        .Tools("ID_Cancel").Picture = Picture16(ToolbarIcon("kCancel"))
        .Tools("ID_Settings").Picture = Picture16(ToolbarIcon("ID_Settings"))   'want new toolbar to use kSettings for consistency
        .Tools("ID_Print").Picture = Picture16(ToolbarIcon("ID_Print"))
        .Tools("ID_Export").Picture = Picture16(ToolbarIcon("kSave"))
    End With
    
    Me.Icon = Picture16(ToolbarIcon("kBarChart"))
    
    'Restore/set form size & location
    strText = GetIniFileProperty(Me.Name, "", "Placement", g.strIniFile)
    If strText = "" Then
        'For forms to fit screen using 800x600 (small fonts)
        '- Max Height:  6960 (with a Status Bar),  7320 (without Status Bar)
        '- Max Width:  11880
        Me.Height = 7000
        Me.Width = 10000
        CenterTheForm Me
    Else
        SetFormPlacement Me, strText
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChart.Form_Load"

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:
    
    WriteSettings
    
    Set m.fgData = Nothing
    Set m.obOptimizer = Nothing
    Set m.aColors = Nothing
    
    'save form size & location
    SetIniFileProperty Me.Name, GetFormPlacement(Me), "Placement", g.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChart.Form_Unload"

End Sub

Private Sub Form_Resize()
On Error Resume Next

    If Pegoa1.Visible = True Then
        Pegoa1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
        Pegoa1.PEactions = 3
    ElseIf Pe3doa1.Visible = True Then
        Pe3doa1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
        Pe3doa1.PEactions = 3
    End If

End Sub

Private Sub SetDataPegoa1()
On Error GoTo ErrSection:
        
    Dim i&, dY#, dMax#
    
    Dim aY As New cGdArray, _
        aX As New cGdArray, _
        aIdx As cGdArray
        
    Dim tblData As New cGdTable
    
    Dim bBooleanX As Boolean, _
        bPercentY As Boolean
                
    dMax = kNullData
               
    If InStr(m.fgData.TextMatrix(0, m.nIdxY), "%") Then bPercentY = True
                
    For i = m.fgData.FixedRows To m.fgData.Rows - 1
        dY = ValOfText(m.fgData.TextMatrix(i, m.nIdxY))
        If bPercentY Then dY = dY * 100
        If dMax < dY Then dMax = dY
        
        aY.Add dY
        aX.Add ValOfText(m.fgData.TextMatrix(i, m.nIdxX))
    Next
   
    tblData.AttachField aY
    tblData.AttachField aX
   
    Set aIdx = tblData.CreateIndex
    tblData.SortIndex aIdx, 1, eGdSort_Default
       
    '** Set Subsets, Points, and fill with YData **'
    Pegoa1.Points = m.fgData.Rows - m.fgData.FixedRows
    
    'determine whether axis label need to be set as "Y" and "N"
    If Left(m.fgData.TextMatrix(0, m.nIdxX), 2) = "R-" Then bBooleanX = True
    
    For i = 0 To aIdx.Size - 1
        Pegoa1.YData(0, i) = tblData.Item(0, aIdx(i))
        If bBooleanX = True Then
            If tblData.Item(1, aIdx(i)) = 0 Then
                Pegoa1.PointLabels(i) = "N"
            Else
                Pegoa1.PointLabels(i) = "Y"
            End If
        Else
            Pegoa1.PointLabels(i) = CStr(tblData.Item(1, aIdx(i)))
        End If
    Next
    
    '** this controls formatting of y-scale labels **'
    '** false: entire numbers are shown **'
    '** true: Gigasoft appends K for thousands, M for millions etc.
    '**       e.g. 100K, 10M
    If dMax < 1000 Then
        m.bAutoScaleData = False
    Else
        m.bAutoScaleData = True
    End If
   
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChart.SetDataPegoa1"

End Sub

Private Sub SetDataPe3doa1()
On Error GoTo ErrSection:
    
    Dim i&, j&, k&
    Dim nPoints&, nSubsets&     'subsets=Z-axis, points=X-axis (Gigasoft's nomenclature)
    Dim nPtIdx&                 'index for X-axis labels
    Dim nFreqSubLabel, nFreqPtLabel     'how far apart labels are set on X,Z scale
    
    Dim bBooleanX As Boolean, _
        bBooleanZ As Boolean, _
        bPercentY As Boolean, _
        bDoLabel As Boolean
               
    Dim dY#, dMin#, dMax#
       
    Dim aY As New cGdArray, _
        aX As New cGdArray, _
        aZ As New cGdArray, _
        aIdx As cGdArray
        
    Dim tblData As New cGdTable
                    
    If InStr(m.fgData.TextMatrix(0, m.nIdxY), "%") Then bPercentY = True
                    
    For i = m.fgData.FixedRows To m.fgData.Rows - 1
        dY = ValOfText(m.fgData.TextMatrix(i, m.nIdxY))
        If bPercentY Then dY = dY * 100
        
        aY.Add dY
        aX.Add ValOfText(m.fgData.TextMatrix(i, m.nIdxX))
        aZ.Add ValOfText(m.fgData.TextMatrix(i, m.nIdxZ))
    Next
   
    tblData.AttachField aY
    tblData.AttachField aX
    tblData.AttachField aZ
   
    'sort first on column for X data
    Set aIdx = tblData.CreateIndex
    tblData.SortIndex aIdx, 1
    
    'count number of points (x-axis)
    nPoints = 1
    For i = 1 To aIdx.Size - 1
        If tblData(1, aIdx(i)) <> tblData(1, aIdx(i - 1)) Then
            nPoints = nPoints + 1
        End If
    Next
    
    'stable sort on column for Z data
    tblData.SortIndex aIdx, 2, eGdSort_Stable
    
    'count number of subsets (points along z-axis)
    nSubsets = 1
    For i = 1 To aIdx.Size - 1
        If tblData(2, aIdx(i)) <> tblData(2, aIdx(i - 1)) Then
            nSubsets = nSubsets + 1
        End If
    Next
    
    'double check
    If nPoints * nSubsets <> aIdx.Size Then Exit Sub
       
    Pe3doa1.Subsets = nSubsets
    Pe3doa1.Points = nPoints
               
    'determine whether axis label need to be set as "Y" and "N"
    m.nAllowArea = 1
    If Left(m.fgData.TextMatrix(0, m.nIdxX), 2) = "R-" Then bBooleanX = True
    If Left(m.fgData.TextMatrix(0, m.nIdxZ), 2) = "R-" Then bBooleanZ = True
    If bBooleanX Or bBooleanZ Then m.nAllowArea = 0
                 
    'set frequency of subset and point labels
    nFreqSubLabel = 1
    nFreqPtLabel = 1
    If nSubsets > kMaxLabels Then nFreqSubLabel = nSubsets / kMaxLabels
    If nPoints > kMaxLabels Then nFreqPtLabel = nPoints / kMaxLabels
    
    'initialize counters
    k = 0
    nPtIdx = 0
    dMax = kNullData
    dMin = -1 * dMax
    
    For i = 0 To nSubsets - 1
        'set subset labels
        bDoLabel = True
        If nSubsets > kMaxLabels Then
            If i Mod nFreqSubLabel <> 0 Then bDoLabel = False
        End If
        If bDoLabel Then
            If bBooleanZ = True Then
                If tblData(2, aIdx(k)) = 0 Then
                    Pe3doa1.SubsetLabels(i) = "N"
                Else
                    Pe3doa1.SubsetLabels(i) = "Y"
                End If
            Else
                Pe3doa1.SubsetLabels(i) = CStr(tblData(2, aIdx(k))) 'label for Z-Axis
            End If
        Else
            Pe3doa1.SubsetLabels(i) = ""
        End If
        For j = 0 To nPoints - 1
            'set data values
            dY = tblData(0, aIdx(k))
            Pe3doa1.YData(i, j) = dY
            Pe3doa1.XData(i, j) = tblData(1, aIdx(k))
            Pe3doa1.ZData(i, j) = tblData(2, aIdx(k))
            'set point labels
            bDoLabel = True
            If nPoints > kMaxLabels Then
                If j Mod nFreqPtLabel <> 0 Then bDoLabel = False
            End If
            If bDoLabel = True Then
                If nPtIdx < nPoints Then
                    If bBooleanX = True Then
                        If tblData(1, aIdx(k)) = 0 Then
                            Pe3doa1.PointLabels(nPtIdx) = "N"
                        Else
                            Pe3doa1.PointLabels(nPtIdx) = "Y"
                        End If
                    Else
                        Pe3doa1.PointLabels(nPtIdx) = CStr(tblData(1, aIdx(k))) 'label for X-Axis
                    End If
                End If
            Else
                Pe3doa1.PointLabels(nPtIdx) = ""
            End If
            nPtIdx = nPtIdx + 1
            k = k + 1
            'get min & max to use for contour area
            If dY > dMax Then dMax = dY
            If dY < dMin Then dMin = dY
        Next
    Next
    
    If Pe3doa1.PolyMode = PEPM_SURFACEPOLYGONS Then
        If tblData(1, nPoints - 1) > 1# And _
           tblData(2, nSubsets - 1) > 1# And dMax > 1000# Then
            m.bAutoScaleData = True
        Else
            m.bAutoScaleData = False
        End If
        '** add/subtract padding so contour will color all values **'
        dMin = dMin - (dMin * 0.05)     'subtract 5%
        dMax = dMax + (dMax * 0.05)     'add 5%
        '** set contour line properties **'
        Pe3doa1.ManualContourLine = 0
        Pe3doa1.ManualContourScaleControl = PEMSC_MINMAX
        Pe3doa1.ManualContourMin = dMin
        Pe3doa1.ManualContourMax = dMax
    ElseIf dMax > 1000 Then
        m.bAutoScaleData = True
    Else
        m.bAutoScaleData = False
    End If
    
    m.nSubsets = nSubsets
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChart.SetDataPe3doa1"

End Sub

Public Sub SetChartType3d(ByVal nPoly&, ByVal nPlot&)
On Error GoTo ErrSection:

    m.n3dPlotStyle = nPlot
    Pe3doa1.PlottingMethod = m.n3dPlotStyle
    If nPoly <> m.n3dPolyMode Then
        m.n3dPolyMode = nPoly
        ShowPe3doa1  'only need to redo image if polymode was changed
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChart.SetChartType3d"

End Sub

Public Property Get AllowAreaChart() As Long
    AllowAreaChart = m.nAllowArea
End Property

Public Property Get SingleColor() As Long
    SingleColor = m.nSingleColor
End Property

Public Property Let SingleColor(ByVal nSingle&)
On Error GoTo ErrSection:
    
    If nSingle <> m.nSingleColor Then
        m.nSingleColor = nSingle
        ShowPe3doa1
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmOptChart.SingleColor.Let"

End Property

Public Property Let Color(ByVal nColor&)
On Error GoTo ErrSection:
    
    If Pe3doa1.Visible = True Then
        ShowPe3doa1
    ElseIf Pegoa1.Visible = True Then
        If nColor <> m.n2dColor Then
            m.n2dColor = nColor
            Pegoa1.SubsetColors(0) = nColor
        End If
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmOptChart.Color.Let"

End Property

Public Property Get LegendLocation() As Long
    LegendLocation = m.nLegLocation
End Property

Public Property Let LegendLocation(ByVal nLocation&)
On Error GoTo ErrSection:
    
    m.nLegLocation = nLocation
    
    If m.nLegLocation = -1 Then
        Pe3doa1.ShowContourLegends = False
    Else
        Pe3doa1.ShowContourLegends = True
        Pe3doa1.LegendLocation = m.nLegLocation
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmOptChart.LegendLocation.Let"

End Property

Public Property Get PeFontSize() As Long
    FontSize = m.nFontSize
End Property

Public Property Let PeFontSize(ByVal nSize&)
On Error GoTo ErrSection:

    m.nFontSize = nSize
    If Pegoa1.Visible = True Then
        Pegoa1.FontSize = m.nFontSize
    ElseIf Pe3doa1.Visible = True Then
        Pe3doa1.FontSize = m.nFontSize
    End If
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmOptChart.PeFontSize.Let"

End Property

Public Property Get IdxAxisY() As Long
    IdxAxisY = m.nIdxY
End Property

Public Property Let IdxAxisY(ByVal idx As Long)
On Error GoTo ErrSection:
    
    If idx < 0 Or idx = m.nIdxY Or _
       idx >= m.fgData.Cols Then
       
       Exit Property
       
    End If
    
    m.nIdxY = idx
    If Pegoa1.Visible = True Then
        ShowPegoa1
    ElseIf Pe3doa1.Visible = True Then
        ShowPe3doa1
    End If
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "frmOptChart.IdxAxisY.Let"

End Property

Public Property Get IdxAxisX() As Long
    IdxAxisX = m.nIdxX
End Property

Public Property Get IdxAxisZ() As Long
    IdxAxisZ = m.nIdxZ
End Property

Public Property Get IdxVarCol() As Long
    IdxVarCol = m.nIdxVarStart
End Property

Public Property Get PointsToGraph() As Long
    If m.nPointsToGraph > 0 Then
        PointsToGraph = m.nPointsToGraph
    Else
        PointsToGraph = Pegoa1.Points
    End If
End Property

Public Property Let PointsToGraph(ByVal nPoints As Long)
    m.nPointsToGraph = nPoints
    If Pegoa1.Points > Pegoa1.PointsToGraph Then
        Pegoa1.HorzScrollPos = Pegoa1.Points - Pegoa1.PointsToGraph + 1
    End If
End Property

Public Property Get RotateDetail() As Long
    RotateDetail = m.nRotateDetail
End Property

Public Property Let RotateDetail(ByVal nDetail&)
    m.nRotateDetail = nDetail
    Pe3doa1.RotationDetail = m.nRotateDetail
End Property

Private Sub Pe3doa1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then ShowEditForm
End Sub

Private Sub Pegoa1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then ShowEditForm
End Sub

Private Sub tbToolbar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
On Error GoTo ErrSection:

    Select Case Tool.ID
        Case "ID_Settings"
            ShowEditForm
            
        Case "ID_Print"
            If Pegoa1.Visible = True Then
                PEnset Pegoa1, 2978, True '(to make dialog modal, otherwise "Setup" doesn't work right)
                Pegoa1.PEactions = 8
            ElseIf Pe3doa1.Visible = True Then
                PEnset Pe3doa1, 2978, True '(to make dialog modal, otherwise "Setup" doesn't work right)
                Pe3doa1.PEactions = 8
            End If
        
        Case "ID_Export"
            If Pegoa1.Visible = True Then
                Pegoa1.PEactions = 6
            ElseIf Pe3doa1.Visible = True Then
                Pe3doa1.PEactions = 6
            End If
            
        Case "ID_Rotate"
            AutoRotate
                        
        Case "ID_Cancel"
            Unload Me
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChart.tbToolbar_ToolClick"

End Sub

Private Sub ShowEditForm()
On Error GoTo ErrSection:
    
    If Pegoa1.Visible = True Then
        frmOptChartEdit2d.ShowMe Me, Pegoa1
    ElseIf Pe3doa1.Visible = True Then
        m.n3dViewHeight = Pe3doa1.ViewingHeight
        m.n3dRotationDegree = Pe3doa1.DegreeOfRotation
        frmOptChartEdit3d.ShowMe Me, m.aColors, Pe3doa1
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChart.ShowEditForm", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub AutoRotate()
On Error GoTo ErrSection:

    Dim bRotate As Boolean

    bRotate = Pe3doa1.AutoRotation
    
    If bRotate = False Then
        Pe3doa1.AutoRotation = True
        tbToolbar.Tools("ID_Rotate").ChangeAll ssChangeAllName, "Stop Rotation"
    Else
        Pe3doa1.AutoRotation = False
        tbToolbar.Tools("ID_Rotate").ChangeAll ssChangeAllName, "Rotate"
    End If
    
    With tbToolbar
        .Tools("ID_Cancel").Enabled = bRotate
        .Tools("ID_Settings").Enabled = bRotate
        .Tools("ID_Print").Enabled = bRotate
        .Tools("ID_Export").Enabled = bRotate
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChart.AutoRotate"

End Sub

Private Function SubTitle() As String
On Error GoTo ErrSection:

    Dim strText$

    strText = m.obOptimizer.ColumnDesc(m.nIdxY)
    strText = Trim(Parse(strText, "=", 2))
    
    SubTitle = strText

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmOptChart.SubTitle"

End Function

Private Function AxisLabelY() As String
On Error GoTo ErrSection:


    AxisLabelY = m.fgData.TextMatrix(0, m.nIdxY)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmOptChart.AxisLabelY"

End Function

Private Function AxisLabelHorz(ByVal nIdx&) As String
On Error GoTo ErrSection:

    Dim strText$

    strText = m.obOptimizer.ColumnDesc(nIdx)
    strText = Trim(Parse(strText, "=", 2))
    If Left(strText, 2) = "->" Then
        strText = Trim(Parse(strText, "->", 2))
    End If
    strText = strText + " (" + m.fgData.TextMatrix(0, nIdx) + ")"
    
    AxisLabelHorz = strText

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmOptChart.AxisLabelHorz"

End Function

Private Sub ReadSettings(ByVal b2D As Boolean)
On Error GoTo ErrSection:

    Dim strColY$, strType$, strFont$, strGrid$, strColor$
    Dim strPts$, strShadow$
    Dim strLegLoc$, strSingle$, strRotate$
    
    Dim aFields As New cGdArray
    Dim i&

    'properties common to both chart types
    If b2D = True Then
        strColY = GetIniFileProperty("YAxisColumn2d", "", "OptimizerChart", g.strIniFile)
        strType = GetIniFileProperty("ChartType2d", "", "OptimizerChart", g.strIniFile)
        strFont = GetIniFileProperty("FontSize2d", "", "OptimizerChart", g.strIniFile)
        strGrid = GetIniFileProperty("GridLines2d", "", "OptimizerChart", g.strIniFile)
        strColor = GetIniFileProperty("Color2d", "", "OptimizerChart", g.strIniFile)
        
        If strType <> "" Then m.n2dPlot = Val(strType)
        If strColor <> "" Then m.n2dColor = Val(strColor)
    Else
        strColY = GetIniFileProperty("YAxisColumn3d", "", "OptimizerChart", g.strIniFile)
        strType = GetIniFileProperty("ChartType3d", "", "OptimizerChart", g.strIniFile)
        strFont = GetIniFileProperty("FontSize3d", "", "OptimizerChart", g.strIniFile)
        strGrid = GetIniFileProperty("GridLines3d", "", "OptimizerChart", g.strIniFile)
        strColor = GetIniFileProperty("Color3d", "", "OptimizerChart", g.strIniFile)
        'saved format for 3d chart type is [PolyMode, PlotStyle]
        'saved format for 3d color is [color1, color2, ... color10]
        If strType <> "" Then
            aFields.SplitFields strType, ","
            If aFields.Size = 2 Then
                m.n3dPolyMode = Val(aFields(0))
                m.n3dPlotStyle = Val(aFields(1))
            End If
            aFields.Clear
        End If
        If strColor <> "" Then
            aFields.SplitFields strColor, ","
            For i = 0 To aFields.Size - 1
                m.aColors(i) = Val(aFields(i))
                If i > 9 Then Exit For
            Next
        End If
    End If
    If strColY <> "" Then m.nIdxY = Val(strColY)
    If strFont <> "" Then m.nFontSize = Val(strFont)
    If strGrid <> "" Then m.nGrid = Val(strGrid)
    
    '2d specific properites
    strPts = GetIniFileProperty("PtsPerScreen", "", "OptimizerChart", g.strIniFile)
    strShadow = GetIniFileProperty("DataShadow", "", "OptimizerChart", g.strIniFile)
    If strPts <> "" Then m.nPointsToGraph = Val(strPts)
    If strShadow <> "" Then m.n2dDataShadow = Val(strShadow)
    
    '3d specific properties
    strLegLoc = GetIniFileProperty("LegendLocation", "", "OptimizerChart", g.strIniFile)
    strSingle = GetIniFileProperty("SingleColor", "", "OptimizerChart", g.strIniFile)
    strRotate = GetIniFileProperty("RotateDetail", "", "OptimizerChart", g.strIniFile)
    If strLegLoc <> "" Then m.nLegLocation = Val(strLegLoc)
    If strSingle <> "" Then m.nSingleColor = Val(strSingle)
    If strRotate <> "" Then m.nRotateDetail = Val(strRotate)
    
    m.bSettingsRead = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChart.ReadSettings"

End Sub

Private Sub WriteSettings()
On Error GoTo ErrSection:

    Dim str3dType$, str3dColor$, i&

    If Pegoa1.Visible = True Then
        m.n2dPlot = Pegoa1.PlottingMethod
        m.nFontSize = Pegoa1.FontSize
        m.nGrid = Pegoa1.GridLineControl
        SetIniFileProperty "YAxisColumn2d", m.nIdxY, "OptimizerChart", g.strIniFile
        SetIniFileProperty "ChartType2d", m.n2dPlot, "OptimizerChart", g.strIniFile
        SetIniFileProperty "FontSize2d", m.nFontSize, "OptimizerChart", g.strIniFile
        SetIniFileProperty "GridLines2d", m.nGrid, "OptimizerChart", g.strIniFile
        SetIniFileProperty "Color2d", m.n2dColor, "OptimizerChart", g.strIniFile
    ElseIf Pe3doa1.Visible = True Then
        m.nFontSize = Pe3doa1.FontSize
        m.nGrid = Pe3doa1.GridLineControl
        str3dType = Str(m.n3dPolyMode) + "," + Str(m.n3dPlotStyle)
        For i = 0 To m.aColors.Size - 1
            str3dColor = str3dColor + Str(m.aColors(i)) + ","
        Next
        SetIniFileProperty "YAxisColumn3d", m.nIdxY, "OptimizerChart", g.strIniFile
        SetIniFileProperty "ChartType3d", str3dType, "OptimizerChart", g.strIniFile
        SetIniFileProperty "FontSize3d", m.nFontSize, "OptimizerChart", g.strIniFile
        SetIniFileProperty "GridLines3d", m.nGrid, "OptimizerChart", g.strIniFile
        SetIniFileProperty "Color3d", str3dColor, "OptimizerChart", g.strIniFile
    End If
    
    m.n2dDataShadow = Pegoa1.DataShadows
    SetIniFileProperty "PtsPerScreen", m.nPointsToGraph, "OptimizerChart", g.strIniFile
    SetIniFileProperty "DataShadow", m.n2dDataShadow, "OptimizerChart", g.strIniFile

    SetIniFileProperty "LegendLocation", m.nLegLocation, "OptimizerChart", g.strIniFile
    SetIniFileProperty "SingleColor", m.nSingleColor, "OptimizerChart", g.strIniFile
    SetIniFileProperty "RotateDetail", m.nRotateDetail, "OptimizerChart", g.strIniFile
    
    m.bSettingsRead = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChart.WriteSettings"

End Sub

'This routine is here because it is used by
'both frmOptChartEdit2d & frmOptChartEdit3d
Public Sub InitComboY(cboData As Control)
On Error GoTo ErrSection:
    
    Dim i&, k&, j&, strCol$

    cboData.Clear
    
    k = 0
    For i = 1 To m.fgData.Cols - 1
        strCol = m.fgData.TextMatrix(0, i)
        
        If (OptimizeColumn(i) = False) And (m.fgData.ColHidden(i) = False) Then
            cboData.AddItem strCol
            cboData.ItemData(k) = i
            If i = m.nIdxY Then j = k   'index of charted column
            k = k + 1
        End If
    Next
         
    If cboData.ListCount > 0 Then
        cboData.ListIndex = j
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmOptChart.InitComboY"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OptimizeColumn
'' Description: Determine if the given column is an optimize column
'' Inputs:      Column
'' Returns:     True if optimize column, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function OptimizeColumn(ByVal lColumn As Long) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim strColumnName As String         ' Column name
    
    bReturn = False
    strColumnName = m.fgData.TextMatrix(0, lColumn)
    
    If m.obOptimizer.Mode = eGDOptMode_HighlightBarReport Then
        bReturn = (UCase(strColumnName) = "# BARS")
    Else
        bReturn = (UCase(Left(strColumnName, 2)) = "I-") Or (UCase(Left(strColumnName, 2)) = "R-")
    End If
    
    OptimizeColumn = bReturn

ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmOptChart.OptimizeColumn"
    
End Function
