Attribute VB_Name = "mCommon"
Option Explicit
Option Compare Text

Global Const gExitSignal = 1
Global Const gEntrySignal = 0
Global Const gUserErr = vbObjectError + 1000
Global Const gbForceMM As Boolean = False

Public Enum eGDReportViewOptions
    eGDReportViewOption_Trades = 0
    eGDReportViewOption_Monthly
    eGDReportViewOption_Yearly
End Enum

Public Enum eGDEquityFilterMode
    eGDEquityFilterMode_BelowMa = 0
    eGDEquityFilterMode_MaDown
End Enum

Public Enum eGDTakeNextTradeValue
    eGDTakeNextTrade_No = 0
    eGDTakeNextTrade_Yes
    eGDTakeNextTrade_NotEnoughData
    eGDTakeNextTrade_NoEquityFilter
End Enum

Type GlobalStruct
    strAppPath As String
    Help As Object
    bShowInLocalTimeZone As Boolean     ' Show times in local time?
    nAltGridRowColor As Long ' grid background color for alternate rows
    IrxBars As cGdBars ' to hold t-bill %'s for Sharpe calc
    
    'RH
    Styler As New cStyler
    
End Type

'RH
Public Enum eStyleColorTypes
    'form
    eForm_Background = 0
    
    'frame
    eFrame_Background = 1
    eFrame_Border = 2
    
    'button
    eButton_Background = 3
    eButton_Border = 4
    eButton_Text = 5
    
    'checkbox
    eCheck_Border = 6
    eCheck_Background = 7
    eCheck_Forecolor = 8
    
    'flexgrid
    eGrid_Background = 9
    
End Enum



Global g As GlobalStruct

Public Sub ReSizeMDIChildForm(frm As Form, ctl As Control)
    
    frm.Width = ctl.Left + ctl.Width + frm.Width - frm.ScaleWidth
    frm.Height = ctl.Top + ctl.Height + frm.Height - frm.ScaleHeight

End Sub

' Asks the user if they want to rename or copy when the name has changed
Public Function RenameOrCopy(ByVal strType As String) As String
On Error GoTo ErrSection:

    Dim strMsg As String
    
    strMsg = "Name of the " & strType & " has changed.||Do you wish to rename the existing " & strType & ",| or create a copy with the new name?"
    RenameOrCopy = AskBox("i=? ; b=+Copy|-Rename ; h=Copy or Rename ; " & strMsg)
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mCommon.RenameOrCopy", eGDRaiseError_Raise, frmReports.AppPath

End Function

Public Function FormatDollar(ByVal strShowCents As String) As String
On Error GoTo ErrSection:

    If strShowCents = "Yes" Then
        FormatDollar = "$#,##0.00"
    Else
        FormatDollar = "$#,##0"
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mCommon.FormatDollar", eGDRaiseError_Raise, frmReports.AppPath

End Function

'Create an Outlook style title in the first row (row 0 of grid)
Public Sub OutLookTitle(pTitle As String, pRow As Long, pFromCol As Long, _
    pToCol As Long, pGrid As VSFlexGrid, pBackColor As Long, pForeColor As Long, _
    pFontName As String, pFontSize As Long, pFontBold As Boolean, _
    pCellHeightPctIncrease As Single)
On Error GoTo ErrSection:

    Dim lRedraw As Long
    
    With pGrid
        lRedraw = .Redraw
        .Redraw = flexRDNone
        .Cell(flexcpText, pRow, pFromCol, pRow, pToCol) = pTitle
        .Cell(flexcpBackColor, pRow, pFromCol, pRow, pToCol) = pBackColor
        .Cell(flexcpFontBold, pRow, pFromCol, pRow, pToCol) = pFontBold
        '.Cell(flexcpFontName, pRow, pFromCol, pRow, pToCol) = pFontName
        '.Cell(flexcpFontSize, pRow, pFromCol, pRow, pToCol) = pFontSize
        .Cell(flexcpForeColor, pRow, pFromCol, pRow, pToCol) = pForeColor
        '.RowHeight(pRow) = .CellHeight * pCellHeightPctIncrease
        .MergeCells = flexMergeRestrictRows
        .MergeRow(pRow) = True
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mCommon.OutlookTitle", eGDRaiseError_Raise, frmReports.AppPath

End Sub

Public Sub ColorNegValue(pGrid As VSFlexGrid, pValue As Double, pRow As Long, pCol As Long)
On Error GoTo ErrSection:
    
    If pValue < 0 Then
        pGrid.Cell(flexcpForeColor, pRow, pCol) = vbRed
    Else
        pGrid.Cell(flexcpForeColor, pRow, pCol) = vbDefault
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mCommon.ColorNegValue", eGDRaiseError_Raise, frmReports.AppPath

End Sub

Public Sub ClearGrid(vsGrid As VSFlexGrid)
On Error GoTo ErrSection:

    Dim lRedraw As Long

    With vsGrid
        lRedraw = .Redraw
        .Redraw = flexRDNone
        .FlexDataSource = Nothing
        .Rows = 0
        .Cols = 0
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mCommon.ClearGrid", eGDRaiseError_Raise, frmReports.AppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetChartSettings
'' Description: Set common chart settings
'' Inputs:      Chart to change settings on, Plotting Method
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetChartSettings(Chart As Pegoa, PlotMethod As ePlottingMethod)
On Error GoTo ErrSection:

    With Chart
        .DeskColor = RGB(255, 255, 255)
        .GraphBackColor = RGB(255, 255, 255)
        .GraphForeColor = 0
        .PlottingMethod = PlotMethod
        .AllowCustomization = False
        .AllowPopup = False
        .AllowRibbon = True
        .FocalRect = False
        .ShadowColor = .DeskColor
        .GridLineControl = PEGLC_NONE
        .FontSizeLegendCntl = 0.8
        .AllowJpegOutput = True
        
        If .PlottingMethod = GPM_BAR Or .PlottingMethod = GPM_HORIZONTALBAR Then
            .SubsetColors(0) = RGB(0, 192, 192)
            .SubsetLineTypes(0) = PELT_THINSOLID
            .SubsetPointTypes(0) = PEPT_DOTSOLID
            .DataShadows = PEDS_3D
            If .Subsets = 2 Then
                .SubsetColors(0) = vbCyan
                .SubsetColors(1) = RGB(0, 128, 128)
                .SubsetLineTypes(1) = PELT_THINSOLID
                .SubsetPointTypes(1) = PEPT_DOTSOLID
                .DataShadows = PEDS_SHADOWS
            End If
        ElseIf .PlottingMethod = GPM_LINE Then
            .SubsetColors(0) = vbRed
            .SubsetLineTypes(0) = PELT_THINSOLID
            .SubsetPointTypes(0) = PEPT_DOTSOLID
            If .Subsets = 2 Then
                .SubsetColors(1) = vbBlue
                .SubsetLineTypes(1) = PELT_THINSOLID
                .SubsetPointTypes(1) = PEPT_DOTSOLID
            Else
                .SubsetsToLegend(0) = -1
            End If
            .GridLineControl = PEGLC_XAXIS ' = PEGLC_NONE
            .GridStyle = PEGS_DOT
            .GridInFront = False
            .ShowAnnotations = True
            .DataShadows = PEDS_NONE
            .YAxisOnRight = True
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mCommon.SetChartSettings", eGDRaiseError_Raise, frmReports.AppPath

End Sub
    
Public Sub ShowChart(ByVal lItemsToChart As Long, Chart As Pegoa, lbl As Label, _
                lblMM As Label, Optional ByVal lMMItems As Long = -1&)
On Error GoTo ErrSection:

    If lItemsToChart = 0 Then
        Chart.Visible = False
        lblMM.Visible = False
        lbl.Visible = True
    ElseIf lMMItems = 0 Then
        lbl.Visible = False
        Chart.Visible = False
        lblMM.Visible = True
    Else
        lbl.Visible = False
        lblMM.Visible = False
        Chart.Visible = True
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mCommon.ShowChart", eGDRaiseError_Raise, frmReports.AppPath

End Sub

Public Sub ShowMsg(Optional ByVal lErrNum& = 0, Optional ByVal strSource$ = "", _
                    Optional ByVal strDesc$ = "")
    
    Dim RetVal As Variant
    
    If lErrNum = 0 Then
        lErrNum = Err.Number
        strSource = Err.Source
        strDesc = Err.Description
    End If
    
    RetVal = LockWindowUpdate(0)
    Screen.MousePointer = vbDefault
    
    If lErrNum < 0 Then
        Replace strDesc, vbCrLf, "|"
        InfBox strDesc, , , "Error", , , , , , , , eGDAlign_Left
    Else
        Replace strDesc, vbCrLf, "|"
        InfBox "An unexpected error occurred.||Please report the following: " & _
            "|Source:  " & strSource & _
            "|Message: " & strDesc, , , "Error", , , , , , , , eGDAlign_Left
    End If

End Sub

Public Sub DoPrintHeader(Optional ByVal nFontSize& = 12)

    With frmPrintPreview.vp
        .LineSpacing = 100
        .HdrFontName = "Times New Roman"
        .HdrFontSize = nFontSize
        .Header = "|Trade Navigator" & vbCrLf & "Genesis Financial Technologies - (800) 808-DATA - www.TradeNavigator.com"
        .Footer = "|Page: %d|"
    End With

End Sub

Public Sub GridToFile(ByVal vsGrid As VSFlexGrid, ByVal strFileName As String)
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Index into a for loop
    Dim lCol As Long                    ' Index into a for loop
    Dim astrFile As New cGdArray        ' Array to hold information to dump to file
    Dim strTemp As String               ' Temporary string variable
    
    astrFile.Create eGDARRAY_Strings
    
    With vsGrid
        For lRow = .FixedRows To .Rows - 1
            strTemp = ""
            
            For lCol = 0 To .Cols - 1
                If .ColHidden(lCol) = False Then
                    strTemp = strTemp & .TextMatrix(lRow, lCol) & ","
                End If
            Next lCol
            
            astrFile.Add Left(strTemp, Len(strTemp) - 1)
        Next lRow
    End With
    
    astrFile.ToFile strFileName

ErrExit:
    Set astrFile = Nothing
    Exit Sub
    
ErrSection:
    Set astrFile = Nothing
    RaiseError "mCommon.GridToFile", eGDRaiseError_Raise, g.strAppPath
    
End Sub
