Attribute VB_Name = "mPFPCore"
Option Explicit

Public Const MAX_ADDR_REC = 16377      ' max # of items for
Public Const MAX_CORE_REC = MAX_ADDR_REC * 2

Public Const MATCH_OPEN = 1
Public Const MATCH_HIGH = 2
Public Const MATCH_LOW = 4
Public Const MATCH_CLOSE = 8

Public Enum eColsPFP
    eColsPFP_Use = 0
    eColsPFP_Date
    eColsPFP_Day
    eColsPFP_CorrPercent
    eColsPFP_Index
    eColsPFP_DateDouble
End Enum

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Structures & function prototypes for NEW DLL
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Type CoreBars
    max_bar As Long
    first_bar As Long
    last_bar As Long
    dummy As Long
    jdate_ptr As Long
    hourmin_ptr As Long     ' 0 if not used
    open_ptr As Long
    high_ptr As Long
    low_ptr As Long
    close_ptr As Long
    vol_ptr As Long
    oi_ptr As Long
    tot_vol_ptr As Long
    tot_oi_ptr As Long
    dummy2 As String * 42
End Type

Declare Function PFP_CorrelationMatches2 Lib "G32_PFP.DLL" _
    (Pattern As CoreBars, _
     Search As CoreBars, _
     ByVal match_type As Long, _
     Filter As Long, _
     ByVal min_corr As Double, _
     ByVal max_hits As Long, _
     hits As Long, _
     hit_corr As Double) As Long

Declare Function PFP_BuildComposite2 Lib "G32_PFP.DLL" _
    (prices As CoreBars, _
     ByVal match_type As Long, _
     ByVal ptrn_len As Long, _
     ByVal fcast_len As Long, _
     ByVal num_hits As Long, _
     hits As Long, _
     composite As CoreBars, _
     strength As Double) As Long


' - MethodWeighting array (0.0-1.0 weighting for each type of correlation "rule"):
'       0 = weight for the added "difference" indicators (when multiple indicators in same pane)
'       1 = standard correlation formula
'       2 = normalized correlation (a "PercentR" type of comparison)
'       3 = directional correlation (% of up/down from previous value)
'       4 = sign correlation (% of positive/negative)
'       5 = highest/lowest peak correlation (comparison of bar# of highest and lowest peaks)
' - DatesArray: only used for logging and debugging purposes
' - IndicatorArrays: array of array handles for each indicator being passed
' - PaneIDs: pane ID for each indicator (but should be 0 for any overlayed indicators)
' - MatchTable: 2 column table for all hits (0 = end bar#, 1 = correlation)
Declare Function PFP_IndicatorMatches Lib "G32_PFP.DLL" _
    (ByVal hMethodWeighting As Long, ByVal hDatesArray As Long, ByVal hIndicatorArrays As Long, _
    ByVal hPaneIDs As Long, ByVal hIndFlags As Long, ByVal iPatternEndBar As Long, ByVal iPatternLength As Long, _
    ByVal iMinCorr As Long, ByVal hMatchTable As Long, ByVal strLogFile As String) As Long


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Subroutines shared by chart & pattern for profit forms
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub InitGridPFP(fg As VSFlexGrid)
On Error GoTo ErrSection:
    
    If fg Is Nothing Then Exit Sub

    With fg
        .Redraw = flexRDNone
        SetupGrid fg, eGridMode_Grid
        .ExplorerBar = flexExSortShow
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .SelectionMode = flexSelectionFree
        .ScrollBars = flexScrollBarVertical
        .HighLight = flexHighlightNever
        .Editable = flexEDKbdMouse
        .ExtendLastCol = True
        .FixedCols = 0
        .FixedRows = 1
        .Rows = .FixedRows
        .Cols = 5
        
        .ColWidthMin = 15               '6392
        .ColWidthMax = 1250
        
        'alignment
        .ColAlignment(eColsPFP_Use) = flexAlignCenterCenter
        .ColAlignment(eColsPFP_Date) = flexAlignLeftCenter
        .ColAlignment(eColsPFP_Day) = flexAlignLeftCenter
        .ColAlignment(eColsPFP_CorrPercent) = flexAlignRightCenter
        'column headers
        .TextMatrix(0, eColsPFP_Use) = "Use"
        .TextMatrix(0, eColsPFP_Date) = "Date"
        .TextMatrix(0, eColsPFP_Day) = "Day"
        .TextMatrix(0, eColsPFP_CorrPercent) = "%Fit" ' "Corr"
        'data type
        .ColDataType(eColsPFP_Use) = flexDTBoolean
        .ColDataType(eColsPFP_Day) = flexDTLong
        .ColDataType(eColsPFP_Date) = flexDTDate
        
        .ColSort(eColsPFP_Day) = flexSortNone
        .ColSort(eColsPFP_Date) = flexSortNone
        .ColSort(eColsPFP_Use) = flexSortNone
        .ColSort(eColsPFP_Index) = flexSortNone
        
        .ColFormat(eColsPFP_Day) = "ddd"
        
        'columns width
        If IsFrmChart(fg.Parent) Then
            .ColWidth(eColsPFP_Use) = 360
            .ColWidth(eColsPFP_Day) = 420       '345
            .ColWidth(eColsPFP_Date) = 1020
            .ColWidth(eColsPFP_CorrPercent) = 495   '465
            .ColWidth(eColsPFP_Index) = 90
        Else
            .ColWidth(eColsPFP_Use) = 500
            .ColWidth(eColsPFP_Day) = 500
            .ColWidth(eColsPFP_Date) = 1050 '900   '1400
            .ColWidth(eColsPFP_CorrPercent) = 500
            .ColWidth(eColsPFP_Index) = 100
        End If
        'hidden columns
        .ColHidden(eColsPFP_Index) = True
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mPFPCore.InitGridPFP"

End Sub

Public Function LoadIndGridPFP(Chart As cChart, fgIndicators As VSFlexGrid, _
    ByVal bReset As Boolean, Optional ByVal bRestoreTemplate As Boolean = False) As String
On Error GoTo ErrSection:

    Static strPrevList As String

    Dim i&, j&, k&
    Dim iPriceIndId&, iNoneCount&
    
    Dim strList$, strCompSym$
    Dim strText$, strData$
    
    Dim Ind As cIndicator
    Dim Pane As cPane
    Dim aIndAvail As cGdArray
    
    
    If fgIndicators Is Nothing Then Exit Function
    
    strList = "|#-999;None|#-888;Close|#-777;Open|#-666;High|#-555;Low"
    
    If Not Chart Is Nothing Then
        For i = 1 To Chart.Tree.Count
            If Chart.Tree.NodeLevel(i) > 0 Then
                Set Ind = Chart.Tree(i)
                If Not Ind Is Nothing Then
                    Set Pane = Chart.Tree(Ind.geIndpaneId)
                    If Not Pane Is Nothing Then
                        If Pane.Display And Ind.Display Then
                            If Ind.DataType = eINDIC_Array Or Ind.DataType = eINDIC_BooleanArray Then
                                strList = strList & "|#" & Str(Ind.geIndId) & ";" & Ind.ChartLabel
                            ElseIf Ind.DataType = eINDIC_BarData And Ind.isPriceInd <> 1 Then
                                If Not Ind.Bars Is Nothing Then
                                    strCompSym = Ind.Bars.Prop(eBARS_Symbol)
                                    If Len(strCompSym) > 0 Then
                                        strList = strList & "|#" & Str(Ind.geIndId) & ";Close Of " & strCompSym
                                        
                                        'need to make the number after "|#" unique else flex grid treats all item as same
                                        '!Important! the cPatternProfit.LoadIndicators uses the reverse math operation to
                                        '            extract the Ind.geIndId value so need to keep this set of code in sync
                                        strList = strList & "|#" & Str(Ind.geIndId * (Chart.Tree.Count + 1)) & ";Open Of " & strCompSym
                                        strList = strList & "|#" & Str(Ind.geIndId * (Chart.Tree.Count + 2)) & ";High Of " & strCompSym
                                        strList = strList & "|#" & Str(Ind.geIndId * (Chart.Tree.Count + 3)) & ";Low Of " & strCompSym
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End If

    If bRestoreTemplate And Len(Chart.PfpIndicators) > 0 Then
        fgIndicators.ColComboList(0) = strList
        RestoreIndGridPFP Chart, fgIndicators
    Else
    
        If strList = strPrevList Then
            'make sure the grid is populated before exiting
            'user might have brought up a new chart with same template which would
            'cause the current & previous indicator list to be same
            If fgIndicators.Rows = 4 Then
                For i = 0 To 3
                    strText = fgIndicators.TextMatrix(i, 0)
                    If Len(strText) = 0 Then Exit For
                Next
                If Len(strText) > 0 Then
                    LoadIndGridPFP = strList
                    GoTo ErrExit
                End If
            End If
        End If
    
        'JM 08-02-2011: Aardvark 6410 - this code rewrote to ignore bReset
        '   leave awhile then remove parameter from function if all ok,
        '   or revert code to use bReset if necessary
    
        Set aIndAvail = New cGdArray
        aIndAvail.SplitFields strList, "|"
        
        With fgIndicators
            .ColComboList(0) = strList
            
            'loop through grid & make sure the hidden data string matches the new indicator list
            For i = 0 To 3
                strData = ""
                strText = .TextMatrix(i, 0)
                
                If Len(strText) = 0 Or UCase(strText) = "NONE" Then
                    .TextMatrix(i, 0) = "None"
                    iNoneCount = iNoneCount + 1
                Else
                    For j = 0 To aIndAvail.Size - 1
                        If InStr(aIndAvail(j), strText) <> 0 Then
                            'make sure text is exact match (eg low & bollinger lower band will both return <> 0)
                            strData = Parse(aIndAvail(j), ";", 2)
                            If strData = strText Then
                                strData = aIndAvail(j)
                                Exit For
                            Else
                                strData = ""
                            End If
                        End If
                    Next
                    
                    If Len(strData) = 0 Then
                        .TextMatrix(i, 0) = "None"
                        iNoneCount = iNoneCount + 1
                    Else
                        .Cell(flexcpData, i, 0) = strData
                    End If
                End If
            Next
            
            'set first row in grid to "Close" if nothing selected
            If iNoneCount = 4 Then
                .TextMatrix(0, 0) = "Close"
                .Cell(flexcpData, 0, 0) = "#-888;Close"
            End If
            
        End With
    End If
    
    If Not Chart Is Nothing Then
        If Not Chart.Form Is Nothing Then
            If Not Chart.Form.PatternProfitObj Is Nothing Then
                Chart.Form.PatternProfitObj.LoadSettings Chart
            End If
        End If
    End If
    
    strPrevList = strList
    
    LoadIndGridPFP = strList
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mPFPCore.InitIndicatorGrid"

End Function

Private Sub RestoreIndGridPFP(Chart As cChart, fgIndicators As VSFlexGrid)
On Error GoTo ErrSection:

    Dim i&, j&, k&, nID&
    Dim strText$, strKey$, strSubkey$

    Dim aList As New cGdArray
    Dim aIndKeys As New cGdArray
    Dim Ind As cIndicator

    If Chart Is Nothing Or fgIndicators Is Nothing Then Exit Sub

    strText = fgIndicators.ColComboList(0)
    aList.SplitFields strText, "|"

    k = 0
    strText = Chart.PfpIndicators
    If Len(strText) > 0 Then
        aIndKeys.SplitFields strText, "|"

        For i = 0 To aIndKeys.Size
            If k >= fgIndicators.Rows Then Exit For
            
            nID = -1
            strKey = aIndKeys(i)
            If strKey = "-555" Or strKey = "-666" Or strKey = "-777" Or strKey = "-888" Then
                strKey = "#" & strKey
            ElseIf InStr(strKey, ";") = 0 Then
                Set Ind = Chart.Tree(strKey)
                If Not Ind Is Nothing Then strKey = "#" & Ind.geIndId
            Else
                strText = Parse(strKey, ";", 1)
                Set Ind = Chart.Tree(strText)
                If Not Ind Is Nothing Then
                    nID = Ind.geIndId
                    strText = UCase(Parse(strKey, ";", 2))
                    Select Case strText
                        Case "CLOSE"
                            strKey = "#" & nID
                        Case "OPEN"
                            strKey = "#" & nID * (Chart.Tree.Count + 1)
                        Case "HIGH"
                            strKey = "#" & nID * (Chart.Tree.Count + 2)
                        Case "LOW"
                            strKey = "#" & nID * (Chart.Tree.Count + 3)
                        Case Else
                            strKey = ""
                    End Select
                End If
            End If
            
            If Len(strKey) > 0 Then
                For j = 0 To aList.Size
                    If strKey = Parse(aList(j), ";", 1) Then
                        strText = Parse(aList(j), ";", 2)
                        If Len(strText) > 0 Then
                            fgIndicators.TextMatrix(k, 0) = strText
                            fgIndicators.Cell(flexcpData, k, 0) = aList(j)
                            k = k + 1
                        End If
                        
                        Exit For
                    End If
                Next
            End If
        Next
    End If
        
    With fgIndicators
        For i = k To .Rows - 1
            If Len(.TextMatrix(i, 0)) = 0 Then
                .TextMatrix(i, 0) = "None"
                .Cell(flexcpData, i, 0) = ""
            End If
        Next
    End With

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mPFPCore.RestoreIndGridPFP"

End Sub

Public Sub SaveIndGridPFP(Chart As cChart, fgIndicators As VSFlexGrid)
On Error GoTo ErrSection:

    Dim i&, nID&, strKeys$, strText$, strData$
    Dim Ind As cIndicator


    If Chart Is Nothing Then Exit Sub

    With fgIndicators
        For i = .FixedRows To .Rows - 1
            strText = UCase(.TextMatrix(i, 0))

            If Len(strText) <> 0 And strText <> "NONE" Then
                strData = Parse(.Cell(flexcpData, i, 0), ";", 1)
                If Left(strData, 1) = "#" Then
                    strData = Replace(strData, "#", "")
                    nID = ValOfText(strData)
                Else
                    nID = -1
                End If
                
                If nID = -555 Or nID = -666 Or nID = -777 Or nID = -888 Then
                    'special case for Price indicator
                    strKeys = strKeys & nID & "|"
                Else
                    Set Ind = Chart.Tree(nID)

                    If Ind Is Nothing Then
                        '!Important!
                        'this code relies on the mPFPCore.InitIndicatorGrid
                        'routine assigning the numeric using equivalent
                        'reverse math so need to keep code in sync
                        If Left(strText, 9) = "CLOSE OF " Then
                            nID = ValOfText(.TextMatrix(i, 0)) / (Chart.Tree.Count + 1)
                            
                            Set Ind = Chart.Tree(nID)
                            If Not Ind Is Nothing Then
                                strKeys = strKeys & Ind.MyKey & ";CLOSE|"
                            End If
                            Set Ind = Nothing
                        ElseIf Left(strText, 8) = "OPEN OF " Then
                            nID = nID / (Chart.Tree.Count + 1)
                            
                            Set Ind = Chart.Tree(nID)
                            If Not Ind Is Nothing Then
                                strKeys = strKeys & Ind.MyKey & ";OPEN|"
                            End If
                            Set Ind = Nothing
                        ElseIf Left(strText, 8) = "HIGH OF " Then
                            nID = nID / (Chart.Tree.Count + 2)
                            
                            Set Ind = Chart.Tree(nID)
                            If Not Ind Is Nothing Then
                                strKeys = strKeys & Ind.MyKey & ";HIGH|"
                            End If
                            Set Ind = Nothing
                        ElseIf Left(strText, 7) = "LOW OF " Then
                            nID = nID / (Chart.Tree.Count + 3)
                            
                            Set Ind = Chart.Tree(nID)
                            If Not Ind Is Nothing Then
                                strKeys = strKeys & Ind.MyKey & ";LOW|"
                            End If
                            Set Ind = Nothing
                        Else
                            Set Ind = Chart.Tree(nID)
                            strKeys = strKeys & Ind.MyKey & "|"
                        End If
                    End If

                    If Not Ind Is Nothing Then
                        If Ind.DataType = eINDIC_BarData Then
                            strKeys = strKeys & Ind.MyKey & ";" & Parse(strText, " ", 1) & "|"
                        Else
                            strKeys = strKeys & Ind.MyKey & "|"
                        End If
                    End If
                    
                End If
            End If
        Next
    End With

    Chart.PfpIndicators = strKeys

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mPFPCore.SaveIndGridPFP"

End Sub
