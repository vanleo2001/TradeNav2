VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMkBreadth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum eGridLocation
    eTopLeft = 0
    eTopRight
    eBottomLeft
    eBottomRight
    eTab
    eSummary
End Enum

Private Type mGridData
    nCols As Long
    nRows As Long
    nFixedCols As Long
    nFixedRows As Long
    aData As New cGdArray
End Type

Private Type mPrivate
    strDataFile As String
    strCurrExchange As String
    aExchangeNames As New cGdArray
    'grid data - 'pt' prefix stands for private type
    ptTopLeft As mGridData
    ptTopRight As mGridData
    ptBottomLeft As mGridData
    ptBottomRight As mGridData
    'tabs data
    nTabCount As Long
    nCurrTab As Long
    aTabTittles As New cGdArray
    ptTabGrid As mGridData
    'settings
    nValueStyle As Long
    nAdvancedColor As Long
    nDeclinedColor As Long
End Type

Private m As mPrivate

Private Sub Class_Initialize()
    m.nCurrTab = -1
    m.nAdvancedColor = vbGreen
    m.nDeclinedColor = vbRed
End Sub

Public Property Let DataFileName(ByVal strName$)
    m.strDataFile = strName
End Property

Public Property Get ExchangeCount() As Long
    ExchangeCount = m.aExchangeNames.Size
End Property

Public Property Get ExchangeNames() As cGdArray
    
    Dim strExchange$
    
    If m.aExchangeNames.Size = 0 Then
        strExchange = GetIniFileProperty("Names", "", "Exchanges", m.strDataFile)
        m.aExchangeNames.SplitFields strExchange, "|"
    End If
    
    Set ExchangeNames = m.aExchangeNames
    
End Property

Public Property Let ExchangeData(ByVal strExchange)
    
    If m.strCurrExchange <> strExchange Then
        m.strCurrExchange = strExchange
        ReadGridData
        ReadTabsTittles
    End If

End Property

Public Property Get TabTittles() As cGdArray
    Set TabTittles = m.aTabTittles
    
    Dim i&
    i = m.aTabTittles.Size
End Property


Public Sub DataToGrid(oGrid As VSFlexGrid, ByVal eLoc As eGridLocation, _
    Optional ByVal nTabIdx& = 0)

    Select Case eLoc
        Case eTopLeft:
            SetGridData oGrid, m.ptTopLeft
        Case eTopRight:
            SetGridData oGrid, m.ptTopRight
        Case eBottomLeft:
            SetGridData oGrid, m.ptBottomLeft
        Case eBottomRight:
            SetGridData oGrid, m.ptBottomRight
        Case eTab
            SetGridTab oGrid, nTabIdx
        Case eSummary
            SetGridSummary oGrid
    End Select

End Sub

Private Sub ReadTabsTittles()

    Dim i&
    Dim strProp$, strTittle$

    m.aTabTittles.Size = 0
    m.nTabCount = GetIniFileProperty("Tabs Count", 0, m.strCurrExchange, m.strDataFile)
    
    For i = 0 To m.nTabCount - 1
        strProp = "Tab" & Str(i) & " Tittle"
        strTittle = GetIniFileProperty(strProp, "", m.strCurrExchange, m.strDataFile)
        If Len(strTittle) > 0 Then m.aTabTittles.Add strTittle
    Next
    
    i = m.aTabTittles.Size

End Sub

Private Sub ReadGridData()

    ReadGridDimension m.ptTopLeft, "TLeftGrid"
    ReadGridRowData m.ptTopLeft, "TLeftGrid"
    
    ReadGridDimension m.ptTopRight, "TRightGrid"
    ReadGridRowData m.ptTopRight, "TRightGrid"
    
    ReadGridDimension m.ptBottomLeft, "BLeftGrid"
    ReadGridRowData m.ptBottomLeft, "BLeftGrid"
    
    ReadGridDimension m.ptBottomRight, "BRightGrid"
    ReadGridRowData m.ptBottomRight, "BRightGrid"

End Sub

Private Sub ReadGridDimension(ptData As mGridData, ByVal strIniProp)

    Dim strProp$
    
    strProp = strIniProp & " Columns"
    ptData.nCols = GetIniFileProperty(strProp, 0, m.strCurrExchange, m.strDataFile)
    
    strProp = strIniProp & " Rows"
    ptData.nRows = GetIniFileProperty(strProp, 0, m.strCurrExchange, m.strDataFile)
    
    strProp = strIniProp & " FixedCols"
    ptData.nFixedCols = GetIniFileProperty(strProp, 0, m.strCurrExchange, m.strDataFile)
    
    strProp = strIniProp & " FixedRows"
    ptData.nFixedRows = GetIniFileProperty(strProp, 0, m.strCurrExchange, m.strDataFile)
        
End Sub

Private Sub ReadGridRowData(ptData As mGridData, ByVal strIniProp)

    Dim i&, k&
    Dim strProp$, strRowData$
    
    ptData.aData.Size = 0
    k = ptData.nRows - 1
    
    For i = 0 To k
        strProp = strIniProp & " Row" & Str(i)
        strRowData = GetIniFileProperty(strProp, "", m.strCurrExchange, m.strDataFile)
        If Len(strRowData) > 0 Then ptData.aData.Add strRowData
    Next
    
End Sub

Private Sub SetGridData(oGrid As VSFlexGrid, ptData As mGridData)

    Dim i&, k&, j&, m&
    Dim strRowData$
    Dim strColData As New cGdArray
    
    With oGrid
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Cols = ptData.nCols
        .FixedCols = 0
        .Rows = 0
    End With
    
    k = ptData.nRows - 1
    For i = 0 To k
        strRowData = ptData.aData(i)
        If Len(strRowData) > 0 Then
            strRowData = Replace(strRowData, "|", vbTab)
            oGrid.AddItem strRowData
        End If
    Next
    
    oGrid.FixedRows = ptData.nFixedRows
        
End Sub

Private Sub SetGridTab(oGrid As VSFlexGrid, ByVal nIdx&)

    Dim strProp$
    
    If m.nCurrTab <> nIdx Or oGrid.Rows = 0 Then
        strProp = "Tab" & Str(nIdx)
        ReadGridDimension m.ptTabGrid, strProp
        ReadGridRowData m.ptTabGrid, strProp
        SetGridData oGrid, m.ptTabGrid
        m.nCurrTab = nIdx
    End If

End Sub

Private Sub SetGridSummary(oGrid As VSFlexGrid)

    Dim i&, k&
    Dim dTotal#, dPercent#
    Dim strExchange$, strText$
    Dim aData As New cGdArray
    
    With oGrid
        .Cols = 5
        .FixedCols = 1
        .FixedRows = 1
        .TextMatrix(0, 0) = "Exchange"
        .TextMatrix(0, 1) = "Advanced"
        .TextMatrix(0, 2) = "Declined"
        .TextMatrix(0, 3) = "Unchanged"
        .TextMatrix(0, 4) = "Total Issues"
        .Rows = .FixedRows
        .Cell(flexcpFontBold, 0, 0, 0, 4) = True
        .ColAlignment(1) = flexAlignRightCenter
        .ColAlignment(2) = flexAlignRightCenter
        .ColAlignment(3) = flexAlignRightCenter
        .ColAlignment(4) = flexAlignRightCenter
        .ScrollBars = flexScrollBarVertical
    End With
    
    k = m.aExchangeNames.Size
    
    For i = 0 To k - 1
        strExchange = m.aExchangeNames(i)
        strText = GetIniFileProperty(strExchange, "", "Summary", m.strDataFile)
        If Len(strText) > 0 Then
            aData.SplitFields strText, "|"
            If aData.Size = 3 Then
                With oGrid
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = strExchange
                                        
                    dTotal = Val(aData(0)) + Val(aData(1)) + Val(aData(2))
                    'advanced issues
                    dPercent = aData(0) / dTotal * 100
                    .Cell(flexcpFloodPercent, .Rows - 1, 1) = dPercent
                    .TextMatrix(.Rows - 1, 1) = FormatValue(aData(0), dTotal, dPercent)
                    'declined issues
                    dPercent = aData(1) / dTotal * 100
                    .Cell(flexcpFloodPercent, .Rows - 1, 2) = dPercent
                    .TextMatrix(.Rows - 1, 2) = FormatValue(aData(1), dTotal, dPercent)
                    'unchanged issues
                    dPercent = aData(2) / dTotal * 100
                    .Cell(flexcpFloodPercent, .Rows - 1, 3) = dPercent
                    .TextMatrix(.Rows - 1, 3) = FormatValue(aData(2), dTotal, dPercent)
                    'total issues
                    .TextMatrix(.Rows - 1, 4) = dTotal
                    
                    If aData(0) > aData(1) Then
                        .Cell(flexcpFloodColor, .Rows - 1, 1) = m.nAdvancedColor
                        .Cell(flexcpFloodColor, .Rows - 1, 2) = m.nAdvancedColor
                        .Cell(flexcpFloodColor, .Rows - 1, 3) = m.nAdvancedColor
                    Else
                        .Cell(flexcpFloodColor, .Rows - 1, 1) = m.nDeclinedColor
                        .Cell(flexcpFloodColor, .Rows - 1, 2) = m.nDeclinedColor
                        .Cell(flexcpFloodColor, .Rows - 1, 3) = m.nDeclinedColor
                    End If
                End With
            End If
        End If
    Next

End Sub

Private Function FormatValue(ByVal dValue#, ByVal dTotal#, ByVal dPercent#) As String

    Dim strText
    
    Select Case m.nValueStyle
        Case 0:
            strText = Format(dPercent, "0.##") & "%"
        Case 1:
            strText = Str(Int(dValue))
        Case 2:
            strText = Str(Int(dValue)) & " (" & Format(dPercent, "0.##") & "%)"
        Case Else:
            strText = Format(dPercent, "0.##") & "%"
    End Select
    
    FormatValue = strText

End Function

Public Property Get AdvancedColor() As Long
    AdvancedColor = m.nAdvancedColor
End Property

Public Property Let AdvancedColor(ByVal nColor&)
    m.nAdvancedColor = nColor
End Property

Public Property Get DeclinedColor() As Long
    DeclinedColor = m.nDeclinedColor
End Property

Public Property Let DeclinedColor(ByVal nColor&)
    m.nDeclinedColor = nColor
End Property

Public Property Get ValueStyle() As Long
    ValueStyle = m.nValueStyle
End Property

Public Property Let ValueStyle(ByVal nStyle&)
    m.nValueStyle = nStyle
End Property
