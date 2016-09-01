VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{1BE65FA0-CBF9-11D2-BBC7-00104B9E0792}#2.0#0"; "sstbars2.ocx"
Begin VB.Form frmCotReport 
   Caption         =   "Form1"
   ClientHeight    =   4830
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   6660
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   6660
   StartUpPosition =   3  'Windows Default
   Begin VSFlex7LCtl.VSFlexGrid fgReport 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      _cx             =   11245
      _cy             =   4895
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   1
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin ActiveToolBars.SSActiveToolBars tbToolbar 
      Left            =   5280
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   131083
      ToolBarsCount   =   1
      ToolsCount      =   6
      DisplayContextMenu=   0   'False
      Tools           =   "frmCotReport.frx":0000
      ToolBars        =   "frmCotReport.frx":2706
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Begin VB.Menu mnuFields 
         Caption         =   "Edit &Fields"
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "Edit &Settings"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print Report"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangeFont 
         Caption         =   "&Change Font"
      End
   End
End
Attribute VB_Name = "frmCotReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmCotReport.frm
'' Description: Calculates and shows the COT Report to the user
''
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    strSymGrpId As String
    alSymbolIds As cGdArray
    astrSpreadSyms As cGdArray
    strSave As String
    astrEnglish As cGdArray
    aValues As cGdTree
    dEndDate As Double
    lMaxWeeks As Long
    lTotalWidth As Long
    
    FieldTable As cGdTable
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetCodedText
'' Description: Given an English version of a function call, hand back the
''              coded text for that expression
'' Inputs:      English function call to translate
'' Returns:     Coded text version of the expression
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetCodedText(ByVal strEnglishText As String) As String
On Error GoTo ErrSection:
   
    Dim lIndex As Long                  ' Index for a for loop
    Dim strChk As String                ' Temporary string for input checking
    Dim strNotKnown As String           ' Inputs not recognized
    Dim bExtraInputs As Boolean         ' Are there unrecognized inputs?
    Dim Expr As cExpression             ' Expression object for translation
    Dim Inputs As cInputs               ' Inputs collection in the expression
 
    If Len(Trim(strEnglishText)) = 0 Then
        Exit Function
    End If
 
    ' Verify the expression to get the coded text from the english text
    Set Expr = New cExpression
    With Expr
        .PortfolioNavigator = False
        .Functions = g.Functions
        .ValidateFunctionRule strEnglishText
    End With
        
    ' Check to make sure there are no unrecognized inputs
    bExtraInputs = False
    strNotKnown = ""
    If Not Expr.Inputs Is Nothing Then
        Set Inputs = Expr.Inputs
        For lIndex = 1 To Expr.Inputs.Count
            strChk = UCase(Inputs.Item(lIndex).ParmName)
            If strChk <> "WEEKLY" And _
                    strChk <> "GC" And _
                    strChk <> "MARKET1" Then
                strNotKnown = strNotKnown & "|" & Inputs.Item(lIndex).ParmName
                bExtraInputs = True
            End If
        Next
    End If
    
    If bExtraInputs Then
        Err.Raise vbObjectError + 1000, , "Unrecognized items in expression:|" & strNotKnown & "|"
    Else
        GetCodedText = Expr.CodedText
    End If
    
    
ErrExit:
    Set Expr = Nothing
    Set Inputs = Nothing
    Exit Function

ErrSection:
    RaiseError "frmCotReport.GetCodedText", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalcValues
'' Description: Recalculate values for all symbols and expressions
'' Inputs:      Do a weekly calculation?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CalcValues(Optional ByVal bWeekly As Boolean = False) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim hArray As Long                  ' Array handle
    Dim lFromDate As Long               ' Date to load data from
    Dim lLastDateOfData As Long         ' Last date of data
    Dim rc As Long                      ' Return code from function calls
    Dim lSymbol As Long                 ' Index into a for loop
    Dim strSymbol As String             ' Symbol to get data for
    Dim lSymbolID As Long               ' Symbol ID to get data for
    Dim dPrice As Double                ' Price from the results array
    Dim dPrev As Double                 ' Previous value
    Dim strCodedText As String          ' Coded text for an English expression
            
    Dim Bars As New cGdBars             ' Data for the main market
    Dim GC As New cGdBars               ' Data for Gold 67 contract
    Dim Weekly As New cGdBars           ' Data for weekly bars of main market
    Dim GCWeekly As New cGdBars         ' Data for weekly Gold
    
    Dim astrParms As New cGdArray       ' Paramaters array for the engine
    Dim astrBarNames As New cGdArray    ' Array of bar names
    Dim aScanExpr As New cGdArray       ' Array of coded text expressions
    Dim aArrayOfResults As New cGdArray ' Array of results
    Dim aArrayOfBars As New cGdArray    ' Array of bars structures
    Dim aScanArrays As New cGdArray     ' Array of results for a scan
    Dim aScanPrev As New cGdArray       ' Array of results for prev
    Dim aIsSetup As New cGdArray
    Dim adTemp As New cGdArray          ' Temporary array
    Dim astrCodedText As New cGdArray   ' Array of coded text expressions
    
    Dim lPercent As Long
    
'frmTest.AddList "Starting COT"
    
    ' Initialize the status form
    frmStatus.Status = eStatus_Initialized
    frmStatus.Status = eStatus_Running
    frmStatus.AddDetail "Calculating COT Report"
    frmStatus.SetTitle "Calculating COT Report"
    frmStatus.UpdateProgress "Initializing"
    
    ' Create the arrays
    aScanExpr.Create eGDARRAY_Strings
    aScanArrays.Create eGDARRAY_Longs
    aScanPrev.Create eGDARRAY_Longs
    aArrayOfResults.Create eGDARRAY_Longs
    aIsSetup.Create eGDARRAY_TinyInts, 0, 0
    
    ' Create the coded text array
    astrCodedText.Create eGDARRAY_Strings, m.astrEnglish.Size
    For lIndex = 0 To m.astrEnglish.Size - 1
        astrCodedText(lIndex) = GetCodedText(m.astrEnglish(lIndex))
        DebugLog astrCodedText(lIndex)
    Next lIndex
    
'frmTest.AddList "Coded text created", True
    
    ' Find out the minimum number of days required
'    lNumDays = RunExpression(astrCodedText, g.SymbolPool.SymbolIDforSymbol("SP-067"), bWeekly)
    
    ' Set up the expressions, values, and results arrays
    For lIndex = 0 To m.astrEnglish.Size - 1
        With m.astrEnglish
            strCodedText = astrCodedText(lIndex)
            If Len(Trim(strCodedText)) > 0 Then
                ' Get values array handle, clear array (so no longer
                ' a const array), pre-size array, and store handle
                Set adTemp = New cGdArray
                adTemp.Create eGDARRAY_Doubles
                m.aValues.Add adTemp, m.FieldTable(1, lIndex)
                
                If InStr(UCase(m.FieldTable(1, lIndex)), "CHANGE") = 0 Then
                    aScanExpr.Add Trim(strCodedText)
                    ' see if this is a "setup" type
                    If InStr(UCase(m.FieldTable(1, lIndex)), "SETUP") > 0 Then
                        aIsSetup.Add True
                    Else
                        aIsSetup.Add False
                    End If
                    
                    'hArray = m.aValues(Parse(m.astrEnglish(lIndex), "(", 1)).ArrayHandle
                    hArray = m.aValues(m.FieldTable(1, lIndex)).ArrayHandle
                    gdClear hArray, True
                    gdSetSize hArray, m.alSymbolIds.Size, False
                    aScanArrays.Add hArray
                    
                    ' Create a temporary result array to be used
                    ' by the expression evaluator
                    hArray = gdCreateArray(eGDARRAY_Doubles, 0)
                    aArrayOfResults.Add hArray
                Else
                    hArray = m.aValues(m.FieldTable(1, lIndex)).ArrayHandle
                    gdClear hArray, True
                    gdSetSize hArray, m.alSymbolIds.Size, False
                    aScanPrev.Add hArray
                End If
            End If
        End With
        
        If frmStatus.Status >= eStatus_Aborting Then
            aScanExpr.Size = 0
            Exit For
        End If
    Next lIndex

    ' Calc FromDate, adjusting for weekends and holidays
    ' (need to fudge a little to the safe side)
    lLastDateOfData = m.dEndDate 'LastDailyDownload
    'lFromDate = lLastDateOfData - Int(lNumDays * 1.46 + 0.5) - 2
    lFromDate = lLastDateOfData - (m.lMaxWeeks + 3) * 7

'frmTest.AddList "Before InitExpr", True

    If aScanExpr.Size > 0 Then
        ' Init the expression evaluator with list of scan expressions
        astrBarNames(0) = "Market1"
        astrBarNames(1) = "Weekly"
        astrBarNames(2) = "GC"
        astrParms(0) = "CotReportCalc"
        If Not SetupExpressions(astrParms, astrBarNames, aScanExpr) Then
            Err.Raise vbObjectError + 1000, , "An error exists in an expression."
        End If

'frmTest.AddList "InitExpr", True
    
        aArrayOfBars.Create eGDARRAY_Longs
        For lSymbol = 0 To m.alSymbolIds.Size - 1
            lSymbolID = m.alSymbolIds(lSymbol)
            strSymbol = g.SymbolPool.SymbolForID(lSymbolID)
            
            Bars.Size = 0
            SetBarProperties Bars, lSymbolID
            
            If lSymbolID <> 0 Then
                If Not DM_GetBars(Bars, lSymbolID, 0, lFromDate, m.dEndDate, , , , False) Then
                    Bars.Size = 0
                End If
            End If

            ' Load the spread symbol if necessary
            If m.astrSpreadSyms(lSymbol) <> GC.Prop(eBARS_Symbol) Then
                DM_GetBars GC, m.astrSpreadSyms(lSymbol), , lFromDate, m.dEndDate, , , , False
                If GC.Size > 0 Then GCWeekly.BuildBars "Weekly", GC.BarsHandle
            End If
    
            If Bars.Size > 0 Then
                Weekly.BuildBars "Weekly", Bars.BarsHandle
                If bWeekly Then
                    Bars.BuildBars "Weekly"
                    GC.BuildBars "Weekly"
                End If
                
                aArrayOfBars.Num(0) = Bars.BarsHandle '(in case changed)
                aArrayOfBars.Num(1) = Weekly.BarsHandle
                aArrayOfBars.Num(2) = GC.BarsHandle
                
'frmTest.AddList strSymbol & " Data", True
                ' Run engine to evaluate expressions for this symbol
                astrParms.Size = 1
                rc = RunExpressions(astrParms.ArrayHandle, _
                    astrBarNames.ArrayHandle, aArrayOfBars.ArrayHandle, _
                    aArrayOfResults.ArrayHandle, ByVal 0&, ByVal 0&)
'frmTest.AddList strSymbol & " RunExpr", True
                If rc = 0 Then
                    ' set current value for each expression
                    For lIndex = 0 To aScanArrays.Size - 1
                        ' get most recent value
                        hArray = aArrayOfResults.Num(lIndex)
                        If Not bWeekly Then
                            dPrice = gdGetNum(hArray, Bars.Size - 1)
                            dPrev = gdGetNum(hArray, Bars.Size - 2)
                        Else
                            dPrice = gdGetNum(hArray, Weekly.Size - 1)
                            dPrev = gdGetNum(hArray, Weekly.Size - 2)
                        End If

                        ' store into the scan's array for this symbol
                        hArray = aScanArrays.Num(lIndex)
                        If dPrice = kNullData Then dPrice = gdNullValue(hArray)
                        gdSetNum hArray, lSymbol, dPrice
                        
                        ' store into the change array for this symbol
                        hArray = aScanPrev.Num(lIndex)
                        If dPrice = kNullData Or dPrev = kNullData Then
                            dPrice = gdNullValue(hArray)
                        ElseIf aIsSetup.Num(lIndex) = 0 Then
                            dPrice = dPrice - dPrev
                        ElseIf Sgn(dPrice) = Sgn(dPrev) Then
                            ' show setup change in the current direction
                            dPrice = Abs(RoundNum(dPrice)) - Abs(RoundNum(dPrev))
                        Else '(need to handle differently if crossed 0 line)
                            dPrice = Abs(RoundNum(dPrice) - RoundNum(dPrev))
                        End If
                        gdSetNum hArray, lSymbol, dPrice
                    Next lIndex
                End If
            End If
            
            If frmStatus.Status >= eStatus_Aborting Then Exit For
            
            lPercent = (lSymbol / m.alSymbolIds.Size) * 100
            frmStatus.UpdateProgress "Calculating", lPercent
                        
'frmTest.AddList strSymbol & " Done", True
                        
            'we'll yield to other threads only every 1/2 second
            Sleep -0.5
        Next lSymbol
        
        ' clear the expression evaluator when done with it
        SetupExpressions astrParms '(clear expressions)
    End If
    
'frmTest.AddList "Done", True
    
   ' Destroy all the temporary result arrays
    For lIndex = 0 To aArrayOfResults.Size - 1
        gdDestroyArray aArrayOfResults(lIndex)
    Next lIndex
    aArrayOfResults.Size = 0
    
    Set adTemp = Nothing
    Set astrCodedText = Nothing
    
    If frmStatus.Status = eStatus_Aborted Or frmStatus.Status = eStatus_Aborting Then
        frmStatus.Status = eStatus_Aborted
        CalcValues = False
    Else
        CalcValues = True
        frmStatus.Status = eStatus_Completed
        frmStatus.AddDetail "Finished"
    End If
               
ErrExit:
    Exit Function
    
ErrSection:
    If frmStatus.Status = eStatus_Running Then frmStatus.Status = eStatus_Aborted
    RaiseError "frmCotReport.CalcValues", eGDRaiseError_Raise
        
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgReport_AfterEdit
'' Description: After the user has edited the "Spread Symbol" column, figure
''              out the new symbol if necessary and recalc the WillVal if
''              necessary
'' Inputs:      Row and Column changed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgReport_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim strValue As String              ' Symbol back from symbol selector
    Dim aSymbols As New cGdArray        ' Symbols returned from symbol selector
    Dim astrExpr As New cGdArray        ' Expressions to recalculate
    Dim adValues As New cGdArray        ' Values returned from recalc
    
    ' If the redraw is off, don't bother to do anything
    If fgReport.Redraw = flexRDNone Then Exit Sub
    
    With fgReport
        ' Only come in here if the Spread Symbol column has been edited
        If Col = 27 Then
            .TextMatrix(Row, Col) = .EditText
            ' If the user chose default, use either Gold or Bonds
            If LCase(.TextMatrix(Row, Col)) = "default" Then
                If .TextMatrix(Row, 0) = "GC-067" Then
                    .TextMatrix(Row, Col) = "TQ-067"
                Else
                    .TextMatrix(Row, Col) = "GC-067"
                End If
                
            ' If the user chose lookup, bring up the symbol selector
            ElseIf .Cell(flexcpText, Row, Col) = "< Lookup >" Then
                Set aSymbols = frmSymbolSelector.ShowMe(m.strSave, False)
                strValue = aSymbols(0)
                If Len(strValue) = 0 Then
                    .TextMatrix(Row, Col) = m.strSave
                Else
                    .TextMatrix(Row, Col) = strValue
                End If
            End If
            
            ' If the value changed, recalculate the WillVal
            If .TextMatrix(Row, Col) <> m.strSave Then
                astrExpr.Create eGDARRAY_Strings, 1
                adValues.Create eGDARRAY_Doubles, 1
                
                astrExpr(0) = GetCodedText(m.astrEnglish(25))
                astrExpr(1) = GetCodedText(m.astrEnglish(26))
                RunExpression astrExpr, g.SymbolPool.SymbolIDforSymbol(.TextMatrix(Row, 0)), True, 1200, .TextMatrix(Row, Col), adValues
                If adValues(0) <> gdNullValue(adValues.ArrayHandle) Then
                    .TextMatrix(Row, Col - 2) = Format(adValues(0), "#0.0")
                    .TextMatrix(Row, Col - 1) = Format(adValues(1), "#0.0")
                Else
                    .TextMatrix(Row, Col - 2) = ""
                    .TextMatrix(Row, Col - 1) = ""
                End If
            End If
            m.strSave = ""
        End If
    End With
    
ErrExit:
    ' Clean up
    Set aSymbols = Nothing
    Set astrExpr = Nothing
    Set adValues = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmCotReport.fgReport.AfterEdit", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgReport_AfterRowColChange
'' Description: After a row or column change, call edit cell
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgReport_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    If fgReport.Visible Then fgReport.EditCell

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCotReport.fgReport.AfterRowColChange", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgReport_AfterSort
'' Description: After the grid is sorted, recolor the background colors
'' Inputs:      Column Sorted, Order Sorted
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgReport_AfterSort(ByVal Col As Long, Order As Integer)
On Error GoTo ErrSection:

    SetBackColors fgReport

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCotReport.fgReport_AfterSort"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgReport_AfterUserResize
'' Description: After the user resizes the columns, save off the widths
'' Inputs:      Row and Column changed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgReport_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index for a for loop
    
    With fgReport
        m.lTotalWidth = 0&
        For lIndex = 0 To .Cols - 1
            m.FieldTable(7, lIndex) = Str(.ColWidth(lIndex))
            m.lTotalWidth = m.lTotalWidth + .ColWidth(lIndex)
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCotReport.fgReport_AfterUserResize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgReport_BeforeEdit
'' Description: Only allow the user to edit the Spread Symbol column
'' Inputs:      Row and Column to be edited, Whether or not to Cancel the edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgReport_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    ' Initialize the Combo List for the grid
    fgReport.ComboList = ""

    ' Only allow the user to edit the Spread Symbol column
    If Col <> 27 Then
        Cancel = True
    Else
        ' Give the user the Combo drop down
        With fgReport
            m.strSave = .TextMatrix(Row, Col)
            .ComboList = "|default|GC-067|TQ-067|US-067|TY-067|DX-067|SP-067|$DJIA|< Lookup >"
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCotReport.fgReport.BeforeEdit", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub fgReport_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
On Error GoTo ErrSection:

    Dim lMouseRow As Long
    Dim lMouseCol As Long

    With fgReport
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If Button = vbRightButton Then
            If lMouseRow >= .FixedRows And lMouseRow < .Rows Then
                .RowSel = lMouseRow
                If Not .IsSelected(lMouseRow) Then .Row = lMouseRow
            End If
            
            Cancel = True
            PopupMenu mnuPopUp
        End If
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCotReport.fgReport.BeforeMouseDown", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgReport_ComboCloseUp
'' Description: After the user closes the combo, force the AfterEdit to happen
'' Inputs:      Row and Column edited, Whether to force a FinishEdit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgReport_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
On Error GoTo ErrSection:

    FinishEdit = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCotReport.fgReport.ComboCloseUp", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgReport_DblClick
'' Description: When the user double clicks on the grid, set the active chart
''              to the symbol in the row that they double clicked in
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgReport_DblClick()
On Error GoTo ErrSection:

    Dim lRow As Long
    
    With fgReport
        lRow = .MouseRow
        
        If lRow >= .FixedRows Then
            .Row = lRow
            SetActiveChartSymbol .TextMatrix(lRow, 0)
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCotReport.fgReport.DblClick", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    fgReport_MouseMove
'' Description: When the user moves the mouse over a column, show the tool
''              tip text for that column
'' Inputs:      Button Pressed, Shift/Ctrl/Alt Status, Mouse Location
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub fgReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim lMouseRow As Long
    Dim lMouseCol As Long

    With fgReport
        lMouseRow = .MouseRow
        lMouseCol = .MouseCol
        
        If lMouseRow < .FixedRows And lMouseRow >= 0 Then
            If .TextMatrix(0, lMouseCol) <> .TextMatrix(1, lMouseCol) Then
                Select Case UCase(.TextMatrix(1, lMouseCol))
                    Case "CHANGE"
                        .ToolTipText = SORT_BY_PREFIX & "Change in " & .TextMatrix(0, lMouseCol)
                    Case "IDX"
                        .ToolTipText = SORT_BY_PREFIX & .TextMatrix(0, lMouseCol) & " Index"
                    Case "IDX CHANGE"
                        .ToolTipText = SORT_BY_PREFIX & "Change in " & .TextMatrix(0, lMouseCol) & " Index"
                    Case "NOW", "VALUE"
                        .ToolTipText = SORT_BY_PREFIX & .TextMatrix(0, lMouseCol)
                    Case "SETUP"
                        .ToolTipText = SORT_BY_PREFIX & .TextMatrix(0, lMouseCol) & " Setup"
                    Case "SETUP CHG"
                        .ToolTipText = SORT_BY_PREFIX & "Change in " & .TextMatrix(0, lMouseCol) & " Setup"
                    Case "SYMBOL"
                        .ToolTipText = SORT_BY_PREFIX & .TextMatrix(0, lMouseCol) & " Symbol"
                    Case Else
                        .ToolTipText = SORT_BY_PREFIX & .TextMatrix(0, lMouseCol) & " " & .TextMatrix(1, lMouseCol)
                End Select
            Else
                .ToolTipText = SORT_BY_PREFIX & .TextMatrix(1, lMouseCol)
            End If
        ElseIf lMouseRow >= .FixedRows And lMouseRow < .Rows Then
            .ToolTipText = m.FieldTable(6, fgReport.MouseCol)
        Else
            .ToolTipText = ""
        End If
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCotReport.fgReport.MouseMove", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Deactivate
'' Description: When the form deactivates, set the previous form to me
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Deactivate()
On Error GoTo ErrSection:

    SetPrevActiveForm Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCotReport.Form.Deactivate", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrSection:

    If KeyCode = vbKeyF1 Then
        KeyCode = 0
        g.Help.ShowF1Help Me
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCotReport.Form.KeyDown", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: When the form loads, center it and set the caption
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Dim strText As String

    g.Styler.StyleForm Me
    
    mnuPopUp.Visible = False

    strText = GetIniFileProperty("COTReportPlacement", "", "Forms", g.strIniFile)
    
    If strText = "" Then
        Me.Width = 9900
        CenterTheForm Me
    Else
        SetFormPlacement Me, strText, "LHTW"
    End If
    
    Me.Icon = Picture16(ToolbarIcon("ID_COTReport"), , True)
    Me.Caption = "COT Report"
    
    strText = "Classic"
    If g.nTbIconStyle = 1 Then
        If g.nColorTheme = kDarkThemeColor Then
            strText = "Light"
        Else
            strText = "Dark"
        End If
    End If
    With tbToolbar
        .Tools("ID_Fields").Picture = g.CoreBridge.ImgListToolbarExt(strText, ToolbarIcon("ID_SymbolGrid"), "", 16).ExtractIcon
        .Tools("ID_Settings").Picture = g.CoreBridge.ImgListToolbarExt(strText, ToolbarIcon("ID_Settings"), "", 16).ExtractIcon
        .Tools("ID_TextInc").Picture = g.CoreBridge.ImgListToolbarExt(strText, ToolbarIcon("ID_TextIncrease"), "", 16).ExtractIcon
        .Tools("ID_TextDec").Picture = g.CoreBridge.ImgListToolbarExt(strText, ToolbarIcon("ID_TextDecrease"), "", 16).ExtractIcon
        .Tools("ID_Print").Picture = g.CoreBridge.ImgListToolbarExt(strText, ToolbarIcon("ID_Print"), "", 16).ExtractIcon
        .Tools("ID_Close").Picture = Picture16(ToolbarIcon("kCancel"))
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCotReport.Form.Load", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Initialize and show the form
'' Inputs:      Symbol Group ID, Array of English expressions, Ending Date
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMe(ByVal strSymGrpId As String, ByVal FieldTable As cGdTable, ByVal dEndDate As Double, ByVal lMaxWeeks As Long)
On Error GoTo ErrSection:

    Dim strLast As String               ' Settings from the last run
    Dim bLoadGrid As Boolean            ' Should we load grid from last saved?

    Screen.MousePointer = vbHourglass
    
    ' Get the settings from the last run
    strLast = GetIniFileProperty("LastRun", "", "LastRun", AddSlash(App.Path) & "CotRpt.INI")

    ' Initialize the necessary local arrays
    Set m.alSymbolIds = New cGdArray
    m.alSymbolIds.Create eGDARRAY_Longs
    Set m.astrSpreadSyms = New cGdArray
    m.astrSpreadSyms.Create eGDARRAY_Strings
    Set m.astrEnglish = New cGdArray
    m.astrEnglish.Create eGDARRAY_Strings
    Set m.aValues = New cGdTree
    m.dEndDate = dEndDate
'    Set m.astrEnglish = astrEnglishText
    m.strSymGrpId = strSymGrpId
    
    Set m.FieldTable = FieldTable
    Set m.astrEnglish = m.FieldTable.FieldArray(2)
    
    m.lMaxWeeks = lMaxWeeks
    
    Me.Caption = "COT Report as of " & DateFormat(m.dEndDate)
    
    ' Get the Symbol IDs from the symbol pool for the given Symbol Group
    GetSymbolIDs
    
    ' If the settings are different, recalc the table otherwise just load it
    If Trim(Str(m.dEndDate)) & ";" & m.strSymGrpId & ";" & m.astrEnglish.JoinFields(";") <> strLast Or Not FileExist(AddSlash(App.Path) & "CotData.GRD") Then
        bLoadGrid = False
        If ProcessIsBusy Then
            Unload Me
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        If Not CalcValues(True) Then
            Unload Me
            Screen.MousePointer = vbDefault
            frmCotSettings.ShowMe
            Exit Sub
        End If
    Else
        bLoadGrid = True
    End If
    
    'JM 12-18-2015: need to call this here because the grids are getting loaded before showing the form
    FixFormControls Me, ALT_GRID_ROW_COLOR
    ' Load up the grid with the appropriate values
    InitGrid bLoadGrid
    
    Screen.MousePointer = vbDefault
    
    ' Show the form
    ShowForm Me, eForm_Nonmodal, frmMain, , ALT_GRID_ROW_COLOR

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCotReport.ShowMe", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: When the form is resized, resize the grid to the full form size
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    With fgReport
'        .Move .Left, .Top, Me.ScaleWidth - .Left * 2, _
'                    Me.ScaleHeight - fraButtons.Height - .Top * 3
    
    
        .Move .Left, .Top, Me.ScaleWidth - .Left * 2, _
                    Me.ScaleHeight - .Top * 3
    End With
    
'    With fraButtons
'        .Move Me.ScaleWidth / 2 - .Width / 2, fgReport.Height + fgReport.Top * 2
'    End With
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: When the form is unloaded, save off the settings and the grid
'' Inputs:      Whether or not to cancel the unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index for a for loop
    Dim strIniFile As String
    
    strIniFile = AddSlash(App.Path) & "CotRpt.INI"

    ' Don't save anything if the user aborted the process
    If frmStatus.Status <> eStatus_Aborted And frmStatus.Status <> eStatus_Aborting Then
        ' Save the settings
        SetIniFileProperty "LastRun", Trim(Str(m.dEndDate)) & ";" & m.strSymGrpId & ";" & m.astrEnglish.JoinFields(";"), "LastRun", strIniFile
        
        ' Save the table to the Ini File
        SetIniFileProperty "NumFields", m.FieldTable.NumRecords, "Fields", strIniFile
        For lIndex = 0 To m.FieldTable.NumRecords - 1
            SetIniFileProperty "Field" & lIndex, m.FieldTable(0, lIndex) & ";" & _
                    m.FieldTable(1, lIndex) & ";" & m.FieldTable(2, lIndex) & ";" & _
                    m.FieldTable(3, lIndex) & ";" & m.FieldTable(4, lIndex) & _
                    ";" & m.FieldTable(5, lIndex) & ";" & m.FieldTable(6, lIndex) & _
                    ";" & m.FieldTable(7, lIndex), "Fields", strIniFile
        Next lIndex
        
        ' Save off the grid and the Spread symbols
        With fgReport
            .SaveGrid AddSlash(App.Path) & "CotData.GRD", flexFileData, False
            
            For lIndex = .FixedRows To .Rows - 1
                SetIniFileProperty .TextMatrix(lIndex, 0), .TextMatrix(lIndex, 27), "SpreadSymbols", AddSlash(App.Path) & "CotRpt.INI"
            Next lIndex
        End With
    
        SetIniFileProperty "COTReport", FontToString(fgReport.Font), "Fonts", g.strIniFile
        SetIniFileProperty "COTReportPlacement", GetFormPlacement(Me), "Forms", g.strIniFile
    End If
   
ErrExit:
    ' Clean up after ourselves
    Set m.alSymbolIds = Nothing
    Set m.astrEnglish = Nothing
    Set m.aValues = Nothing
    Set m.astrSpreadSyms = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "frmCotReport.Form.Unload", eGDRaiseError_Show
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetSymbolIDs
'' Description: Get the symbol ID's for the symbol group and set the spread
''              symbols from what was stored
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub GetSymbolIDs()
On Error GoTo ErrSection:

    Dim lFieldNum As Long               ' Field number for the symbol group
    Dim aSymbols As cGdArray            ' Array of true/false in grid values
    Dim lIndex As Long                  ' Array into a for loop
    Dim lCounter As Long                ' Index into the array
    Dim strTemp As String               ' Temporary string
    Dim lSymbolID As Long               ' Symbol Id for the symbol
    Dim dValue As Double                ' Temporary COT value
    
    ' Get the field number for the symbol group
    lFieldNum = g.SymbolPool.FieldNumForID(m.strSymGrpId)
    Set aSymbols = g.SymbolPool.ArrayTable.FieldArray(lFieldNum)
    
    ' Add the symbols from that group into this one
    For lIndex = 0 To aSymbols.Size - 1
        If Abs(aSymbols(lIndex)) = 1 Then
            lSymbolID = g.SymbolPool.SymbolID(lIndex)
            'If DM_GetSnapFromHistory(lSymbolID, "COT_1", m.dEndDate, dValue) Then
                m.alSymbolIds(lCounter) = lSymbolID
                If g.SymbolPool.Symbol(lIndex) <> "GC-067" Then
                    strTemp = GetIniFileProperty(aSymbols(lIndex), "GC-067", "SpreadSymbols", AddSlash(App.Path) & "CotRpt.INI")
                Else
                    strTemp = GetIniFileProperty(aSymbols(lIndex), "TQ-067", "SpreadSymbols", AddSlash(App.Path) & "CotRpt.INI")
                End If
                m.astrSpreadSyms(lCounter) = strTemp
                lCounter = lCounter + 1
            'End If
        End If
    Next lIndex
    
    ' Clean up after ourselves
    Set aSymbols = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCotReport.GetSymbolIDs", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitGrid
'' Description: Initialize and fill in the grid
'' Inputs:      Take the values from the saved file? (or from m.aValues)
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitGrid(ByVal bFromFile As Boolean)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index for a for loop
    Dim lCol As Long                    ' Index for a for loop
    Dim lValue As Long                  ' Value for the "Market Setup"
    Dim strFont As String
    Dim dValue As Double

    With fgReport
        ' Main grid settings
        .Redraw = flexRDNone
        .Clear
        
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExSortShow
        .AllowUserResizing = flexResizeColumns
        .ExtendLastCol = True
        .SelectionMode = flexSelectionListBox
        .AllowBigSelection = False
        .ScrollTrack = True
        .SheetBorder = RGB(128, 128, 128)
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        
        ' Column and row settings
        .Cols = m.FieldTable.NumRecords '10
        .FixedCols = 0
        .Rows = m.alSymbolIds.Size + 2 '1
        .FixedRows = 2 '1
        
        If bFromFile Then
            .LoadGrid AddSlash(App.Path) & "CotData.GRD", flexFileData, False
            For lCol = 0 To .Cols - 1
                If InStr(UCase(m.FieldTable(1, lCol)), "CHANGE") <> 0 Then
                    For lIndex = .FixedRows To .Rows - 1
                        dValue = ValOfText(.TextMatrix(lIndex, lCol))
                        If dValue > 0 Then
                            .Cell(flexcpForeColor, lIndex, lCol) = QBColor(2)
                        ElseIf dValue < 0 Then
                            .Cell(flexcpForeColor, lIndex, lCol) = vbRed
                        End If
                    Next lIndex
                End If
            Next lCol
        End If
        
        ' Column headers
        m.lTotalWidth = 0&
        For lIndex = 0 To m.FieldTable.NumRecords - 1
            .TextMatrix(0, lIndex) = m.FieldTable(4, lIndex)
            .TextMatrix(1, lIndex) = m.FieldTable(5, lIndex)
            .ColHidden(lIndex) = Not CBool(m.FieldTable(0, lIndex))
            If Len(m.FieldTable(7, lIndex)) > 0 Then
                .ColWidth(lIndex) = Val(m.FieldTable(7, lIndex))
                m.lTotalWidth = m.lTotalWidth + .ColWidth(lIndex)
            End If
        Next lIndex
        .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True
        .MergeCol(0) = True
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterTop
        .Cell(flexcpAlignment, 1, 0, 1, .Cols - 1) = flexAlignCenterTop 'flexAlignLeftTop
        .Cell(flexcpAlignment, 0, 0) = flexAlignLeftCenter

        ' If not from the file, load values from m.aValues
        If Not bFromFile Then
            For lIndex = 0 To m.alSymbolIds.Size - 1
                .TextMatrix(lIndex + .FixedRows, 0) = g.SymbolPool.SymbolForID(m.alSymbolIds(lIndex))
                For lCol = 1 To .Cols - 1
                    If m.FieldTable(1, lCol) <> "Will-Val Symbol" Then
                        'ShowValue Parse(m.FieldTable(2, lCol), "(", 1), lIndex, lIndex + .FixedRows, lCol
                        ShowValue m.FieldTable(1, lCol), lIndex, lIndex + .FixedRows, lCol
                    Else
                        .TextMatrix(lIndex + .FixedRows, lCol) = m.astrSpreadSyms(lIndex)
                    End If
                Next lCol
            Next lIndex
        End If
            
        If .ColHidden(19) = False Then
            .Col = 19
            .Sort = flexSortGenericDescending '(numeric sorting does not differentiate between "0" and blank)
        ElseIf .ColHidden(25) = False Then
            .Col = 25
            .Sort = flexSortGenericDescending '(numeric sorting does not differentiate between "0" and blank)
        End If
        
        FilterGrid
        
        .Col = 1
        .FrozenCols = 1
          
        strFont = GetIniFileProperty("COTReport", "", "Fonts", g.strIniFile)
        If strFont <> "" Then
            FontFromString .Font, strFont
            .Font = .Font '(this is required to trigger the grid to reset itself!)
        End If
        
        If m.lTotalWidth = 0 Then
            .AutoSize 0, .Cols - 1, False, 75
            
            ' make WillVal symbol a little wider
            .ColWidth(29) = .ColWidth(29) + 110
            .ColWidth(32) = .ColWidth(32) + 200
        End If
        .Redraw = flexRDBuffered
    End With
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCotReport.InitGrid", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowValue
'' Description: Show the value if it exists, or blank if it is Null
'' Inputs:      Key into the tree, Indexed item, Row and Column in the grid
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ShowValue(ByVal strKey As String, ByVal lIndex As Long, ByVal lRow As Long, ByVal lCol As Long)
On Error GoTo ErrSection:

    Dim dValue As Double, nNode As Long

    nNode = m.aValues.Index(strKey)
    If nNode <= 0 Then Exit Sub
    dValue = m.aValues(nNode).Item(lIndex)

    With fgReport
        If dValue <> gdNullValue(m.aValues(nNode).ArrayHandle) Then
            Select Case strKey
                Case "Commercials", "Commercials Change", "Large Spec", _
                        "Large Spec Change", "Small Spec", _
                        "Small Spec Change", "Genesis Sentiment", _
                        "Genesis Sentiment Change", "TN Consensus", "TN Consensus Change", _
                        "Genesis Setup Strength Change", "TN Setup Strength Change", _
                        "Genesis Proxy Setup Strength Change", "LW Proxy Setup Strength Change"
                    .TextMatrix(lRow, lCol) = Format(dValue, "#,##0")
                
                Case "Genesis Setup Strength", "TN Setup Strength", "Genesis Proxy Setup Strength", "LW Proxy Setup Strength"
                    If dValue <= 0 Then
                        .TextMatrix(lRow, lCol) = NumStr(Abs(dValue), 3) & " Bullish"
                    Else
                        .TextMatrix(lRow, lCol) = NumStr(dValue, 3) & " Bearish"
                    End If
                
                Case Else
                    .TextMatrix(lRow, lCol) = Format(dValue, "#,##0.0")
            End Select
            
            If InStr(strKey, "Change") <> 0 Then
                If dValue > 0 Then
                    .TextMatrix(lRow, lCol) = "+" & .TextMatrix(lRow, lCol)
                    .Cell(flexcpForeColor, lRow, lCol) = QBColor(2)
                ElseIf dValue < 0 Then
                    .Cell(flexcpForeColor, lRow, lCol) = vbRed
                End If
            End If
        Else
            .TextMatrix(lRow, lCol) = ""
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCotReport.ShowValue", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RunExpression
'' Description: Run expressions to find out the minimum number of days required
'' Inputs:      Expressions, Symbol to run on, Weekly?, Number of days to
''              Load for this trial, Default Spread Symbol, Final Values
'' Returns:     Number of days required to run these expressions
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RunExpression(ByVal astrExpr As cGdArray, ByVal lSymbolID&, _
                                Optional bWeekly As Boolean = False, _
                                Optional ByVal nNumDaysToLoad& = -1&, _
                                Optional ByVal strSpreadSym$ = "GC-067", _
                                Optional adValues As cGdArray = Nothing) As Long
On Error GoTo ErrSection:

    Dim i&, ii&, rc&, d#, hArray&, nAutoDetect&, lTemp&
    Dim nRecord&, nCount&, nStartDate&
    Dim strCodedText$
    Dim Bars As New cGdBars
    Dim Weekly As New cGdBars
    Dim GC As New cGdBars
    
    Dim astrParms As New cGdArray, astrBarNames As New cGdArray
    Dim aScanExpr As New cGdArray, aArrayOfResults As New cGdArray
    Dim aArrayOfBars As New cGdArray
    Dim aScanArrays As New cGdArray
    Dim aMinBarsReq As New cGdArray
    
    Dim iDayOfWeek As Integer
    Dim lIndex As Long
    Dim dPrice As Double
       
    ' Get coded text and handle of values array from each Criteria
    aScanExpr.Create eGDARRAY_Strings
    aScanArrays.Create eGDARRAY_Longs
    aArrayOfResults.Create eGDARRAY_Longs
    aMinBarsReq.Create eGDARRAY_Longs
    
    If Not adValues Is Nothing Then
        adValues.Create eGDARRAY_Doubles, astrExpr.Size
    End If
            
    nRecord = g.SymbolPool.PoolRecForSymbolID(lSymbolID)
    If astrExpr.Size > 0 And nRecord >= 0 Then
        If nNumDaysToLoad = -1& Then
            nStartDate = DateSerial(1998, 1, 1) '0
        Else
            nStartDate = m.dEndDate - Int(nNumDaysToLoad * 1.46 + 0.5) - 2
        End If
        
        aArrayOfBars.Create eGDARRAY_Longs
        Bars.Size = 0
            
        If lSymbolID <> 0 Then
            ' load a year's worth of data
            If Not DM_GetBars(Bars, lSymbolID, , nStartDate, 0, , , , False) Then
                Bars.Size = 0
            End If
        End If
            
        If Bars.Size > 0 Then
            DM_GetBars GC, g.SymbolPool.SymbolIDforSymbol(strSpreadSym), , nStartDate, 0, , , , False
            Weekly.BuildBars "Weekly", Bars.BarsHandle
            If bWeekly = True Then
                Bars.BuildBars "Weekly"
                GC.BuildBars "Weekly"
            End If
            Set aScanExpr = astrExpr
            
            ' create a temporary result array to be used
            ' by the expression evaluator
            For lIndex = 0 To astrExpr.Size - 1
                If astrExpr(lIndex) <> "" Then
                    hArray = gdCreateArray(eGDARRAY_Doubles, Bars.Size)
                Else
                    hArray = 0&
                End If
                aArrayOfResults.Add hArray
            Next lIndex
            
            ' Init the expression evaluator with list of scan expressions
            astrBarNames(0) = "Market1"
            astrBarNames(1) = "Weekly"
            astrBarNames(2) = "GC"
            astrParms(0) = "CotReportRunExp"
            If Not SetupExpressions(astrParms, astrBarNames, aScanExpr) Then
                Err.Raise vbObjectError + 1000, , "An error exists in a Criteria expression"
            End If
    
            ' run engine to evaluate expressions for this symbol
            aArrayOfBars.Num(0) = Bars.BarsHandle '(in case changed)
            aArrayOfBars.Num(1) = Weekly.BarsHandle
            aArrayOfBars.Num(2) = GC.BarsHandle
            astrParms.Size = 1
            rc = RunExpressions(astrParms.ArrayHandle, _
                astrBarNames.ArrayHandle, aArrayOfBars.ArrayHandle, _
                aArrayOfResults.ArrayHandle, aMinBarsReq.ArrayHandle, ByVal 0&)
            
            If rc = 0 Then
                ' see if last value is not null
                If aMinBarsReq.Size > 0 Then
                    For lIndex = 0 To aMinBarsReq.Size - 1
                        If aMinBarsReq(lIndex) < Bars.Size Then
                            lTemp = aMinBarsReq(lIndex) + 1
                            If Not bWeekly And InStr(astrExpr(lIndex), "~07006WEEKLY") <> 0 Then
                                If lTemp = 0 Then
                                    lTemp = 5
                                Else
                                    ' figure number of daily bars for full weeks
                                    d = Bars(eBARS_DateTime, lTemp - 1) - Bars(eBARS_DateTime, 0)
                                    lTemp = Int((d + 6) / 7) * 5
                                End If
                            End If
                        End If
                        If lTemp > nAutoDetect Then
                            nAutoDetect = lTemp
                        End If
                    Next lIndex
                    
                    If Not adValues Is Nothing Then
                        For lIndex = 0 To adValues.Size - 1
                            hArray = aArrayOfResults.Num(lIndex)
                            dPrice = gdGetNum(hArray, Bars.Size - 1)
                            If dPrice = kNullData Then dPrice = gdNullValue(hArray)
                            adValues(lIndex) = dPrice
                        Next
                    End If
                Else
                    For lIndex = 0 To aArrayOfResults.Size - 1
                        hArray = aArrayOfResults.Num(lIndex)
                        d = gdGetNum(hArray, gdGetSize(hArray) - 1)
                        If d <> gdNullValue(hArray) Then
                            ' if so, find first non-null item
                            For i = 0 To gdGetSize(hArray) - 1
                                d = gdGetNum(hArray, i)
                                If d <> gdNullValue(hArray) Then
                                    lTemp = i + 1
                                    If InStr(UCase(astrExpr(lIndex)), "~07006WEEKLY") <> 0 Then
                                        gdCopy Bars.ArrayHandle(eBARS_Close), hArray
                                        Bars.BuildBars "Weekly"
                                        For ii = 0 To Bars.Size - 1
                                            If Bars(eBARS_Close, ii) <> gdNullValue(Bars.ArrayHandle(eBARS_Close)) Then
                                                'nAutoDetect = (ii + 1) * 5
                                                lTemp = Int((ii + 4) / 5)
                                                Exit For
                                            End If
                                        Next ii
                                    End If
                                    Exit For
                                End If
                            Next
                        End If
                        If lTemp > nAutoDetect Then
                            nAutoDetect = lTemp
                        End If
                    Next lIndex
                End If
            End If
            ' clear the expression evaluator when done with it
            SetupExpressions astrParms '(clear expressions)
        End If
    End If
    
    ' destroy all the temporary result arrays
    For i = 0 To aArrayOfResults.Size - 1
        gdDestroyArray aArrayOfResults(i)
    Next
    aArrayOfResults.Size = 0
    
    If bWeekly Then nAutoDetect = (nAutoDetect + 1) * 5
    
    RunExpression = nAutoDetect

ErrExit:
    Set Bars = Nothing
    Set Weekly = Nothing
    Set GC = Nothing
    
    Set astrParms = Nothing
    Set astrBarNames = Nothing
    Set aScanExpr = Nothing
    Set aArrayOfResults = Nothing
    Set aArrayOfBars = Nothing
    Set aScanArrays = Nothing
    Set aMinBarsReq = Nothing

    Exit Function
    
ErrSection:
    RaiseError "frmCotReport.RunExpression", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GenerateReport
'' Description: Callback function for the Print Preview
'' Inputs:      Arguments into the Print Preview
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GenerateReport(ByVal vArgs As Variant)
On Error GoTo ErrSection:

    Dim strDefaults As String
    Dim lRow As Long
    Dim lCol As Long
    Dim strText As String
    

    strDefaults = "GRP:CONT067.GRP;7;14;3;3;3;3;-1;2;22;156;GC-067;TQ-067"
    strDefaults = GetIniFileProperty("Defaults", strDefaults, "Defaults", AddSlash(App.Path) & "CotRpt.INI")

    With frmPrintPreview.vp
        .StartDoc
        DoPrintHeader 10
        
        .TextAlign = taCenterMiddle
        .Font.Name = "Times New Roman"
        .Font.Size = 14
        .Font.Bold = True
        '.FontUnderline = True
        .Text = "COT / Market Sentiment Report" & vbLf
        .Font.Size = 12
        .FontUnderline = False
        .Font.Bold = False
        .TextAlign = taLeftMiddle
        .Text = "Date through: " & DateFormat(m.dEndDate) & vbLf
        .Text = "Bars for ADX: " & Parse(strDefaults, ";", 2) & vbLf
        .Text = "Bars for Stochastic: " & Parse(strDefaults, ";", 3) & ", %K: " & Parse(strDefaults, ";", 4) & vbLf '& ", %D: " & Parse(strDefaults, ";", 5) & vbLf
        .Text = "Lookback for COT: " & Parse(strDefaults, ";", 6) & " Years" & vbLf
        If Val(Parse(strDefaults, ";", 8)) = 0 Then
            .Text = "Lookback for Larry Williams Sentiment: " & Parse(strDefaults, ";", 7) & " Years" & vbLf
        Else
            .Text = "Lookback for Genesis Sentiment: " & Parse(strDefaults, ";", 7) & " Years" & vbLf
        End If
        .Text = "Short Term for WillVal: " & Parse(strDefaults, ";", 9) & ", "
        .Text = "Long Term: " & Parse(strDefaults, ";", 10) & ", "
        .Text = "Lookback: " & Parse(strDefaults, ";", 11) & " Weeks" & vbCrLf
        
        fgReport.ExtendLastCol = False
        If frmPrintPreview.GoingToFile Then
            With fgReport
                For lRow = 0 To .Rows - 1
                    strText = ""
                    For lCol = 0 To .Cols - 1
                        If Not .ColHidden(lCol) Then
                            strText = strText & .Cell(flexcpTextDisplay, lRow, lCol) & vbTab
                        End If
                    Next lCol
                    strText = Left(strText, Len(strText) - 1) ' strip the trailing tab
                    frmPrintPreview.vp.Text = strText & vbCrLf
                Next lRow
            End With
        Else
            .RenderControl = fgReport.hWnd
        End If
        fgReport.ExtendLastCol = True
        
        .EndDoc
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCotReport.GenerateReport", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PrintMe
'' Description: Send the grid to the Print Preview
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PrintMe() As Boolean
On Error GoTo ErrSection

    PrintMe = frmPrintPreview.ShowMe("CNV CotReport", frmCotReport, , , , 0.75, 0.75)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCotReport.PrintMe", eGDRaiseError_Raise
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    WillValExists
'' Description: Determine whether or not to show the Will-Val functions
'' Inputs:      None
'' Returns:     True if it exists, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function WillValExists() As Boolean
On Error GoTo ErrSection:

    Dim rs As Recordset
    
    If InStr(g.strAuthorizationString, ",INC,") <> 0 Then
        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctions] " & _
                    "WHERE [FunctionName]='WillVal';", dbOpenDynaset)
        If Not rs.EOF Then
            WillValExists = True
        End If
    End If

ErrExit:
    Set rs = Nothing
    Exit Function

ErrSection:
    RaiseError "frmCotReport.WillValExists", eGDRaiseError_Raise

End Function


Private Sub mnuChangeFont_Click()
On Error GoTo ErrSection:

    ChangeGridFont fgReport, True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCotReport.mnuChangeFont.Click", eGDRaiseError_Show
    Resume ErrExit

End Sub

Private Sub ChangeSettings()
On Error GoTo ErrSection:

    Dim lIndex As Long

    If Not ProcessIsBusy Then
        ' Save the settings
        SetIniFileProperty "LastRun", Trim(Str(m.dEndDate)) & ";" & m.strSymGrpId & ";" & m.astrEnglish.JoinFields(";"), "LastRun", AddSlash(App.Path) & "CotRpt.INI"
        
        ' Save off the grid and the Spread symbols
        With fgReport
            .SaveGrid AddSlash(App.Path) & "CotData.GRD", flexFileData, False
            
            For lIndex = .FixedRows To .Rows - 1
                SetIniFileProperty .TextMatrix(lIndex, 0), .TextMatrix(lIndex, 27), "SpreadSymbols", AddSlash(App.Path) & "CotRpt.INI"
            Next lIndex
        End With
        
        frmCotSettings.ShowMe
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCotReport.ChangeSettings", eGDRaiseError_Raise
    
End Sub

Private Sub ChangeFields()
On Error GoTo ErrSection:

    Dim astrFields As New cGdArray
    Dim lIndex As Long
    Dim strIniFile As String
    
    strIniFile = AddSlash(App.Path) & "CotRpt.INI"
    
    astrFields.Create eGDARRAY_Strings, m.FieldTable.NumRecords
    For lIndex = 0 To m.FieldTable.NumRecords - 1
        astrFields(lIndex) = m.FieldTable(0, lIndex) & vbTab & m.FieldTable(1, lIndex) & _
                vbTab & m.FieldTable(2, lIndex) & vbTab & m.FieldTable(3, lIndex)
    Next lIndex
    
    If frmQuoteBoardFields.ShowMe(astrFields, eQbfMode_CotReport) Then
        m.FieldTable(0, 0) = "True"
        fgReport.ColHidden(0) = False
        For lIndex = 1 To astrFields.Size - 1
            If Parse(astrFields(lIndex), vbTab, 1) = "2" Then
                m.FieldTable(0, lIndex) = "False"
                fgReport.ColHidden(lIndex) = True
            Else
                m.FieldTable(0, lIndex) = "True"
                fgReport.ColHidden(lIndex) = False
            End If
        Next lIndex
        
        FilterGrid
    End If

    Set astrFields = Nothing
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCotReport.ChangeFields", eGDRaiseError_Raise
    
End Sub

Private Sub mnuFields_Click()
On Error GoTo ErrSection:

    ChangeFields

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCotReport.mnuFields.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub mnuPrint_Click()
On Error GoTo ErrSection:

    PrintMe

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCotReport.mnuPrint.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

Private Sub mnuSettings_Click()
On Error GoTo ErrSection:

    ChangeSettings

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCotReport.mnuSettings.Click", eGDRaiseError_Show
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FilterGrid
'' Description: Only show rows in grid that are not blank for a visible setup
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FilterGrid()
On Error GoTo ErrSection:

    Dim lRow As Long                    ' Index into a for loop
    
    With fgReport
        For lRow = .FixedRows To .Rows - 1
            If HasVisibleSetup(lRow) = False And HasVisibleCotData(lRow) = False Then
                .RowHidden(lRow) = True
            Else
                .RowHidden(lRow) = False
            End If
        Next lRow
        
        SetBackColors fgReport
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCotReport.FilterGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasVisibleCotData
'' Description: Does the current row have any visible COT data?
'' Inputs:      Row
'' Returns:     True if has visible COT data, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function HasVisibleCotData(ByVal lRow As Long) As Boolean
On Error GoTo ErrSection:

    Dim lCol As Long                    ' Index into a for loop
    
    HasVisibleCotData = False
    With fgReport
        For lCol = 0 To .Cols - 1
            If InStr(UCase(.TextMatrix(0, lCol)), "COMMERCIALS") <> 0 Then
                If .ColHidden(lCol) = False And Len(.TextMatrix(lRow, lCol)) > 0 Then
                    HasVisibleCotData = True
                    Exit For
                End If
            End If
        Next lCol
    End With

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCotReport.HasCotVisibleData"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasVisibleSetup
'' Description: Does the given row have any visible setup?
'' Inputs:      Row to Check
'' Returns:     True if has visible setup, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function HasVisibleSetup(ByVal lRow As Long) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    HasVisibleSetup = False
    With fgReport
        For lIndex = 0 To .Cols - 1
            If InStr(UCase(.TextMatrix(1, lIndex)), "SETUP") <> 0 Then
                If .ColHidden(lIndex) = False And Len(.TextMatrix(lRow, lIndex)) > 0 Then
                    HasVisibleSetup = True
                    Exit For
                End If
            End If
        Next lIndex
    End With

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmCotReport.HasVisibleSentiment"
    
End Function

Private Sub tbToolbar_ToolClick(ByVal Tool As ActiveToolBars.SSTool)
On Error GoTo ErrSection:

    Dim lRedraw&, lAutoSizeMode&, lNewSize&

    Select Case Tool.ID
        Case "ID_Fields"
            ChangeFields
            
        Case "ID_Close"
            Unload Me
        
        Case "ID_TextInc"
            lNewSize = fgReport.Font.Size + 1
        
        Case "ID_TextDec"
            lNewSize = fgReport.Font.Size - 1
        
        Case "ID_Settings"
            ChangeSettings
        
        Case "ID_Print"
            PrintMe
        
    End Select
    
    If lNewSize > 0 Then
        With fgReport
            lRedraw = .Redraw
            lAutoSizeMode = .AutoSizeMode
            
            fgReport.Font.Size = lNewSize
            .AutoSizeMode = flexAutoSizeColWidth
            .AutoSize 0, .Cols - 1, , 75
            
            .AutoSizeMode = lAutoSizeMode
            .Redraw = lRedraw
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmCotReport.tbToolbar_ToolClick"
    
End Sub

