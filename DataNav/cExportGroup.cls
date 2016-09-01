VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cExportGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cExportGroup.cls
'' Description: Class for exporting a symbol group to a different format
''
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date      Author      Description
'' 07/19/01  D Jarmuth   Created
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Type mPrivate
    strSymbolGroupID As String
    strSymbolGroup As String
    strFormat As String
    strPath As String
    bAutoExport As Boolean
    strPeriod As String
    dStartDate As Double
    dEndDate As Double
    bSplitAdjust As Boolean
    
    ' ASCII Options
    strDateFormat As String
    strDelimiter As String
    bShowHeader As Boolean
    bCustomHeader As Boolean
    strHeaderLine As String
    strFields As String
    strDecimal As String
    strExtension As String
    bSingleFile As Boolean
    strFileName As String
End Type

Private m As mPrivate

Public Property Get SymbolGroupID() As String
    If InStr(m.strSymbolGroupID, ":") = 0 Then
        If Right(m.strSymbolGroupID, 3) = "GRP" Then
            m.strSymbolGroupID = "GRP:" & m.strSymbolGroupID
        Else
            m.strSymbolGroupID = "FIL:" & m.strSymbolGroupID
        End If
    End If
    SymbolGroupID = m.strSymbolGroupID
End Property
Public Property Get SymbolGroup() As String
    SymbolGroup = m.strSymbolGroup
End Property
Public Property Get Format() As String
    Format = m.strFormat
End Property
Public Property Get Path() As String
    Path = m.strPath
End Property
Public Property Get AutoExport() As Boolean
    AutoExport = m.bAutoExport
End Property
Public Property Get Period() As String
    Period = m.strPeriod
End Property
Public Property Get StartDate() As Double
    StartDate = m.dStartDate
End Property
Public Property Get EndDate() As Double
    EndDate = m.dEndDate
End Property
Public Property Get DateFormat() As String
    DateFormat = m.strDateFormat
End Property
Public Property Get Delimiter() As String
    Delimiter = m.strDelimiter
End Property
Public Property Get ShowHeader() As Boolean
    ShowHeader = m.bShowHeader
End Property
Public Property Get CustomHeader() As Boolean
    CustomHeader = m.bCustomHeader
End Property
Public Property Get HeaderLine() As String
    HeaderLine = m.strHeaderLine
End Property
Public Property Get Fields() As String
    Fields = m.strFields
End Property
Public Property Get DecimalChar() As String
    DecimalChar = m.strDecimal
End Property
Public Property Get Extension() As String
    Extension = m.strExtension
End Property
Public Property Get SingleFile() As Boolean
    SingleFile = m.bSingleFile
End Property
Public Property Get FileName() As String
    FileName = m.strFileName
End Property
Public Property Get SplitAdjust() As Boolean
    SplitAdjust = m.bSplitAdjust
End Property

Public Property Let SymbolGroupID(ByVal strValue As String)
    m.strSymbolGroupID = strValue
End Property
Public Property Let SymbolGroup(ByVal strValue As String)
    m.strSymbolGroup = strValue
End Property
Public Property Let Format(ByVal strValue As String)
    m.strFormat = strValue
End Property
Public Property Let Path(ByVal strValue As String)
    m.strPath = strValue
End Property
Public Property Let AutoExport(ByVal bValue As Boolean)
    m.bAutoExport = bValue
End Property
Public Property Let Period(ByVal strValue As String)
    m.strPeriod = strValue
End Property
Public Property Let StartDate(ByVal dValue As Double)
    m.dStartDate = dValue
End Property
Public Property Let EndDate(ByVal dValue As Double)
    m.dEndDate = dValue
End Property
Public Property Let DateFormat(ByVal strValue As String)
    m.strDateFormat = strValue
End Property
Public Property Let Delimiter(ByVal strValue As String)
    m.strDelimiter = strValue
End Property
Public Property Let ShowHeader(ByVal bValue As Boolean)
    m.bShowHeader = bValue
End Property
Public Property Let CustomHeader(ByVal bValue As Boolean)
    m.bCustomHeader = bValue
End Property
Public Property Let HeaderLine(ByVal strValue As String)
    m.strHeaderLine = strValue
End Property
Public Property Let Fields(ByVal strValue As String)
    m.strFields = strValue
End Property
Public Property Let DecimalChar(ByVal strValue As String)
    m.strDecimal = strValue
End Property
Public Property Let Extension(ByVal strValue As String)
    m.strExtension = strValue
End Property
Public Property Let SingleFile(ByVal bValue As Boolean)
    m.bSingleFile = bValue
End Property
Public Property Let FileName(ByVal strValue As String)
    m.strFileName = strValue
End Property
Public Property Let SplitAdjust(ByVal bSplitAdjust As Boolean)
    m.bSplitAdjust = bSplitAdjust
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromString
'' Description: Fill in the class members from a tab delimeted string
'' Inputs:      String from the save file
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromString(ByVal strString As String)
On Error GoTo ErrSection:

    Dim astrProps As New cGdArray
    
    astrProps.Create eGDARRAY_Strings
    astrProps.SplitFields strString, vbTab

    If IsAlpha(astrProps(0)) Then
        m.bAutoExport = CBool(astrProps(0))
    Else
        m.bAutoExport = CBool(CLng(Val(astrProps(0))))
    End If
    m.strSymbolGroupID = astrProps(1)
    m.strFormat = astrProps(2)
    m.strPath = astrProps(3)
    m.strSymbolGroup = astrProps(4)
    m.strPeriod = astrProps(5)
    m.dStartDate = Val(Trim(astrProps(6)))
    m.dEndDate = Val(Trim(astrProps(7)))
    
    ' ASCII Options
    m.strDateFormat = astrProps(8)
    m.strDelimiter = astrProps(9)
    m.bShowHeader = CBool(CLng(Val(astrProps(10))))
    m.bCustomHeader = CBool(CLng(Val(astrProps(11))))
    If Left(astrProps(12), 1) = Chr(34) And Right(astrProps(12), 1) = Chr(34) Then
        m.strHeaderLine = Mid(astrProps(12), 2, Len(astrProps(12)) - 2)
    Else
        m.strHeaderLine = astrProps(12)
    End If
    m.strFields = astrProps(13)
    m.strDecimal = astrProps(14)
    m.strExtension = astrProps(15)
    m.bSingleFile = CBool(CLng(Val(astrProps(16))))
    m.strFileName = astrProps(17)
    
    If Len(astrProps(18)) = 0 Then
        m.bSplitAdjust = True
    Else
        m.bSplitAdjust = CBool(CLng(Val(Trim(astrProps(18)))))
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cExportGroup.FromString", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToString
'' Description: Prepares a tab delimited string from the class members to write
''              to a "save" file
'' Inputs:      None
'' Returns:     Tab delimited string
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToString() As String
On Error GoTo ErrSection:

    Dim astrProps As New cGdArray       ' Array of properties to return
    
    astrProps.Create eGDARRAY_Strings
    
    astrProps.Add Str(CLng(m.bAutoExport))
    astrProps.Add m.strSymbolGroupID
    astrProps.Add m.strFormat
    astrProps.Add m.strPath
    astrProps.Add m.strSymbolGroup
    astrProps.Add m.strPeriod
    astrProps.Add Str(m.dStartDate)
    astrProps.Add Str(m.dEndDate)
    
    ' ASCII Options
    astrProps.Add m.strDateFormat
    astrProps.Add m.strDelimiter
    astrProps.Add Str(CLng(m.bShowHeader))
    astrProps.Add Str(CLng(m.bCustomHeader))
    astrProps.Add Chr(34) & m.strHeaderLine & Chr(34)
    astrProps.Add m.strFields
    astrProps.Add m.strDecimal
    astrProps.Add m.strExtension
    astrProps.Add Str(CLng(m.bSingleFile))
    astrProps.Add m.strFileName
    
    astrProps.Add Str(CLng(m.bSplitAdjust))

    ToString = astrProps.JoinFields(vbTab)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cExportGroup.ToString", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Export
'' Description: Exports the symbol group
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Export(Optional ByVal bRtExport As Boolean = False)
On Error GoTo ErrSection:

    Dim lFieldNum As Long               ' Field number for the symbol group
    Dim lIndex As Long                  ' Index into a for loop
    Dim lSymbolID As Long               ' Symbol ID for the current symbol
    Dim Bars As New cGdBars             ' Bars structure to hold the data
    Dim strFmt As String                ' Format to send
    Dim iPercent As Integer             ' Percent done
    Dim iSave As Integer                ' Last saved percent done
    Dim fhOutput As Integer             ' File handle for the output file
    Dim lBar As Long                    ' Index into a for loop
    Dim strSymbol As String             ' Symbol to be used as the filename
    Dim strDesc As String               ' Description of symbol
    Dim iPos As Integer                 ' Position of something in a string
    Dim strDisplay As String            ' Format for a format command
    Dim strField As String              ' Field to output
    Dim strDelim As String              ' Delimiter to use
    Dim lField As Long                  ' Counter variable
    Dim strTemp As String               ' Temporary string variable
    Dim bSwitchDec As Boolean           ' Do we need to switch decimal?
    Dim bIntraday As Boolean
    Dim dStartDate As Double
    Dim strExtension
    
    Dim astrLine As New cGdArray        ' One line of the ASCII file
    Dim astrOutput As New cGdArray      ' ASCII File to export
    Dim bHeaderShown As Boolean         ' Has the header been shown yet?
    Dim aFields() As String

    If bRtExport Then
        ' special things for Bollinger's auto-export (refreshes every minute)
        dStartDate = LastDailyDownload + 1
        strExtension = "RT"
    Else
        dStartDate = m.dStartDate
        strExtension = m.strExtension
    End If

    If IsIntraday(GetPeriodicity(m.strPeriod)) Then
        bIntraday = True
    End If

    ' build array of fields to do for Ascii format (just those that are checked)
    ReDim aFields(-1 To -1) As String
    If UCase(m.strFormat) = "ASCII" Then
        For lField = 1 To 999
            strField = Parse(m.strFields, "|", lField)
            If Len(strField) = 0 Then
                Exit For
            ElseIf Parse(strField, ";", 1) = flexChecked Then
                If UBound(aFields) < 0 Then
                    ReDim aFields(0) As String
                Else
                    ReDim Preserve aFields(UBound(aFields) + 1) As String
                End If
                strTemp = Parse(strField, ";", 2)
                If UCase(Left(strTemp, 6)) = "TOTAL " Then
                    strTemp = Trim(Mid(strTemp, 6))
                End If
                aFields(UBound(aFields)) = strTemp
                ' add a time field only if exporting intraday data, the date field
                ' has been included, and an "h" is not in the date format string
                If bIntraday And UCase(strTemp) = "DATE" And InStr(UCase(m.strDateFormat), "H") = 0 Then
                    ReDim Preserve aFields(UBound(aFields) + 1) As String
                    aFields(UBound(aFields)) = "Time"
                End If
            End If
        Next
        If UBound(aFields) < 0 Then Exit Sub
        Select Case UCase(Left(m.strDelimiter, 1))
            Case "C"
                strDelim = "," '& Chr(34)
            Case "P"
                strDelim = "|"
            Case "T"
                strDelim = vbTab
            Case "S"
                If InStr(m.strDelimiter, ";") > 0 Then
                    strDelim = ";"
                Else
                    strDelim = " "
                End If
        End Select
        ' build custom header
        If m.bCustomHeader = False Then
            If strDelim = vbTab Then
                m.strHeaderLine = Join(aFields, strDelim)
            Else
                m.strHeaderLine = Chr(34) & Join(aFields, Chr(34) & strDelim & Chr(34)) & Chr(34)
            End If
        End If
        If strDelim = "," Then strDelim = "," & Chr(34)
        For lField = 0 To UBound(aFields)
            aFields(lField) = UCase(aFields(lField))
        Next
    End If
        
    ' Get the field number for the symbol group ID passed in
    lFieldNum = g.SymbolPool.FieldNumForID(m.strSymbolGroupID)
    
    ' init frmStatus
    If Not bRtExport Then
        frmStatus.AddDetail "Exporting " & m.strSymbolGroup & " (" & m.strFormat & ")"
        frmStatus.Status = eStatus_Running
        If bIntraday Then
            'frmStatus.ShowDetails True
        End If
    End If
        
    bHeaderShown = False
                   
    ' Walk through the symbol pool
    For lIndex = 0 To g.SymbolPool.NumRecords - 1
    
        If Not bRtExport Then
            ' Update the progress bar if necessary
            iPercent = (lIndex / (g.SymbolPool.NumRecords - 1)) * 100
            If iPercent <> iSave Then
                iSave = iPercent
                frmStatus.UpdateProgress "Exporting", iPercent, True
            End If
        End If
        
        ' If the symbol is in the symbol group, export it
        If g.SymbolPool.ArrayTable(lFieldNum, lIndex) = 1 Then
            ' Get the symbol ID of the symbol to export
            lSymbolID = g.SymbolPool.SymbolID(lIndex)
            strSymbol = g.SymbolPool.Symbol(lIndex)
            strDesc = g.SymbolPool.Desc(lIndex)
            
            If m.strFormat = "ASCII" And NoPrices = True Then
                If m.bShowHeader = True And bHeaderShown = False Then
                    astrOutput.Add m.strHeaderLine
                    bHeaderShown = True
                End If
                astrLine.Clear False
                For lField = 0 To UBound(aFields)
                    Select Case aFields(lField)
                    Case "SYMBOL"
                        astrLine.Add strSymbol
                    Case "DESCRIPTION"
                        astrLine.Add strDesc
                    End Select
                Next
                If astrLine.Size > 0 Then astrOutput.Add astrLine.JoinFields(strDelim)
            Else
                ' Get the bars from the Data Manager and export the data
                If bIntraday And Not bRtExport Then
                    frmStatus.AddDetail "Building bars for " & strSymbol
                    Screen.MousePointer = vbHourglass
                End If
                ' TLB 11/7/2013: for exporting, include snapshot (e.g. current session update) even when streaming is on
                If DM_GetBars(Bars, lSymbolID, m.strPeriod, CLng(dStartDate), CLng(m.dEndDate), , , Not m.bSplitAdjust, 2) = True Then
                    ' TLB 4/20/2012: we're now thinking that SpliceBars should only be called for the John Bollinger
                    ' special case, or if the symbol is already in the realtime stream ...
                    If bRtExport Then
                        g.RealTime.SpliceBars Bars
                    ElseIf g.RealTime.SalmonIsRunning Then
                        If Not g.RealTime.SymbolInfo(lSymbolID, False, True) Is Nothing Then
                            g.RealTime.SpliceBars Bars
                        End If
                    End If
                    
                    If Not DirExist(m.strPath) Then MakeDir m.strPath, False
                    If UCase(m.strFormat) = "METASTOCK" Then strFmt = "MS7" Else strFmt = m.strFormat
                    
                    If UCase(m.strFormat) <> "ASCII" Then
                        ' special case: rename ETF's for Vantage Point
                        If UCase(m.strFormat) = "CSI" Then
                            If Left(strSymbol, 2) = "$$" Then
                                strSymbol = Mid(strSymbol, 3)
                            ElseIf strSymbol = "$DIA" Or strSymbol = "$SPY" Or strSymbol = "$QQQQ" Then
                                strSymbol = Mid(strSymbol, 2)
                            End If
                        End If
                        If UCase(m.strFormat) = "GDB" Then
                            Bars.ToFile strFmt, AddSlash(m.strPath) & strSymbol & ".GDB", strSymbol, strDesc
                        Else
                            Bars.ToFile strFmt, m.strPath, strSymbol, strDesc
                        End If
                    Else
                        If bIntraday And Not bRtExport Then
                            frmStatus.AddDetail "Exporting bars for " & strSymbol
                        End If
                        
                        iPos = InStr(strSymbol, "-")
                        If iPos <> 0 And Left(strSymbol, 1) <> "$" Then
                            If Len(Mid(strSymbol, iPos + 1)) = 6 Then
                                strSymbol = Left(strSymbol, iPos) & Right(strSymbol, 4)
                            End If
                        End If
                        
                        If m.bSingleFile = False Then astrOutput.Clear
                        If m.bShowHeader = True Then
                            If (m.bSingleFile = True And bHeaderShown = False) Or (m.bSingleFile = False) Then
                                astrOutput.Add m.strHeaderLine
                                bHeaderShown = True
                            End If
                        End If
                        For lBar = 0 To Bars.Size - 1
                            astrLine.Clear False
                            For lField = 0 To UBound(aFields)
                                Select Case aFields(lField)
                                    Case "SYMBOL"
                                        astrLine.Add strSymbol
                                    Case "DESCRIPTION"
                                        astrLine.Add strDesc
                                    Case "DATE"
                                        astrLine.Add VBA.Format(Bars(eBARS_DateTime, lBar), m.strDateFormat)
                                    Case "TIME"
                                        astrLine.Add VBA.Format(Bars(eBARS_DateTime, lBar), "HhNn")
                                    Case "OPEN"
                                        astrLine.Add PriceString(Bars(eBARS_Open, lBar), Bars.Prop(eBARS_ConvFactor))
                                    Case "HIGH"
                                        astrLine.Add PriceString(Bars(eBARS_High, lBar), Bars.Prop(eBARS_ConvFactor))
                                    Case "LOW"
                                        astrLine.Add PriceString(Bars(eBARS_Low, lBar), Bars.Prop(eBARS_ConvFactor))
                                    Case "CLOSE"
                                        astrLine.Add PriceString(Bars(eBARS_Close, lBar), Bars.Prop(eBARS_ConvFactor))
                                    Case "CONTRACT VOLUME"
                                        astrLine.Add Trim(Str(Bars(eBARS_ContVol, lBar)))
                                    Case "CONTRACT OPEN INTEREST"
                                        astrLine.Add Trim(Str(Bars(eBARS_ContOI, lBar)))
                                    Case "VOLUME"
                                        astrLine.Add Trim(Str(Bars(eBARS_Vol, lBar)))
                                    Case "OPEN INTEREST"
                                        astrLine.Add Trim(Str(Bars(eBARS_OI, lBar)))
                                End Select
                            Next
                            astrOutput.Add astrLine.JoinFields(strDelim)
                        Next lBar
                        If Not m.bSingleFile Then astrOutput.ToFile AddSlash(m.strPath) & strSymbol & "." & strExtension
                    End If
                End If
                If bIntraday Then
                    Screen.MousePointer = 0
                End If
            End If
            DoEvents
        End If
        
        If g.bUnloading Then Exit For
        If Not bRtExport Then
            If frmStatus.Status = eStatus_Aborting Then Exit For
        End If
        
    Next lIndex

    If UCase(m.strFormat) = "ASCII" Then
        If m.bCustomHeader = False Then m.strHeaderLine = ""
        If m.bSingleFile = True Then astrOutput.ToFile AddSlash(m.strPath) & m.strFileName & "." & strExtension
        If m.strFormat = "ASCII" And NoPrices = True Then astrOutput.ToFile AddSlash(m.strPath) & m.strFileName & "." & strExtension
    End If
    
    If Not bRtExport Then frmStatus.Status = eStatus_Completed
    
ErrExit:
    Exit Sub
    
ErrSection:
    If Not bRtExport Then
        If frmStatus.Status = eStatus_Running Then frmStatus.Status = eStatus_Aborted
    End If
    RaiseError "cExportGroup.Export", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetPrivateData
'' Description: Set the current private data to that of a copy
'' Inputs:      Private data of the copy
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub SetPrivateData(mCopyFrom As mPrivate)
On Error GoTo ErrSection:

    ' get copy of all private data
    m = mCopyFrom
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cExportGroup.SetPrivateData", eGDRaiseError_Raise

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MakeCopy
'' Description: Make a copy of this object
'' Inputs:      None
'' Returns:     Copy of the current object
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MakeCopy() As cExportGroup
On Error GoTo ErrSection:

    Dim aCopy As New cExportGroup
    
    aCopy.SetPrivateData m
    Set MakeCopy = aCopy
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cExportGroup.MakeCopy", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PriceString
'' Description: Format a number into a string to output to the export file
'' Inputs:      Number to format, Conversion factor
'' Returns:     String to output
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function PriceString(ByVal dPrice As Double, ByVal dConvFactor As Double) As String
On Error GoTo ErrSection:

    Dim strDisplay As String            ' Format string for the format function
    Dim strTemp As String               ' Temporary string variable
    Dim bSwitchDec As Boolean           ' Whether or not to switch decimal character
    Dim strCharToSwitch As String       ' Character to replace
    Dim strNewChar As String            ' Character to replace with
    Dim strReturn As String             ' String to return from function
    Dim iPos As Integer                 ' Position of something in a string

    Select Case dConvFactor
        Case -4:
            strDisplay = "#0.0#####"
        Case 5, -3:
            strDisplay = "#0.0####"
        Case 4, -2:
            strDisplay = "#0.0###"
        Case 3, -1:
            strDisplay = "#0.0##"
        Case 2:
            strDisplay = "#0.0#"
        Case 1:
            strDisplay = "#0.0"
        Case 0:
            strDisplay = "#0"
    End Select

    strTemp = VBA.Format("1.5")
    bSwitchDec = False
    Select Case UCase(Left(m.strDecimal, 1))
        Case "D"
            If InStr(strTemp, ",") > 0 Then bSwitchDec = True
            strCharToSwitch = ","
            strNewChar = "."
        Case "C"
            If InStr(strTemp, ".") > 0 Then bSwitchDec = True
            strCharToSwitch = "."
            strNewChar = ","
    End Select
    
    strReturn = VBA.Format(dPrice, strDisplay)
    If bSwitchDec = True Then
        iPos = InStr(strReturn, strCharToSwitch)
        Do While iPos <> 0
            strReturn = Left(strReturn, iPos - 1) & strNewChar & Right(strReturn, Len(strReturn) - iPos)
            iPos = InStr(strReturn, strCharToSwitch)
        Loop
    End If
    
    PriceString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cExportGroup.PriceString", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NoPrices
'' Description: Determine whether there are any price fields to be exported
'' Inputs:      None
'' Returns:     True if no price fields, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NoPrices() As Boolean
On Error GoTo ErrSection:

    Dim astrFields As New cGdArray      ' Array of available fields
    Dim strField As String              ' Name of field
    Dim iChecked As Integer             ' Is the field checked on?
    Dim lIndex As Long                  ' Index into a for loop
    
    astrFields.SplitFields m.strFields, "|"
    
    NoPrices = True
    For lIndex = 0 To astrFields.Size - 1
        strField = Parse(astrFields(lIndex), ";", 2)
        iChecked = Parse(astrFields(lIndex), ";", 1)
        
        If UCase(strField) <> "SYMBOL" And UCase(strField) <> "DESCRIPTION" Then
            If iChecked = vbChecked Then
                NoPrices = False
                Exit For
            End If
        End If
    Next lIndex

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cExportGroup.NoPrices", eGDRaiseError_Raise

End Function
