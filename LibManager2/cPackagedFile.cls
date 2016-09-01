VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPackagedFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cPackagedFile.cls
'' Description: Handle importing and exporting of libraries
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Date         Author      Description
'' 04/17/2009   DAJ         If strategy didn't import, don't import other
''                          strategy stuff, don't stop reverify functions on a
''                          save error, don't build rule with secondary price
''                          unless it is a Stop with Limit style order
'' 08/14/2009   DAJ         Allow CR/LF in Library fields
'' 03/09/2011   DAJ         Changed step variables from Integer to Long
'' 03/15/2012   DAJ         Don't try to import function parm if function doesn't exist
'' 05/08/2012   DAJ         Fix for reverifying rules with macros of the same name
'' 04/03/2013   DAJ         Move Strategy Baskets into the database
'' 05/01/2013   DAJ         Fix for 115% on progress bar when importing
'' 07/12/2013   DAJ         Fix for importing strategy baskets, 106% progress
'' 08/05/2013   DAJ         Fix for duplicate strategy baskets after import
'' 10/20/2014   DAJ         Replaced File System Object references
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Option Compare Text

Private Enum eLibSection
    eLibSection_Library = 0
    eLibSection_DLL
    eLibSection_IncludeFiles
    eLibSection_CustomerID
    eLibSection_ExpirationDate
    eLibSection_CreatorID
End Enum

Private Enum eDLLHeader
    eDLLHeader_HeaderLen = 0
    eDLLHeader_FileName
    eDLLHeader_FileLen
    eDLLHeader_FileDate
    eDLLHeader_CRC
    eDLLHeader_BuiltIn
    eDLLHeader_Register
End Enum

Private Type SectionHeader
    Section As eLibSection
    lSize As Long
    Reserved As String * 8
    'Reserved1 As Long
    'Reserved2 As Long
End Type

Private Type FileHeader
    lHeaderSize As Long
    lFileLen As Long
    lFlags As Long
    dFileDate As Double
End Type

' Make sure that the size of the export header always adds up to 80!
' Make sure to pay attention to byte alignment as well!
Private Type ExportHeader
    DLLSize As Long
    Register As Byte
    BuiltIn As Byte
    DLLNameSize As Byte
    Reserved2 As Byte
    CustIDLen As Long
    ExpDate As Long
    IncludeFileLen As Long
    Reserved1 As String * 60
End Type

Private Type mPrivate
    lLibraryID As Long
    strLibraryName As String
    bBuiltin As Boolean
    bGenesisManaged As Boolean ' TLB 5/16/2012: true if is either BuiltIn or has DLLSecurityCode < 0
    vsStatusBar As Object
    StatusBar As cStatusBar
    ctlStatusMsg As Control
    QryDef As QueryDef
    aDepLibs As cGdArray
    aDepLibsDate As cGdArray
    aDepLibsVers As cGdArray
    strPath As String
    lStep As Long
    a() As Variant
    b As Variant
    aMsg As cGdArray
    lMaxStep As Long
    bOnlyIfExist As Boolean
    bABuiltinWasImported As Boolean
    vRetVal As Variant
    strSecondaryLib As String
    lInnerCircleID As Long

    astrSystemName As cGdArray
    alBogusID As cGdArray

    strDLLName As String
    vLibraryType        As Variant

    astrExportFile     As cGdArray

    'Package library fields
    strPackName As String
    strPackDesc As String
    strPackDLLName As String
    strPackVersion As String
    dPackLastModified As Date
    strPackAuthor As String
    bPackSecurityLevel As Byte
    bPackCannotDelete As Boolean
    strPackPassword As String
    strPackRtfFileName As String
    strPackRtfEMail As String
    strPackRtfPhone As String
    strPackRtfWebSite As String
    strPackRequiredMod As String
    
    strCustomerID As String
    dExpDate As Double

    ExportHeader As ExportHeader
End Type
Private m As mPrivate

'Order types
Private Const C_MARKET = 0
Private Const C_STOP = 1
Private Const C_LIMIT = 2
Private Const C_STOPLIMIT = 3
Private Const C_MARKETCLOSE = 4
Private Const C_STOPCLOSE = 5
Private Const C_LIMITCLOSE = 6
Private Const C_STOPLIMITCLOSE = 7

Private Const kVersion = 4
Private Const kClassName = "Class:cPackageFile "
Private Const eLibError = vbObjectError + 1009
Private Const eDupItem = vbObjectError + 1010
Private Const mError = vbObjectError + 1000

Private Function LibSection(ByVal LibSec As eLibSection) As Long
    LibSection = LibSec
End Function

Private Function DLLHdrFld(ByVal HdrFld As eDLLHeader) As Long
    DLLHdrFld = HdrFld
End Function

Property Get LibraryID() As Long
    LibraryID = m.lLibraryID
End Property
Property Get PackName() As Variant
    PackName = m.strPackName
End Property
Property Get PackVersion() As Variant
    PackVersion = m.strPackVersion
End Property
Property Get PackLastModified() As Variant
    PackLastModified = m.dPackLastModified
End Property
Property Get PackAuthor() As Variant
    PackAuthor = m.strPackAuthor
End Property
Property Get PackRequiredMod() As Variant
    PackRequiredMod = m.strPackRequiredMod
End Property
Property Get Path() As String
    Path = m.strPath
End Property
Property Get OnlyIfExist() As Boolean
    OnlyIfExist = m.bOnlyIfExist
End Property
Property Get DLLName() As String
    DLLName = m.strDLLName
End Property
Property Get LibraryType() As Variant
    LibraryType = m.vLibraryType
End Property
Property Get BuiltIn() As Boolean
    BuiltIn = m.bBuiltin
End Property
Property Get SecondaryLib() As String
    SecondaryLib = m.strSecondaryLib
End Property
Property Get CustomerID() As String
    CustomerID = m.strCustomerID
End Property
Property Get ExpirationDate() As Double
    ExpirationDate = m.dExpDate
End Property

Property Let LibraryID(pData As Long)
    m.lLibraryID = pData
End Property
Property Let LibraryName(pData As String)
    m.strLibraryName = pData
End Property
Property Let StatusBar(pData As Object)
    Set m.vsStatusBar = pData
End Property
Property Let Path(pData As String)
    m.strPath = pData
End Property
Property Let StatusMsg(pData As Object)
    Set m.ctlStatusMsg = pData
End Property
Property Let OnlyIfExist(pData As Boolean)
    m.bOnlyIfExist = pData
End Property
Property Let DLLName(pData As String)
    m.strDLLName = pData
End Property
Property Let LibraryType(pData As Variant)
    m.vLibraryType = pData
End Property
Property Let BuiltIn(pData As Boolean)
    m.bBuiltin = pData
End Property
Property Let SecondaryLib(pData As String)
    m.strSecondaryLib = pData
End Property
Property Let CustomerID(ByVal pData As String)
    m.strCustomerID = pData
End Property
Property Let ExpirationDate(ByVal pData As Double)
    m.dExpDate = pData
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Build
'' Description: Package up the library into the GLB file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Build()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index for a for loop
    Dim lCRC As Long                    ' CRC coming back from the zip
    Dim fh As Long                      ' File handle to the output file
    Dim lZipLen As Long                 ' Length of the zipped string
    Dim lDLLZipLen As Long              ' Length of the zipped DLL
    Dim strDLLPath As String            ' Path of the DLL
    Dim bIncludeDLL As Boolean          ' Dont include the dll
    Dim strIncludeFiles As String       ' Include files
    
    Dim strTemp As String               ' Temporary string to hold header
    Dim SecHdr As SectionHeader         ' Section Header
    Dim strSubHeader As String          ' Sub Header for a section
    
    Dim mb As New cMemBuffer
    Dim mbZip As New cMemBuffer
    Dim mbPassword As New cMemBuffer
    
    Set m.astrExportFile = New cGdArray
    m.astrExportFile.Create eGDARRAY_Strings
    
    'Package all tables
    If g.CalledFrom = SystemNavigator Then
        InitializeProgressBar 15
        
        'Create Dependent libraries list
        BuildDepLibs
        
        'Make sure all objects are verified
        CheckVerify
        
        'Package all tables
        ExportTable "Files"
        ExportTable "Functions"
        ExportTable "FunctionParms"
        ExportTable "FunctionRefs"
        ExportTable "Rules"
        ExportTable "RuleParms"
        ExportTable "FunctionRules"
        ExportTable "Systems"
        ExportTable "SystemRules"
        ExportTable "SystemParms"
        ExportTable "SystemSecurities"
        ExportTable "StrategyBaskets"
        ExportTable "StrategyBasketItems"
        ExportTable "StrategyBasketItemParms"
        UpdateMeter "Export Completed Successfully"
    Else
        InitializeProgressBar 13
        
        'Create Dependent libraries list
        BuildDepLibs
        
        'Make sure all objects are verified
        CheckVerify
    
        ExportTable "Functions"
        ExportTable "FunctionParms"
        ExportTable "Rules"
        ExportTable "RuleInputs"
        ExportTable "EZFields"
        ExportTable "EZPhrases"
        ExportTable "Portfolios"
        ExportTable "PortfolioSystems"
        ExportTable "PortfolioRules"
        ExportTable "PortfolioInputs"
        ExportTable "Models"
        ExportTable "ModelRules"
        ExportTable "ModelInputs"
    End If
    
    ' 1) Build Library Information...
    strTemp = m.astrExportFile.JoinFields(vbCrLf)
    With SecHdr
        .Section = eLibSection_Library
        .lSize = Len(strTemp)
        '.Reserved1 = 0
        '.Reserved2 = 0
        .Reserved = 0
    End With
    mb.PutFromMemory GetAddress(SecHdr), 16
    mb.PutStr strTemp
    
    ' 2) Build the DLL information if necessary...
    If m.strDLLName <> "" Then
        strDLLPath = DLLPath(m.bBuiltin, (m.vLibraryType = 1))
        If FileExist(strDLLPath & m.strDLLName) Then
            gdCalcFileCRC32 strDLLPath & m.strDLLName, lCRC
            strSubHeader = m.strDLLName & "," & Str(FileLen(strDLLPath & m.strDLLName)) & _
                            "," & Str(Val(FileDate(strDLLPath & m.strDLLName))) & "," & _
                            Str(lCRC) & "," & CStr(m.bBuiltin) & "," & _
                            Str(m.vLibraryType = 1)
            
            With SecHdr
                .Section = eLibSection_DLL
                .lSize = FileLength(strDLLPath & m.strDLLName) + Len(strSubHeader) + 5
                '.Reserved1 = 0
                '.Reserved2 = 0
                .Reserved = 0
            End With
            mb.PutFromMemory GetAddress(SecHdr), 16
            
            mb.PutLong CLng(Len(strSubHeader) + 5)
            mb.PutStr "," & strSubHeader
            
            mb.FromFile strDLLPath & m.strDLLName, True
        End If
    End If
        
    ' 3) Build any included files...
    Set mbZip = BuildIncludeFiles
    If mbZip.Length > 4 Then
        With SecHdr
            .Section = eLibSection_IncludeFiles
            .lSize = mbZip.Length
            '.Reserved1 = 0
            '.Reserved2 = 0
            .Reserved = 0
        End With
        
        mb.PutFromMemory GetAddress(SecHdr), 16
        mb.PutFromMemory mbZip.MemPtr, mbZip.Length
    End If
    mbZip.Clear
    
    ' 4) Include specific Customer ID's if applicable...
    If Len(m.strCustomerID) > 0 Then
        With SecHdr
            .Section = eLibSection_CustomerID
            .lSize = Len(m.strCustomerID)
            '.Reserved1 = 0
            '.Reserved2 = 0
            .Reserved = 0
        End With
        
        mb.PutFromMemory GetAddress(SecHdr), 16
        mb.PutStr m.strCustomerID
    End If
    
    ' 5) Include expiration date if applicable...
    If m.dExpDate <> 0 Then
        With SecHdr
            .Section = eLibSection_ExpirationDate
            .lSize = 4
            '.Reserved1 = 0
            '.Reserved2 = 0
            .Reserved = 0
        End With
        
        mb.PutFromMemory GetAddress(SecHdr), 16
        mb.PutLong CLng(m.dExpDate)
    End If
    
    ' 6) Include library creator customer ID...
    If g.lLCD <> 0 Then
        With SecHdr
            .Section = eLibSection_CreatorID
            .lSize = 4
            .Reserved = 0
        End With
        
        mb.PutFromMemory GetAddress(SecHdr), 16
        mb.PutLong g.lLCD
    End If
            
    ' Zip up the "Big String"
    mbZip.Length = mb.Length
    lZipLen = ZipMemory(mb.MemPtr, mb.Length, mbZip.MemPtr, mbZip.Length, lCRC)
    mbZip.Length = lZipLen
        
    ' Encrypt the "Zip String"
    mbPassword.Buffer = "LibMgr2"
    gdEncrypt True, mbZip, mbPassword
    
    ' Create binary file
    If Dir(Left(m.strPath, Len(m.strPath) - 4) & ".GLB") <> "" Then Kill Left(m.strPath, Len(m.strPath) - 4) & ".GLB"
    lIndex = mb.Length
    mb.Clear
    mb.PutLong CLng(kVersion) ' GLB Format Version Number
    mb.PutLong lCRC
    mb.PutLong lIndex
    mb.PutLong mbZip.Length
    mb.PutFromMemory mbZip.MemPtr, mbZip.Length
    mb.ToFile ReplaceFileExt(m.strPath, ".GLB")
    
ErrExit:
    Set m.astrExportFile = Nothing
    Set mb = Nothing
    Set mbZip = Nothing
    Exit Sub

ErrSection:
    RaiseError "cPackagedFile.Build", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ReplaceChars
'' Description: Replace certain characters with other characters
'' Inputs:      String to change, String to Search For, String to Replace With
'' Returns:     Fixed String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ReplaceChars(vStr As Variant, strSearchFor As String, strReplaceWith As String) As String
On Error GoTo ErrSection:

    Dim lStart As Long
    Dim strLeftPart As String
    Dim strRightPart As String
    
    lStart = 1
    Do Until lStart = 0
        lStart = InStr(lStart, vStr, strSearchFor)
        If lStart > 0 Then
        
            strLeftPart = Left(vStr, lStart - 1)
            strRightPart = Right(vStr, Len(vStr) - (lStart + (Len(strSearchFor) - 1)))
            vStr = strLeftPart & strReplaceWith & strRightPart
            lStart = lStart + Len(strReplaceWith)
        End If
    Loop
    ReplaceChars = vStr
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPackagedFile.ReplaceChars", eGDRaiseError_Raise, g.strAppPath
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ReInsertTabKeys
'' Description: Insert 'Tab' Keys back into Text Fields
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ReInsertTabKeys()
On Error GoTo ErrSection:
    
    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To UBound(m.b) - 1
        If Not IsNumeric(m.b(lIndex)) Then
            m.b(lIndex) = ReplaceChars(m.b(lIndex), "`", Chr(9))
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPackagedFile.ReInsertTabKeys", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckVerify
'' Description: Check to see if anything needs reverifying before packaging
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CheckVerify()
On Error GoTo ErrSection:
    
    Dim rs As Recordset                 ' Recordset into the database
    
    ' Check for functions with a the reverify flag on
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctions] " & _
        "WHERE [LibraryID]=" & Str(m.lLibraryID) & ";", dbOpenDynaset)
    ValidateCheckSums rs, "tblFunctions"
    rs.FindFirst "[CheckSum] = 0.5"
    If Not rs.NoMatch Then
        rs.Close
        Err.Raise vbObjectError + 1000, , "One or more functions in this library are invalid"
    End If
    rs.FindFirst "[Reverify] = True"
    If Not rs.NoMatch Then
        rs.Close
        Err.Raise vbObjectError + 1000, , "One or more functions in this library need to be verified."
    End If
    rs.Close
    
    ' Check for rules with a the reverify flag on
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] " & _
        "WHERE [LibraryID]=" & Str(m.lLibraryID) & ";", dbOpenDynaset)
    ValidateCheckSums rs, "tblRules"
    rs.FindFirst "[CheckSum] = 0.5"
    If Not rs.NoMatch Then
        rs.Close
        Err.Raise vbObjectError + 1000, , "One or more rules in this library are invalid"
    End If
    rs.FindFirst "[Reverify] = True"
    If Not rs.NoMatch Then
        rs.Close
        Err.Raise vbObjectError + 1000, , "One or more rules in this library need to be verified"
    End If
    rs.Close
    
    ' Check for systems with a the reverify flag on
    If g.CalledFrom = SystemNavigator Then
        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblSystems] " & _
            "WHERE [LibraryID]=" & Str(m.lLibraryID) & ";", dbOpenDynaset)
        ValidateCheckSums rs, "tblSystems"
        rs.FindFirst "[CheckSum] = 0.5"
        If Not rs.NoMatch Then
            rs.Close
            Err.Raise vbObjectError + 1000, , "One or more strategies in this library are invalid"
        End If
        rs.FindFirst "[Reverify] = True"
        If Not rs.NoMatch Then
            rs.Close
            Err.Raise vbObjectError + 1000, , "One or more strategies in this library need to be saved"
        End If
        rs.Close
    Else
        Set rs = g.dbNav.OpenRecordset("SELECT tblPortfolios.LibraryID, tblPortfolios.PortfolioName, tblPortfolioRules.Reverify " & _
            "FROM tblPortfolios INNER JOIN tblPortfolioRules ON tblPortfolios.PortfolioNumber = tblPortfolioRules.PortfolioNumber " & _
            "Where (((tblPortfolios.LibraryID) = " & Str(m.lLibraryID) & ") And ((tblPortfolioRules.Reverify) = True)) " & _
            "ORDER BY tblPortfolios.LibraryID;", dbOpenSnapshot)
        If Not rs.EOF Then
            Err.Raise vbObjectError + 1000, , "Portfolio: " & UCase(rs!PortfolioName) & " has one or more rules that need reverifying."
        End If
    
        'Check for Model rules needing reverifying
        Set rs = g.dbNav.OpenRecordset("SELECT tblModels.LibraryID, tblModels.ModelName, tblModelRules.Reverify " & _
            "FROM tblModels INNER JOIN tblModelRules ON tblModels.ModelNumber = tblModelRules.ModelNumber " & _
            "Where (((tblModels.LibraryID) = " & Str(m.lLibraryID) & ") And ((tblModelRules.Reverify) = True)) " & _
            "ORDER BY tblModels.LibraryID;", dbOpenSnapshot)
        If Not rs.EOF Then
            Err.Raise vbObjectError + 1000, , "Model: " & UCase(rs!ModelName) & " has one or more rules that need reverifying."
        End If
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cPackagedFile.CheckVerify", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExportTable
'' Description: Write Packaged information for current table to the export file
'' Inputs:      Name of the Table to Export
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ExportTable(ByVal strTableName As String)
On Error GoTo ErrSection:
    
    Dim rs As Recordset                 ' Recordset into the database
    Dim rs2 As Recordset                ' Recordset into the database
    Dim rs3 As Recordset                ' Recordset into the database
    Dim lIndex As Integer               ' Index into a for loop
    Dim strTextLine As String           ' Line to add to the export file
    Dim astrTemp As cGdArray            ' Temporary array of Linked Rules
    Dim lIndex2 As Long                 ' Index into a for loop
    
    m.astrExportFile.Add "[tbl" & strTableName & "]"
    
    Set m.QryDef = g.dbNav.QueryDefs("qryPackage" & strTableName)
    m.QryDef.Parameters(0).Value = m.lLibraryID
    Set rs = m.QryDef.OpenRecordset()
    
    If Not rs.EOF Then
        ReDim m.a(rs.Fields.Count)
        
        'Write access table field names to line 2 in section
        If g.CalledFrom = PortfolioNavigator Then
            For lIndex = 0 To rs.Fields.Count - 1
                m.a(lIndex) = rs.Fields.Item(lIndex).Name
            Next lIndex
            strTextLine = Join(m.a(), vbTab)
            strTextLine = ReplaceChars(strTextLine, Chr(13) & Chr(10), "||")
            m.astrExportFile.Add strTextLine
        End If
        
        Do Until rs.EOF
            For lIndex = 0 To rs.Fields.Count - 1
                If IsNull(rs.Fields.Item(lIndex).Value) Then
                    m.a(lIndex) = ""
                Else
                    If strTableName = "Functions" Then
                        Select Case rs.Fields.Item(lIndex).Name
                            Case "Formatted", "FormattedWithFillWords", "CodedText", "Password"
                                m.a(lIndex) = DecryptField(rs.Fields.Item(lIndex))
                            Case Else
                                If IsNumeric(rs.Fields.Item(lIndex).Value) Or rs.Fields.Item(lIndex).Type = dbDate Then
                                    'm.a(lIndex) = rs.Fields.Item(lIndex).Value
                                    m.a(lIndex) = GetField(rs, lIndex)
                                Else
                                    m.a(lIndex) = ReplaceChars(rs.Fields.Item(lIndex).Value, Chr(9), Chr(1))
                                End If
                        End Select
                    ElseIf (strTableName = "Systems") Or (strTableName = "StrategyBaskets") Then
                        Select Case rs.Fields.Item(lIndex).Name
                            Case "Password"
                                m.a(lIndex) = DecryptField(rs.Fields.Item(lIndex))

                            Case Else
                                If IsNumeric(rs.Fields.Item(lIndex).Value) Or rs.Fields.Item(lIndex).Type = dbDate Then
                                    'm.a(lIndex) = rs.Fields.Item(lIndex).Value
                                    m.a(lIndex) = GetField(rs, lIndex)
                                Else
                                    m.a(lIndex) = ReplaceChars(rs.Fields.Item(lIndex).Value, Chr(9), Chr(1))
                                End If
                        End Select
                    ElseIf strTableName = "Rules" Then
                        Select Case rs.Fields.Item(lIndex).Name
                            Case "PreviewRTF", "CondRTF", "CondCoded", "PriceRTF", "PriceCoded", "Price2RTF", "Price2Coded", "Password"
                                m.a(lIndex) = DecryptField(rs.Fields.Item(lIndex))
                            Case Else
                                If IsNumeric(rs.Fields.Item(lIndex).Value) Or rs.Fields.Item(lIndex).Type = dbDate Then
                                    'm.a(lIndex) = rs.Fields.Item(lIndex).Value
                                    m.a(lIndex) = GetField(rs, lIndex)
                                Else
                                    m.a(lIndex) = ReplaceChars(rs.Fields.Item(lIndex).Value, Chr(9), Chr(1))
                                End If
                        End Select
                    ElseIf strTableName = "SystemRules" And rs.Fields.Item(lIndex).Name = "LinkedRules" Then
                        Set astrTemp = New cGdArray
                        astrTemp.Create eGDARRAY_Strings
                        
                        astrTemp.SplitFields NullChk(rs.Fields.Item(lIndex).Value), ","
                        For lIndex2 = 0 To astrTemp.Size - 1
                            If astrTemp(lIndex2) <> "" Then
                                Set rs2 = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] " & _
                                            "WHERE [RuleID]=" & astrTemp(lIndex2) & ";", dbOpenDynaset)
                                If Not rs2.EOF Then
                                    If rs2!SystemNumber = 0 Then
                                        astrTemp(lIndex2) = rs2!Name & "'0"
                                    Else
                                        Set rs3 = g.dbNav.OpenRecordset("SELECT * FROM [tblSystems] " & _
                                                    "WHERE [SystemNumber]=" & Str(rs2!SystemNumber) & ";", dbOpenDynaset)
                                        If Not rs3.EOF Then
                                            astrTemp(lIndex2) = rs2!Name & "'" & rs3!SystemName
                                        Else
                                            astrTemp(lIndex2) = rs2!Name & "'0"
                                        End If
                                    End If
                                Else
                                    astrTemp(lIndex2) = ""
                                End If
                            End If
                        Next lIndex2
                        m.a(lIndex) = astrTemp.JoinFields(",")
                        If Len(m.a(lIndex)) > 0 Then
                            If Left(m.a(lIndex), 1) <> "," Then m.a(lIndex) = "," & m.a(lIndex)
                            If Right(m.a(lIndex), 1) <> "," Then m.a(lIndex) = m.a(lIndex) & ","
                        End If
                    Else
                        'Replace tab characters used in comments.  Convert it to a
                        'special character before saving to the text file since
                        'the field delimiter is also <tab>.
                        If IsNumeric(rs.Fields.Item(lIndex).Value) Or rs.Fields.Item(lIndex).Type = dbDate Then
                            'm.a(lIndex) = rs.Fields.Item(lIndex).Value
                            m.a(lIndex) = GetField(rs, lIndex)
                        Else
                            m.a(lIndex) = ReplaceChars(rs.Fields.Item(lIndex).Value, Chr(9), Chr(1))
                        End If
                    End If
                End If
            Next lIndex
            
            strTextLine = Join(m.a(), vbTab)
            strTextLine = ReplaceChars(strTextLine, Chr(13) & Chr(10), "||")

            m.astrExportFile.Add strTextLine
            rs.MoveNext
        Loop
    End If
    UpdateMeter "Exporting " & strTableName
        
ErrExit:
    Set astrTemp = Nothing
    Set rs = Nothing
    Exit Sub

ErrSection:
    RaiseError "cPackagedFile.ExportTable", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BuildDepLibs
'' Description: Build a List of Dependant Libraries
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub BuildDepLibs()
On Error GoTo ErrSection:
    
    Dim rs As Recordset                 ' Recordset into the database
    Dim lIndex As Integer               ' Index into a for loop
    Dim lIndex2 As Integer              ' Index into a for loop
    Dim bFound As Boolean               ' Was the element found?
    Dim strOutputLine As String         ' Line to add to the output file
    Dim astrRecord As New cGdArray      ' Array of the elements of a record

    ' Write out version of database
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblDatabase];", dbOpenDynaset)
    ValidateCheckSums rs, "tblDatabase"
    If rs!CheckSum = 0.5 Then
        rs.Close
        Err.Raise vbObjectError + 1000, , "Database is invalid"
    Else
        m.astrExportFile.Add "[DatabaseVersion]"
        m.astrExportFile.Add Trim(Str(rs!Version))
    End If
    rs.Close
    
    ' Write packaging library to text file (1st n rows, 1 per field)
    m.astrExportFile.Add "[PackagedLib]"
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblLibrarys] " & _
            "WHERE [LibraryID]=" & Str(m.lLibraryID) & ";", dbOpenDynaset)
    ValidateCheckSums rs, "tblLibrarys"
    If rs!CheckSum = 0.5 Then
        rs.Close
        Err.Raise vbObjectError + 1000, , "Library cannot be Exported because it is invalid"
    End If
    astrRecord.Create eGDARRAY_Strings, rs.Fields.Count - 1
    For lIndex = 1 To rs.Fields.Count - 1
        Select Case rs.Fields.Item(lIndex).Name
            Case "Password"
                astrRecord(lIndex) = DecryptField(rs.Fields.Item(lIndex))
            Case "RequiredMod"
                astrRecord(lIndex) = FixRequiredMod(GetField(rs, lIndex))
            Case Else
                astrRecord(lIndex) = GetField(rs, lIndex)
                ' just to be backwards compatible, do booleans as "True" or "False"
                'If rs.Fields.Item(lIndex).Type = dbBoolean Then
                '    astrRecord(lIndex) = CStr(NullChk(rs.Fields.Item(lIndex).Value))
                'Else
                '    astrRecord(lIndex) = NullChk(rs.Fields.Item(lIndex).Value)
                'End If
        End Select
    Next lIndex
    strOutputLine = astrRecord.JoinFields(vbTab)
    m.astrExportFile.Add Replace(strOutputLine, vbCrLf, "||")
    rs.Close

    ' Create a list of dependent libraries by Name
    m.astrExportFile.Add "[DepLibs]"
    Set m.aDepLibs = New cGdArray
    Set m.aDepLibsDate = New cGdArray
    Set m.aDepLibsVers = New cGdArray
    m.aDepLibs.Create eGDARRAY_Strings
    m.aDepLibsDate.Create eGDARRAY_Strings
    m.aDepLibsVers.Create eGDARRAY_Strings
    
    ' Loop through "qryDep..." queries to determine dependent libraries
    For lIndex = 0 To g.dbNav.QueryDefs.Count - 1
        If Left(g.dbNav.QueryDefs.Item(lIndex).Name, 6) = "qryDep" Then
            Set m.QryDef = g.dbNav.QueryDefs.Item(lIndex)
            m.QryDef.Parameters(0).Value = m.lLibraryID
            Set rs = m.QryDef.OpenRecordset(, dbOpenSnapshot)
            
            Do Until rs.EOF
                bFound = False
                For lIndex2 = 0 To m.aDepLibs.Size - 1
                    If rs!LibraryName = m.aDepLibs.Item(lIndex2) Then
                        bFound = True
                        Exit For
                    End If
                Next lIndex2
                
                'Write dependent Library name to text file
                If Not bFound Then
                    m.aDepLibs.Add rs!LibraryName
                    'm.aDepLibsDate.Add Format(rs!LastMod, "mm/dd/yy HH:MM AM/PM")
                    m.aDepLibsDate.Add Str(Val(rs!LastMod))
                    m.aDepLibsVers.Add Str(NullChk(rs!Vers, "N/A"))
                End If
                
                rs.MoveNext
            Loop
        End If
    Next lIndex
    
    ' Write 1 line per dependent library to text file
    Set astrRecord = New cGdArray
    astrRecord.Create eGDARRAY_Strings, 3
    For lIndex = 0 To m.aDepLibs.Size - 1
        astrRecord(0) = m.aDepLibs.Item(lIndex)
        astrRecord(1) = m.aDepLibsDate.Item(lIndex)
        astrRecord(2) = m.aDepLibsVers.Item(lIndex)
        strOutputLine = astrRecord.JoinFields(vbTab)
        m.astrExportFile.Add strOutputLine
    Next lIndex
    
ErrExit:
    Set astrRecord = Nothing
    Set rs = Nothing
    Exit Sub

ErrSection:
    RaiseError "cPackagedFile.BuildDepLibs", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Import
'' Description: Import a library from a GLB file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Import()
On Error GoTo ErrSection:
    
    Dim strSource As String             ' Store the source of the error
    Dim strMessage As String            ' Store the message of the error
    Dim lErrNumber As Long              ' Store the error number
    Dim astrPaths As New cGdArray       ' Array of files to import
    Dim lIndex As Long                  ' Index into a for loop
    Dim BasketIdMap As cGdTree          ' Strategy basket ID map
    Dim BasketItemIdMap As cGdTree      ' Strategy basket item ID map
        
    ' Do some initialization of variables
    m.bABuiltinWasImported = False
    astrPaths.Create eGDARRAY_Strings
    astrPaths.SplitFields m.strPath, ","
    Set m.astrSystemName = New cGdArray
    m.astrSystemName.Create eGDARRAY_Strings
    Set m.alBogusID = New cGdArray
    m.alBogusID.Create eGDARRAY_Longs
    
    ' Backup the current database
    SaveMDB
    
    ' Walk through the libraries to import and import them
    For lIndex = 0 To astrPaths.Size - 1
        m.strPath = StripStr(astrPaths(lIndex), Chr(34))
        
        Set m.astrExportFile = Nothing
        If LoadLibraryFile(False, False) = False Then GoTo ErrExit
        If DatabaseVersionOK = False Then GoTo ErrExit
        
        g.bChanged = True
        Set m.aMsg = New cGdArray
        m.aMsg.Create eGDARRAY_Strings
        ValidatePackagedLibrary
        
        If g.CalledFrom = SystemNavigator Then
            InitializeProgressBar 16
        
            UpdateMeter "Importing Library Definition"
            ImportLibrary
            
            UpdateMeter "Importing Files"
            ImportFiles
            
            UpdateMeter "Importing Functions"
            ImportFunctions
            
            UpdateMeter "Importing Function Inputs"
            ImportFunctionParms
            
            UpdateMeter "Importing Function Inputs"
            ImportFunctionRefs
            
            UpdateMeter "Importing Strategies"
            ImportSystems
            'FixSystemNumbers
            
            UpdateMeter "Importing Rules"
            ImportRules
            
            UpdateMeter "Importing Rule Inputs"
            ImportRuleParms
            
            UpdateMeter "Importing Rule's Function ref"
            ImportRuleFunctions
            
            UpdateMeter "Importing Strategy Rules"
            ImportSystemRules
            
            UpdateMeter "Importing Strategy Inputs"
            ImportSystemParms
            
            UpdateMeter "Importing Strategy Securities"
            ImportSystemSecurities
            
            UpdateMeter "Importing Strategy Baskets"
            Set BasketIdMap = ImportStrategyBaskets
            
            UpdateMeter "Importing Strategy Basket Items"
            Set BasketItemIdMap = ImportStrategyBasketItems(BasketIdMap)
            
            UpdateMeter "Importing Strategy Basket Item Parms"
            ImportStrategyBasketItemParms BasketItemIdMap
        Else
            InitializeProgressBar 11
        
            UpdateMeter "Importing Library Definition"
            ImportLibrary
            
            'Import Functions and function inputs
            UpdateMeter "Importing Functions"
            ImportTableParent "tblFunctions", "FunctionName", "FunctionID", "Function"
            
            UpdateMeter "Importing Function Inputs"
            ImportTableChild "tblFunctions", "FunctionName", "FunctionID", "Function", _
                             "tblFunctionParms", "ParmName", "ParmID", "Function Parm"
            
            'Import Rules, Inputs, EZFields, EZPhrases
            UpdateMeter "Importing Rules"
            ImportTableParent "tblRules", "RuleName", "RuleID", "Rule"
            
            UpdateMeter "Importing Rule Inputs"
            ImportTableChild "tblRules", "RuleName", "RuleID", "Rule", _
                             "tblRuleInputs", "InputName", "InputID", "Rule Input"
            
            UpdateMeter "Importing Rule EZ Fields"
            ImportTableChild "tblRules", "RuleName", "RuleID", "Rule", _
                             "tblEZFields", "FieldName", "FID", "EZ Field"
            
            UpdateMeter "Importing Rule EZ Phrases"
            ImportEZPhrases
            
            'Import Portfolios, Systems, Rules, and Inputs
            UpdateMeter "Importing Portfolios"
            ImportTableParent "tblPortfolios", "PortfolioName", "PortfolioNumber", "Portfolio"
            
            UpdateMeter "Importing Portfolio Strategies"
            ImportPortfolioSystems
            
            UpdateMeter "Importing Portfolio Rules"
            ImportTableChild "tblPortfolios", "PortfolioName", "PortfolioNumber", "Portfolio", _
                             "tblPortfolioRules", "RuleName", "RuleID", "Portfolio Rule"
            
            UpdateMeter "Importing Portfolio Rules"
            ImportTableChild "tblPortfolios", "PortfolioName", "PortfolioNumber", "Portfolio", _
                             "tblPortfolioInputs", "InputName", "InputID", "Portfolio Input"
            
            'Import Model, Rules, and Inputs
            UpdateMeter "Importing Models"
            ImportTableParent "tblModels", "ModelName", "ModelNumber", "Model"
            
            UpdateMeter "Importing Model Rules"
            ImportTableChild "tblModels", "ModelName", "ModelNumber", "Model", _
                             "tblModelRules", "RuleName", "RuleID", "Model Rule"
            
            UpdateMeter "Importing Model Rules"
            ImportTableChild "tblModels", "ModelName", "ModelNumber", "Model", _
                             "tblModelInputs", "InputName", "InputID", "Model Input"
        
        End If
    Next lIndex
    
    UpdateMeter "Done Importing"
    
    ' Reverify Functions And Rules
    ReverifyFunctions
    ReverifyRules
    If g.CalledFrom = PortfolioNavigator Then
        ReverifyPortRules
        ReverifyModelRules
    End If
    ReverifySystems
    
    InitializeProgressBar 1
    m.lStep = 0
    UpdateMeter "Import Completed Successfully!"
    
    InfBox "Import Completed Successfully", "i", , "Confirmation"
    
    ' Print Error report
    PrintErrorReport
    
ErrExit:
    Cleanup
    ''CloseDb
    Set m.astrExportFile = Nothing
    Exit Sub

ErrSection:
    Select Case Err.Number
        Case vbObjectError + 1000:
            RaiseError "cPackagedFile.Import", eGDRaiseError_Raise, g.strAppPath
    
        'File not found on Kill
        Case 53: Resume Next
            
        'Duplicate Function, Rule, or System
        'Save the error and then re-raise it (must first restore MDB)
        'which clears the call stack
        ''Case eDupItem
            ''strMessage = Err.Description
            ''RestoreMDB
            ''Err.Raise eDupItem, , strMessage
        
        'Library validation errors occur BEFORE the MDB is saved, no
        'restore of the MDB is required
        Case eLibError
            RaiseError "cPackagedFile.Import", eGDRaiseError_Raise, g.strAppPath
            
        Case Else
            'Save the error and then re-raise it (must first restore MDB)
            'which clears the messages
#If 0 Then
            lErrNumber = Err.Number
            strMessage = Err.Description
            strSource = Err.Source
            RestoreMDB
            If lErrNumber < 0 Or Left(strSource, 5) = "Class" Then
                Err.Raise lErrNumber, strSource, strMessage
            Else
                Err.Raise lErrNumber, kClassName & " Sub:Import", strMessage
            End If
#End If
            InfBox "Problems were encountered while trying to import this library.  In " & _
                    "order to ensure that the database has no problems, the program " & _
                    "will be shut down and the old database restored when you restart the " & _
                    "program."
                    
            FileFromString AddSlash(g.strAppPath) & "RestoreMDB.FLG", "Restore"
    End Select

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveMDB
'' Description: Make a backup copy of the current Libraries.MDB
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveMDB()
On Error GoTo ErrSection:
    
    ' Close the database and delete the "Old.MDB"
    ''CloseDb
    KillFile AddSlash(g.strAppPath) & "Old.MDB"
    mGenesis.CopyFiles AddSlash(g.strAppPath) & g.strMdbName, AddSlash(g.strAppPath) & "Old.MDB"
    
    Set g.CommonBridge = CreateObject(g.strCommonDLL & "cCommonBridge")
    With g.CommonBridge
        .AppPath = g.strAppPath
        .CustomerID = g.lLCD
        .dbNavRef = g.dbNav
    End With
    

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cPackagedFile.SaveMDB", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RefreshDatabaseConnection
'' Description: Re-Open the datbase after having closed it
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RefreshDatabaseConnection()
On Error GoTo ErrSection:

    If g.dbNav Is Nothing Then
        Set g.WrkJet = CreateWorkspace("", "admin", "", dbUseJet)
        Set g.dbNav = g.WrkJet.OpenDatabase(AddSlash(g.strAppPath) & g.strMdbName, _
            False, False, "; pwd=" & g.strPassword)
        Set g.CommonBridge = CreateObject(g.strCommonDLL & "cCommonBridge")
        g.CommonBridge.dbNavRef = g.dbNav
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPackagedFile.RefreshDatabaseConnection", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RestoreMDB
'' Description: Restore the Libraries.MDB from the Old.MDB
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RestoreMDB()
On Error GoTo ErrSection:

    Screen.MousePointer = vbHourglass
    
    ' Close the database and kill the Libraries.MDB
    CloseDb
    KillFile AddSlash(g.strAppPath) & g.strMdbName
    
    ' Copy the Old.MDB to Libraries.MDB
    FileCopy AddSlash(g.strAppPath) & "Old.MDB", AddSlash(g.strAppPath) & g.strMdbName
    
ErrExit:
    Screen.MousePointer = vbDefault
    Exit Sub

ErrSection:
    Screen.MousePointer = vbDefault
    RaiseError "cPackagedFile.RestoreMDB", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ImportTables
'' Description: Import tables from a GLB file
'' Inputs:      List of files to import, Header label to print to
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ImportTables(ByVal hastrFiles As Long, lblHeader As Object)
On Error GoTo ErrSection:
    
    Dim strSource As String             ' Store the source of the error
    Dim strMessage As String            ' Store the message of the error
    Dim lErrNumber As Long              ' Store the error number
    Dim astrLibraries As New cGdArray   ' List of libraries to import
    Dim lIndex As Long                  ' Index into a for loop
    Dim rs As Recordset                 ' Recordset into the database
    Dim BasketIdMap As cGdTree          ' Strategy basket ID map
    Dim BasketItemIdMap As cGdTree      ' Strategy basket item ID map
    
    ' Do some intialization
    Set m.astrSystemName = New cGdArray
    m.astrSystemName.Create eGDARRAY_Strings
    Set m.alBogusID = New cGdArray
    m.alBogusID.Create eGDARRAY_Longs
    astrLibraries.SetArrayHandle hastrFiles, False
    m.bABuiltinWasImported = False
    
    ' Save off the current database
    SaveMDB
    
    ' Walk through the library files and import each one
    For lIndex = 0 To astrLibraries.Size - 1
        Set m.astrExportFile = Nothing
        m.strPath = astrLibraries(lIndex)
        
        If Left(m.strPath, 1) = "-" Then
            ' delete the library (if it exists) instead of importing
            m.strPath = Mid(m.strPath, 2)
            If FileExist(m.strPath) Then
                If LoadLibraryFile(True, True) = False Then GoTo ErrExit
                ' Get the Library information from the GLB file
                GetPackagedLibraryInfo
            Else
                ' if file doesn't exist, just assume the library name is the filebase
                ' (since we're just trying to delete the library anyway)
                m.strPackName = FileBase(m.strPath)
            End If
            ' delete library if exists
            Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblLibrarys] WHERE [LibraryName]='" & m.strPackName & "';") ', dbOpenSnapshot)
            If Not rs.EOF Then
                g.bChanged = True
                rs.Delete
            End If
            Set rs = Nothing
        Else
            ' import the library
            If LoadLibraryFile(False, True) = False Then GoTo ErrExit
            If DatabaseVersionOK = False Then GoTo ErrExit
            
            g.bChanged = True
            Set m.aMsg = New cGdArray
            m.vRetVal = m.aMsg.Create(eGDARRAY_Strings)
            ValidatePackagedLibrary True
            lblHeader = "Updating Library:  " & m.strPackName
            lblHeader.Refresh
            
            InitializeProgressBar 15
        
            UpdateMeter "Importing Library Definition"
            ImportLibrary
            
            If g.CalledFrom = SystemNavigator Then
                UpdateMeter "Importing Files"
                ImportFiles
                
                UpdateMeter "Importing Functions"
                ImportFunctions
                
                UpdateMeter "Importing Function Inputs"
                ImportFunctionParms
                
                UpdateMeter "Importing Function's Functions ref"
                ImportFunctionRefs
                
                UpdateMeter "Importing Strategies"
                ImportSystems
                'FixSystemNumbers
                
                UpdateMeter "Importing Rules"
                ImportRules
                
                UpdateMeter "Importing Rule Inputs"
                ImportRuleParms
                
                UpdateMeter "Importing Rule's Function ref"
                ImportRuleFunctions
                
                UpdateMeter "Importing Strategy Rules"
                ImportSystemRules
                
                UpdateMeter "Importing Strategy Inputs"
                ImportSystemParms
                
                UpdateMeter "Importing Strategy Securities"
                ImportSystemSecurities
            
                UpdateMeter "Importing Strategy Baskets"
                Set BasketIdMap = ImportStrategyBaskets
                
                UpdateMeter "Importing Strategy Basket Items"
                Set BasketItemIdMap = ImportStrategyBasketItems(BasketIdMap)
                
                UpdateMeter "Importing Strategy Basket Item Parms"
                ImportStrategyBasketItemParms BasketItemIdMap
            Else
                InitializeProgressBar 13
                
                'Import Functions and function inputs
                ImportTableParent "tblFunctions", "FunctionName", "FunctionID", "Function"
                UpdateMeter "Importing Functions"
                ImportTableChild "tblFunctions", "FunctionName", "FunctionID", "Function", _
                                 "tblFunctionParms", "ParmName", "ParmID", "Function Parm"
                UpdateMeter "Importing Function Inputs"
                
                'Import Rules and Rule inputs
                ImportTableParent "tblRules", "RuleName", "RuleID", "Rule"
                UpdateMeter "Importing Rules"
                ImportTableChild "tblRules", "RuleName", "RuleID", "Rule", _
                                 "tblRuleInputs", "InputName", "InputID", "Rule Input"
                UpdateMeter "Importing Rule Inputs"
                ImportTableChild "tblRules", "RuleName", "RuleID", "Rule", _
                                 "tblFields", "FieldName", "FID", "EZ Field"
                UpdateMeter "Importing Rule EZ Fields"
                ImportEZPhrases
                UpdateMeter "Importing Rule EZ Phrases"
                
                'Import Portfolios, Systems, Rules, and Inputs
                ImportTableParent "tblPortfolios", "PortfolioName", "PortfolioNumber", "Portfolio"
                UpdateMeter "Importing Portfolios"
                
                ImportPortfolioSystems
                UpdateMeter "Importing Portfolio Strategies"
                
                ImportTableChild "tblPortfolios", "PortfolioName", "PortfolioNumber", "Portfolio", _
                                 "tblPortfolioRules", "RuleName", "RuleID", "Portfolio Rule"
                UpdateMeter "Importing Portfolio Rules"
                ImportTableChild "tblPortfolios", "PortfolioName", "PortfolioNumber", "Portfolio", _
                                 "tblPortfolioInputs", "InputName", "InputID", "Portfolio Input"
                UpdateMeter "Importing Portfolio Rules"
                        
                'Import Models, Systems, Rules, and Inputs
                ImportTableParent "tblModels", "ModelName", "ModelNumber", "Model"
                UpdateMeter "Importing Models"
                ImportTableChild "tblModels", "ModelName", "ModelNumber", "Model", _
                                 "tblModelRules", "RuleName", "RuleID", "Model Rule"
                UpdateMeter "Importing Model Rules"
                ImportTableChild "tblModels", "ModelName", "ModelNumber", "Model", _
                                 "tblModelInputs", "InputName", "InputID", "Model Input"
                UpdateMeter "Importing Model Rules"
            
            End If
            
            InitializeProgressBar 1
            m.lStep = 0
            UpdateMeter "Import Completed Successfully!"
            
            DoEvents
        End If
    Next lIndex
            
    ' If any tables were imported, then reverify the rules and functions
    lblHeader = "Verifying Functions"
    lblHeader.Refresh
    ReverifyFunctions
    
    lblHeader = "Verifying Rules"
    lblHeader.Refresh
    DoEvents
    ReverifyRules
    
    If g.CalledFrom = PortfolioNavigator Then
        lblHeader = "Verifying Portfolio Rules"
        lblHeader.Refresh
        DoEvents
        ReverifyPortRules
        
        lblHeader = "Verifying Model Rules"
        lblHeader.Refresh
        DoEvents
        ReverifyModelRules
    End If
    lblHeader = "Verifying Strategies"
    lblHeader.Refresh
    ReverifySystems
    
    ''InfBox "Import Completed Successfully", "i", , "Confirmation"
    
    ' Print Error report
    PrintErrorReport
    
ErrExit:
    Cleanup
    ''CloseDb
    Set m.astrExportFile = Nothing
    Exit Sub

ErrSection:
    Select Case Err.Number
    
        'File not found on Kill
        Case 53: Resume Next
            
        'Duplicate Function, Rule, or System
        'Save the error and then re-raise it (must first restore MDB)
        'which clears the call stack
        ''Case eDupItem
        ''    strMessage = Err.Description
        ''    RestoreMDB
        ''    Err.Raise eDupItem, , strMessage
        
        'Library validation errors occur BEFORE the MDB is saved, no
        'restore of the MDB is required
        Case eLibError
            RaiseError "cPackagedFile.Import", eGDRaiseError_Raise, g.strAppPath
            
        Case Else
            'Save the error and then re-raise it (must first restore MDB)
            'which clears the messages
#If 0 Then
            lErrNumber = Err.Number
            strMessage = Err.Description
            strSource = Err.Source
            RestoreMDB
            If lErrNumber < 0 Or Left(strSource, 5) = "Class" Then
                Err.Raise lErrNumber, strSource, strMessage
            Else
                Err.Raise lErrNumber, kClassName & " Sub:ImportTables ", strMessage
            End If
#End If
            RaiseError "cPackagedFile.ImportTables", eGDRaiseError_Show, g.strAppPath
            InfBox "Problems were encountered while trying to import this library.  In " & _
                    "order to ensure that the database has no problems, the program " & _
                    "will be shut down and the old database restored when you restart the " & _
                    "program."
                    
            FileFromString AddSlash(g.strAppPath) & "RestoreMDB.FLG", "Restore"
    End Select

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PrintErrorReport
'' Description: Print a report of import errors to an output file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PrintErrorReport()
On Error GoTo ErrSection:
    
    Dim lIndex As Integer               ' Index into a for loop
    Dim fh As Integer                   ' File handle for the output file
    
    If m.aMsg Is Nothing Then Exit Sub
    If m.aMsg.Size = 0 Then Exit Sub
    
    fh = FreeFile
    Open AddSlash(g.strAppPath) & "ImportErrors.txt" For Output As #fh
    For lIndex = 0 To m.aMsg.Size - 1
        Print #fh, m.aMsg.Item(lIndex)
    Next lIndex
    Close #fh
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPackagedFile.PrintErrorReport", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ImportLibrary
'' Description: Import the Library table from the import file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ImportLibrary()
On Error GoTo ErrSection:
    
    Dim rs As Recordset                 ' Recordset into the database
    Dim rsMax As Recordset              ' Recordset into the database
    Dim lIndex As Integer               ' Index into a for loop
    
    ' Get information about the Library that has been packaged
    GetPackagedLibraryInfo
    
    If Len(m.strPackName) = 0 Then
        Err.Raise vbObjectError + 1000, , "Export library name was not found in the library file.  Please get a new copy of the library from the developer."
    End If
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblLibrarys] " & _
            "WHERE [LibraryName]='" & m.strPackName & "';", dbOpenDynaset)
    If rs.EOF Then
        Set rsMax = g.dbNav.OpenRecordset("SELECT * FROM [tblLibrarys] " & _
                "ORDER BY [LibraryID] DESC;", dbOpenSnapshot)
        With rs
            .AddNew
            If Not rsMax.EOF Then
                .Fields.Item(0).Value = rsMax!LibraryID + 1
                m.lLibraryID = rsMax!LibraryID + 1
            Else
                .Fields.Item(0).Value = 1
                m.lLibraryID = 1
            End If
            For lIndex = 1 To .Fields.Count - 1
                If lIndex <= UBound(m.b) Then
                    .Fields.Item(lIndex).Value = MoveArrayValue(m.b(lIndex), lIndex, rs)
                Else
                    .Fields(lIndex).Value = MoveArrayValue("", lIndex, rs)
                End If
            Next lIndex
            
            If m.ExportHeader.ExpDate > 0 Then rs!Expiration = m.ExportHeader.ExpDate
            
            ' Need to re-encrypt appropriate fields
            EncryptField !Password, NullChk(!Password)
            
            !CheckSum = BuildCheckSum(rs, "tblLibrarys")
            .Update
        End With
        rs.MoveLast
        m.bBuiltin = rs!BuiltIn
    Else
        m.lLibraryID = rs!LibraryID
        m.bBuiltin = rs!BuiltIn
        With rs
            .Edit
            For lIndex = 1 To .Fields.Count - 1
                If lIndex <= UBound(m.b) Then
                    .Fields.Item(lIndex).Value = MoveArrayValue(m.b(lIndex), lIndex, rs)
                Else
                    .Fields(lIndex).Value = MoveArrayValue("", lIndex, rs)
                End If
            Next lIndex
            If m.ExportHeader.ExpDate > 0 Then rs!Expiration = m.ExportHeader.ExpDate
            
            ' Need to re-encrypt appropriate fields
            EncryptField !Password, NullChk(!Password)
            
            !CheckSum = BuildCheckSum(rs, "tblLibrarys")
            .Update
        End With
    End If
    
    m.bGenesisManaged = False
    If UsingNewVerifyMode Then
        If rs!DLLSecurityCode < 0 Or rs!BuiltIn = True Then
            m.bGenesisManaged = True
        End If
    End If
    
    ' If we are importing the Batting 800 library, we need to get the library
    ' id of Inner Circle as well
    If UCase(m.strPackName) = "BATTING 800" Then
        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblLibrarys] " & _
            "WHERE [LibraryName]='Inner Circle';", dbOpenDynaset)
        If Not rs.EOF Then
            m.lInnerCircleID = rs!LibraryID
        End If
    End If
    
    If m.bBuiltin Then m.bABuiltinWasImported = True

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cPackagedFile.ImportLibrary", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MoveArrayValue
'' Description: Return the correct value to place in the recordset
'' Inputs:      Array Value, Recordset Index, Recordset
'' Returns:     Value to place in the field
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function MoveArrayValue(pArrayVal As Variant, pIndex As Variant, pRs As Recordset) As Variant
On Error GoTo ErrSection:
    
    If pArrayVal = "" And pRs.Fields(pIndex).Type <> dbText Then
        If pRs.Fields(pIndex).Type = dbDate Then
            MoveArrayValue = 0
        Else
            MoveArrayValue = pRs.Fields(pIndex).DefaultValue
        End If
    Else
        Select Case pRs.Fields(pIndex).Type
            Case dbText, dbMemo
                MoveArrayValue = Replace(pArrayVal, Chr(1), Chr(9))
                
            Case dbBinary, dbLongBinary
                MoveArrayValue = pArrayVal
            
            Case dbDate
                MoveArrayValue = GetDate(Str(pArrayVal))
            
            Case Else
                MoveArrayValue = Val(pArrayVal)
        
        End Select
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPackagedFile.MoveArrayValue", eGDRaiseError_Raise, g.strAppPath

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ImportFunctions
'' Description: Import the Functions table from the import file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ImportFunctions()
On Error GoTo ErrSection:
    
    Dim strName As String
    Dim astrNames As cGdArray
    Dim bFound As Boolean
    Dim lIndex2 As Long
    Dim rs As Recordset
    Dim rsByName As Recordset
    Dim strOuputLine As String
    Dim bOK As Boolean
    Dim strMatchName As String
    Dim strSave As String
    Dim rsParms As Recordset
    Dim rsRefs As Recordset
    Dim lStart As Long
    Dim lIndex As Long

    ' Find the Functions section of the import file
    lStart = FindSection("tblFunctions")
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctions] " & _
        "ORDER BY [CodedName];", dbOpenDynaset)
    Set rsByName = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctions] " & _
        "ORDER BY [FunctionName];", dbOpenDynaset)
    Set astrNames = New cGdArray
    astrNames.Create eGDARRAY_Strings
    
    lIndex = lStart
    Do Until lIndex > m.astrExportFile.Size - 1
        ' Get detail record to post
        strOuputLine = m.astrExportFile(lIndex)
        strOuputLine = ReplaceChars(strOuputLine, "||", vbCrLf)
        m.b = Split(strOuputLine, vbTab)
                
        ' Insert any tab keys back into text fields
        ReInsertTabKeys
        
        ' Exit if the line retrieve is the next section header
        If Left(m.b(0), 1) = "[" Then Exit Do
        
        ' Save Function Name for comparison in next section.
        astrNames.Add Str(m.b(8))
        
        With rs
            bOK = True
            
            ' Look for a matching record (always leave ourselves on the record to edit)
            rsByName.FindFirst "[FunctionName]='" & m.b(0) & "'"
            strMatchName = m.b(8)
            .FindFirst "[CodedName]='" & strMatchName & "'"
            If .NoMatch Then
                ' See if English Name already exists
                If rsByName.NoMatch Then
                    strMatchName = ""
                Else
                    strMatchName = rsByName!CodedName
                    .FindFirst "[CodedName]='" & strMatchName & "'"
                End If
            Else
                If Not rsByName.NoMatch Then
                    ' See if the Coded Name match and the English Name match are different functions
                    If rsByName!FunctionID <> !FunctionID Then
                        ' If not builtin, can't add
                        If Not m.bBuiltin And Not m.bGenesisManaged Then
                            bOK = False
                        Else
                            ' Save the Coded Name of the english match
                            strSave = rsByName!CodedName
                            
                            ' Rename or delete the Coded Name match
                            .FindFirst "[CodedName]='_" & strMatchName & "'"
                            If .NoMatch Then
                                rsByName.FindFirst "[FunctionName]='_" & strMatchName & "'"
                                If rsByName.NoMatch Then
                                    .FindFirst "[CodedName]='" & strMatchName & "'"
                                    .Edit
                                    !FunctionName = "_" & strMatchName
                                    !CodedName = "_" & strMatchName
                                    !CheckSum = BuildCheckSum(rs, "tblFunctions")
                                    .Update
                                Else
                                    .FindFirst "[CodedName]='" & strMatchName & "'"
                                    .Delete
                                End If
                            Else
                                .FindFirst "[CodedName]='" & strMatchName & "'"
                                .Delete
                            End If
                            
                            ' Rename or delete the English Name match
                            strMatchName = strSave
                            .FindFirst "[CodedName]='_" & strMatchName & "'"
                            If .NoMatch Then
                                rsByName.FindFirst "[FunctionName]='_" & strMatchName & "'"
                                If rsByName.NoMatch Then
                                    .FindFirst "[CodedName]='" & strMatchName & "'"
                                    .Edit
                                    !FunctionName = "_" & strMatchName
                                    !CodedName = "_" & strMatchName
                                    !CheckSum = BuildCheckSum(rs, "tblFunctions")
                                    .Update
                                Else
                                    .FindFirst "[CodedName]='" & strMatchName & "'"
                                    .Delete
                                End If
                            Else
                                .FindFirst "[CodedName]='" & strMatchName & "'"
                                .Delete
                            End If
                            
                            ' Both functions will be deleted or renamed so we are
                            ' adding a new record
                            strMatchName = ""
                        End If
                    End If
                End If
            End If
            
            ' If found a "matching" record
            If Len(strMatchName) > 0 And bOK = True Then
                ' See if in a different library
                If !LibraryID <> m.lLibraryID Then
                    ' If not builtin, can't add
                    If Not m.bBuiltin And Not m.bGenesisManaged Then
                        ' If we are importing the Batting 800 module and we come to
                        ' a function in the Inner Circle library with the same name,
                        ' change the Library ID to the Batting 800 library instead
                        If UCase(m.strPackName) = "BATTING 800" Then
                            If !LibraryID = m.lInnerCircleID Then
                                .Edit
                                !LibraryID = m.lLibraryID
                                !CheckSum = BuildCheckSum(rs, "tblFunctions")
                                .Update
                            Else
                                bOK = False
                            End If
                        Else
                            bOK = False
                        End If
                    Else
                        ' else first try to rename existing function
                        .FindFirst "[CodedName]='_" & strMatchName & "'"
                        If .NoMatch Then
                            rsByName.FindFirst "[FunctionName]='_" & strMatchName & "'"
                            If rsByName.NoMatch Then
                                .FindFirst "[CodedName]='" & strMatchName & "'"
                                .Edit
                                !FunctionName = "_" & strMatchName
                                !CodedName = "_" & strMatchName
                                !CheckSum = BuildCheckSum(rs, "tblFunctions")
                                .Update
                                strMatchName = ""
                            End If
                        End If
                        .FindFirst "[CodedName]='" & strMatchName & "'"
                    End If
                End If
            End If

            If bOK = False Then
                InfBox "ERROR: Could not import function '" & m.b(0) & "'", "!", , "Import Error"
                m.aMsg.Add "Couldn't add function '" & UCase(m.b(0)) & "' because function exists in another library"
            Else
                If strMatchName = "" Then
                    .AddNew
                Else
                    ' If either of these are going to be changed, flag a requery on rsByName
                    If !FunctionName <> m.b(0) Or !CodedName <> m.b(8) Then
                        strMatchName = ""
                    End If
                    
                    ' When importing an existing function, flag its parms and delete its refs
                    Set rsParms = g.dbNav.OpenRecordset("Select * from [tblFunctionParms] " & _
                            "WHERE [FunctionID]=" & Str(!FunctionID) & ";", dbOpenDynaset)
                    With rsParms
                        lIndex2 = .RecordCount
                        Do While Not .EOF
                            .Edit
                            !ParmNbr = 255
                            .Update
                            .MoveNext
                        Loop
                    End With
                    Set rsParms = Nothing
                    Set rsRefs = g.dbNav.OpenRecordset("Select * from [tblFunctionRefs] " & _
                            "WHERE [FunctionID]=" & Str(!FunctionID) & ";", dbOpenDynaset)
                    Do While Not rsRefs.EOF
                        rsRefs.Delete
                        rsRefs.MoveNext
                    Loop
                    Set rsRefs = Nothing
                    
                    .Edit
                End If
                
                For lIndex2 = 1 To .Fields.Count - 1
                    '.Fields(lIndex2).Value = m.B(lIndex2 - 1)
                    
                    'If nbr of fields in tables GREATER than the number
                    'in the text file then the text file is old.  Use the
                    'default value from the table
                    If lIndex2 - 1 <= UBound(m.b) Then
                        .Fields(lIndex2).Value = MoveArrayValue(m.b(lIndex2 - 1), lIndex2, rs)
                    Else
                        .Fields(lIndex2).Value = MoveArrayValue("", lIndex2, rs)
                    End If
                Next lIndex2
                
                'Special case:
                'If Abreviation column is "" then change to null because of
                'unique index on Abbrev.  It was giving message "record already
                'exits" when execute .Update
                If !Abbrev = "" Then
                    !Abbrev = Null
                End If
                
                If Not m.bBuiltin And Not m.bGenesisManaged Then
                    !Reverify = True
                Else
                    !Reverify = False
                End If
                
                ' Need to re-encrypt the appropriate fields
                EncryptField !Formatted, NullChk(!Formatted)
                EncryptField !FormattedWithFillWords, NullChk(!FormattedWithFillWords)
                EncryptField !CodedText, NullChk(!CodedText)
                EncryptField !Password, NullChk(!Password)
                
                !LibraryID = m.lLibraryID
                !CheckSum = BuildCheckSum(rs, "tblFunctions")
                .Update
                
                ' The only time that we will not do a requery on rsByName is if neither the
                ' FunctionName nor the CodedName changed
                If strMatchName = "" Then
                    rsByName.Requery
                End If
            End If
        End With
        
        lIndex = lIndex + 1
    Loop
    
    ' Delete any existing functions in library that are NOT in the packaged
    ' library these are outdated functions and should be deleted.
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctions] " & _
        "WHERE [LibraryID]=" & Str(m.lLibraryID) & " " & _
        "ORDER BY [LibraryID],[CodedName];", dbOpenDynaset)
    Do Until rs.EOF
        bFound = False
        For lIndex2 = 0 To astrNames.Size - 1
            'Field 9 is CodedName
            If astrNames.Item(lIndex2) = rs.Fields.Item(9).Value Then
                bFound = True
                Exit For
            End If
        Next lIndex2
        If Not bFound Then
            rs.Delete
        End If
        rs.MoveNext
    Loop
    
ErrExit:
    Set rs = Nothing
    Set astrNames = Nothing
    Exit Sub

ErrSection:
    Set astrNames = Nothing
    
    If Err.Number = 3022 Then
        Err.Raise eDupItem, , "A function name: '" & _
            UCase(m.b(0)) & "' already exists.  In order " & _
            "to install this library, rename your version of the " & _
            "function to something else."
    Else
        RaiseError "cPackagedFile.ImportFunctions", eGDRaiseError_Raise, g.strAppPath
    End If
    
End Sub

'This routine imports parent type records (tblFuntions, tblRules, and tblPortfolios
'from the library text file and adds them to the corresponding access table
'Note: This is currently used only by Portfolio Navigator.
'The primary differences are
' - The field list from the access table is included on line 2
' - Child records are always deleted (ie Function inputs, references)

'The query order is important.  It is as follows:
' - All columns                 ie tblFunctions.*
' - LibraryID Selection         LibraryID = [SelLibID] (not checked)

'Example of parms passed...
' - pTableName:         "tblFunctions"
' - pUniqueFieldName:   "FunctionName"
' - pKeyFieldName:      "FunctionID"
' - pType:              "Function"

'Benefits:
' - Can add/remove/change fields anywhere in table.  Don't have to add to end.
' - Greater condidence that fields will be update correctly since field list
'     is included on line 2 and used to match up data.
Public Sub ImportTableParent(pTableName As String, pUniqueFieldName As String, _
    pKeyFieldName As String, pType As String)
On Error GoTo ErrSection:
    
    Dim X           As Long
    Dim Y           As Long
    Dim rs          As Recordset
    Dim rsAll       As Recordset
    Dim outputline  As String
    Dim lStart      As Long
    Dim lIndex      As Long
    Dim FieldNames  As cGdArray
    Dim IndName     As Integer
    Dim IndLibID    As Integer
    Dim Continue    As Boolean
    Dim Fnd         As Boolean
    Dim Names       As cGdArray
    
'---------------------------------------------------------------------------
'Line 1: Section name (ie [tblFunctions])

    lStart = FindSection(pTableName)
    
'---------------------------------------------------------------------------
'Load the unique names used (ie function names) in the packaged library
'Currently in the tables. The "Names" gdArray has the name (ei function name)
'of all items currently in the packaged library in the tables.

    Set Names = New cGdArray
    Names.Create (eGDARRAY_Strings)
    Set rs = g.dbNav.OpenRecordset("Select * from [" & pTableName & _
        "] Where [LibraryID]=" & Str(m.lLibraryID) & _
        " Order by [" & pUniqueFieldName & "];", dbOpenDynaset)
    
    'Find field position of unique field used
    For X = 0 To rs.Fields.Count - 1
        If rs.Fields.Item(X).Name = pUniqueFieldName Then
            IndName = X '<--Borrow IndName for building Names gdArray...
            Exit For
        End If
    Next X
    Do Until rs.EOF
        Names.Add rs.Fields.Item(IndName).Value
        rs.MoveNext
    Loop
    
    'Load ALL functions in the application by unique name (ie tblFunctions
    'would be sorted by FunctionName).
    Set rsAll = g.dbNav.OpenRecordset("Select * from [" & pTableName & "] " & _
        " Order by [" & pUniqueFieldName & "];", dbOpenDynaset)
    
'---------------------------------------------------------------------------
'Line 2 (lStart): contains the field names list
'Load the field names into a gdArray: FieldNames

    outputline = m.astrExportFile(lStart)
    m.b = Split(outputline, vbTab)
    If Left(m.b(0), 1) = "[" Then GoTo ErrExit:
    Set FieldNames = New cGdArray
    FieldNames.Create (eGDARRAY_Strings)
    For X = 0 To UBound(m.b) - 1
        FieldNames.Add m.b(X)
        Select Case m.b(X)
            Case pUniqueFieldName: IndName = X
            Case "LibraryID": IndLibID = X
        End Select
    Next X
    
'---------------------------------------------------------------------------
'Line 3-n:  Read/process each record...

    lIndex = lStart + 1
    Do Until lIndex > m.astrExportFile.Size - 1
    
        'Get data record...
        outputline = m.astrExportFile(lIndex)
        outputline = ReplaceChars(outputline, "||", Chr(13) & Chr(10))
        m.b = Split(outputline, vbTab)
                
        'Exit if the line retrieve is the next section header
        If UBound(m.b) = -1 Then GoTo ErrExit:
        If Left(m.b(0), 1) = "[" Then Exit Do
        
        'Insert any tab keys back into text fields
        ReInsertTabKeys
        
        'Search for unique Name.  If found and in the same library, then
        'replace it with this version...
        Continue = True
        rsAll.FindFirst "[" & pUniqueFieldName & "]='" & m.b(IndName) & "'"
        If rsAll.NoMatch Then
            rs.AddNew
        Else
            'Error - Function in table found but in different library
            If m.lLibraryID <> rsAll!LibraryID Then
                'MsgBox "ERROR: Could not import " & pType & " '" & m.B(IndName) & "'", vbCritical, "ERROR"
                m.aMsg.Add "Couldn't import " & pType & " '" & UCase(m.b(IndName)) & "' because this " & pType & " already exists in another library"
                Continue = False
                Err.Raise mError, , m.aMsg.Item(m.aMsg.Size - 1)
            Else
                'Record found delete it so it's child records are
                'also delete (through cascade deleting,  ie Functions would
                'have it's Parms and references deleted)
                rsAll.Delete
                rs.AddNew
            End If
        End If
        
        'Write data from text file line to recordset...
        If Continue Then
            For X = 0 To rs.Fields.Count - 1
                With rs.Fields.Item(X)
                    If .Name <> pKeyFieldName Then
                        Fnd = False
                        'Loop through fields and set each one.  Skip keys...
                        For Y = 0 To FieldNames.Size - 1
                            If .Name = FieldNames.Item(Y) Then
                                .Value = m.b(Y)
                                Fnd = True
                                Exit For
                            End If
                        Next Y
                        'If value not found in text file then new field.  Take
                        'default value from table
                        If Not Fnd Then
                            .Value = .DefaultValue
                        End If
                    End If
                End With
            Next X
            
            'Set the LibraryID
            rs.Fields.Item(IndLibID).Value = m.lLibraryID
            
            rs.Update
            rs.Requery              'Refresh recordsets
            rsAll.Requery
        End If
        
        lIndex = lIndex + 1
    Loop
    
    
'---------------------------------------------------------------------------
'Delete any existing data items in library that are NOT in the packaged
'library these are outdated items and should be deleted.

    Set rs = g.dbNav.OpenRecordset("Select * from [" & pTableName & "] " & _
        "Where [LibraryID]=" & Str(m.lLibraryID) & _
        " Order by [LibraryID],[" & pUniqueFieldName & "];", dbOpenDynaset)
    
    'Find position of description field used to find unique values (ie functionname)
    For X = 0 To rs.Fields.Count - 1
        If rs.Fields.Item(X).Name = pUniqueFieldName Then
            IndName = X
            Exit For
        End If
    Next X
    
    'If item in table not found in the package list then it is outdated.
    'delete it
    If Names.Size - 1 > 0 Then
        Names.Sort
        Do Until rs.EOF
            If Not Names.BinarySearch(rs.Fields.Item(IndName).Value) Then
                rs.Delete
            End If
            rs.MoveNext
        Loop
    End If
    
ErrExit:
    Set FieldNames = Nothing
    Set Names = Nothing
    Set rs = Nothing
    Set rsAll = Nothing
    Exit Sub

ErrSection:
    RaiseError "cPackagedFile.ImportTableParent", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'This routine imports child type records (tblFuntionParms, tblPortfolioRules, etc.)
'from the library text file and adds them to the corresponding access table
'Note: This is currently used only by Portfolio Navigator.

'The query order is important.  It is as follows:
' - Parent unique field name    ie FunctionName
' - All child columns           ie tblFunctionParms.*
' - Sequencing                  ParmNbr (not checked to include)
' - LibraryID Selection         LibraryID = [SelLibID] (not checked)

'Example of parms passed...
' - pParentTable:       "tblFunctions"
' - pParentUniqueName:  "FunctionName"
' - pParentKey:         "FunctionID"
' - pChildTable:        "tblFunctionParms"
' - pChildUniqueName:   "ParmName"
' - pChildKey:          "ParmID"
' - pType:              "Function Input"
Public Sub ImportTableChild(pParentTable As String, pParentUniqueName As String, _
    pParentKey As String, pParentType As String, pChildTable As String, _
    pChildUniqueName As String, pChildKey As String, pChildType As String)
On Error GoTo ErrSection:
    
    Dim X                       As Long
    Dim Y                       As Long
    Dim rsParent                As Recordset
    Dim rsChild                 As Recordset
    Dim outputline              As String
    Dim lStart                  As Long
    Dim lIndex                  As Long
    Dim FieldNames              As cGdArray
    Dim IndParentUniqueName     As Integer
    Dim IndParentKey            As Integer
    Dim IndChildUniqueName      As Integer
    Dim IndChildsParentKey      As Integer
    Dim Fnd                     As Boolean
    Dim ParentKeyValue          As Variant
    
'---------------------------------------------------------------------------
'Line 1: Section name (ie [tblFunctionParms])

    lStart = FindSection(pChildTable)

'---------------------------------------------------------------------------
'Line 2: contains the field names list
'Load the field name into a gdArray: FieldNames.  Note the position of the
'unique field names used to access the table (ie FunctionName, PortfolioName, etc.)

    outputline = m.astrExportFile(lStart)
    m.b = Split(outputline, vbTab)
    'Exit if the line retrieve is the next section header
    If UBound(m.b) = -1 Then GoTo ErrExit:
    If Left(m.b(0), 1) = "[" Then GoTo ErrExit:
                
    Set FieldNames = New cGdArray
    FieldNames.Create (eGDARRAY_Strings)
    For X = 0 To UBound(m.b) - 1
        FieldNames.Add m.b(X)
        Select Case m.b(X)
            Case pParentUniqueName: IndParentUniqueName = X
            Case pChildUniqueName: IndChildUniqueName = X
        End Select
    Next X
    
    'Load ALL iems in the application by unique name (ie tblFunctions
    'would be sorted by FunctionName).  This is used to retrieve the
    'proper key value (ie FunctionID) to write to the child table
    '(ie FunctionParms).  Also save the field position of the parent key.
    Set rsParent = g.dbNav.OpenRecordset("Select * from [" & pParentTable & "] " & _
        " Order by [" & pParentUniqueName & "];", dbOpenDynaset)
    For X = 0 To rsParent.Fields.Count
        If rsParent.Fields.Item(X).Name = pParentKey Then
            IndParentKey = X
            Exit For
        End If
    Next X
    
    'Open a recordset to child table (ie tblFunctionParms).  Identify the
    'position of the Parent Key name (ie FunctionID)
    Set rsChild = g.dbNav.OpenRecordset("Select * from [" & pChildTable & "];", _
        dbOpenDynaset)
    For X = 0 To rsChild.Fields.Count
        If rsChild.Fields.Item(X).Name = pParentKey Then
            IndChildsParentKey = X
            Exit For
        End If
    Next X
        
'---------------------------------------------------------------------------
'Line 3-n:  Read/process each record...

    lIndex = lStart + 1
    Do Until lIndex > m.astrExportFile.Size - 1
    
        'Get data record...
        outputline = m.astrExportFile(lIndex)
        outputline = ReplaceChars(outputline, "||", Chr(13) & Chr(10))
        m.b = Split(outputline, vbTab)
        
        'Exit if the line retrieve is the next section header
        If UBound(m.b) = -1 Then Exit Do
        If Left(m.b(0), 1) = "[" Then Exit Do
        
        'Insert any tab keys back into text fields
        ReInsertTabKeys
        
        'Search for parent record in order to get the current key value from the
        'table (ie Search by FunctionName to get the FunctionID)
        rsParent.FindFirst "[" & pParentUniqueName & "]='" & _
            m.b(IndParentUniqueName) & "'"
        If rsParent.NoMatch Then
            m.aMsg.Add "Could not import " & pChildType & " because " & _
                pParentType & ": " & UCase(m.b(IndParentUniqueName)) & _
                " was not found."
            Err.Raise mError, , m.aMsg.Item(m.aMsg.Size - 1)
        End If
        ParentKeyValue = rsParent.Fields.Item(IndParentKey).Value
    
        'Loop through each field in the text file record
        'and move the values to the corresponding fields in the recordset.
        rsChild.AddNew
        For X = 0 To rsChild.Fields.Count - 1
            With rsChild.Fields.Item(X)
                If .Name <> pChildKey Then      'Skip child key
                    Fnd = False
                    For Y = 0 To FieldNames.Size - 1
                        If .Name = FieldNames.Item(Y) Then
                            .Value = m.b(Y)
                            Fnd = True
                            Exit For
                        End If
                    Next Y
                    'If value not found in text file then new field.  Take
                    'default value from table
                    If Not Fnd Then
                        .Value = .DefaultValue
                    End If
                End If
            End With
        Next X
            
        'Set the Parent Key in the child record (ei FunctionID)
        rsChild.Fields.Item(IndChildsParentKey).Value = ParentKeyValue
        rsChild.Update
        rsChild.Requery              'Refresh recrodset
        
        lIndex = lIndex + 1
    Loop
        
ErrExit:
    Set FieldNames = Nothing
    Set rsParent = Nothing
    Set rsChild = Nothing
    Exit Sub

ErrSection:
    RaiseError "cPackagedFile.ImportTableChild", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ImportPortfolioSystems
'' Description: Import the Portfolio Systems table from the import file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ImportPortfolioSystems()
On Error GoTo ErrSection:
    
    Dim rsPorts                 As Recordset
    Dim rsSystems               As Recordset
    Dim rsPortSystems           As Recordset
    Dim outputline              As String
    Dim lStart                  As Long
    Dim lIndex                  As Long
    
'---------------------------------------------------------------------------
'Line 1: Section name

    lStart = FindSection("tblPortfolioSystems")

'---------------------------------------------------------------------------
'Line 2: contains the field names list
'Load the field name into a gdArray: FieldNames.  Note the position of the
'unique field names used to access the table (ie FunctionName, PortfolioName, etc.)

    outputline = m.astrExportFile(lStart)
    m.b = Split(outputline, vbTab)
    'Exit if the line retrieve is the next section header
    If UBound(m.b) = -1 Then GoTo ErrExit:
    If Left(m.b(0), 1) = "[" Then GoTo ErrExit:
    
    'Open tables...
    Set rsPortSystems = g.dbNav.OpenRecordset("Select * from [tblPortfolioSystems];", dbOpenDynaset)
    Set rsPorts = g.dbNav.OpenRecordset("Select * from [tblPortfolios] Order by [PortfolioName];", dbOpenSnapshot)
    Set rsSystems = g.dbNav.OpenRecordset("Select * from [tblSystems] Order by [SystemName];", dbOpenSnapshot)
        
'---------------------------------------------------------------------------
'Line 3-n:  Read/process each record...

    lIndex = lStart + 1
    Do Until lIndex > m.astrExportFile.Size - 1
    
        'Get data record...
        outputline = m.astrExportFile(lIndex)
        outputline = ReplaceChars(outputline, "||", Chr(13) & Chr(10))
        m.b = Split(outputline, vbTab)
        
        'Exit if the line retrieve is the next section header
        If UBound(m.b) = -1 Then Exit Do
        If Left(m.b(0), 1) = "[" Then Exit Do
        
        'Get Portfolio number from tblPortfolios...
        rsPorts.FindFirst "[PortfolioName]='" & m.b(0) & "'"
        If rsPorts.NoMatch Then
            m.aMsg.Add "Could not import strategies associated with portfolio" & UCase(m.b(0)) & " because the portfolio was not found."
            Err.Raise mError, , m.aMsg.Item(m.aMsg.Size - 1)
        End If
        
        'Get system number from tblSystems...
        rsSystems.FindFirst "[SystemName]='" & m.b(1) & "'"
        If rsSystems.NoMatch Then
            'OK if not found.  This means the user hasn't bought or loaded
            'the systems used by the portfolio
        Else
            With rsPortSystems
                .AddNew
                !PortfolioNumber = rsPorts!PortfolioNumber
                !SystemNumber = rsSystems!SystemNumber
                .Update
            End With
        End If
        rsPortSystems.Requery              'Refresh recrodset
        
        lIndex = lIndex + 1
    Loop
        
ErrExit:
    Set rsPorts = Nothing
    Set rsPortSystems = Nothing
    Set rsSystems = Nothing
    Exit Sub

ErrSection:
    RaiseError "cPackagedFile.ImportPortfolioSystems", , g.strAppPath

End Sub

'Special routine for handling table that is a child of a child of a parent.
'This occurs with tblEZPhrases.  tblRules (Parent) has tblEZFields (Child)
'and each field has tblEZPhrases (Child of Field).
Public Sub ImportEZPhrases()
On Error GoTo ErrSection:
    
    Dim rsRules                 As Recordset
    Dim rsEZFields              As Recordset
    Dim rsEZPhrases             As Recordset
    Dim outputline              As String
    Dim lStart                  As Long
    Dim lIndex                  As Long
    
'---------------------------------------------------------------------------
'Line 1: Section name

    lStart = FindSection("tblEZPhrases")

'---------------------------------------------------------------------------
'Line 2: contains the field names list
'Load the field name into a gdArray: FieldNames.  Note the position of the
'unique field names used to access the table (ie RuleName and FieldName, etc.)

    outputline = m.astrExportFile(lStart)
    m.b = Split(outputline, vbTab)
    'Exit if the line retrieve is the next section header
    If UBound(m.b) = -1 Then GoTo ErrExit:
    If Left(m.b(0), 1) = "[" Then GoTo ErrExit:
    
    'Open tables...
    Set rsRules = g.dbNav.OpenRecordset("Select * from [tblRules] Order by [RuleName];", dbOpenSnapshot)
    Set rsEZFields = g.dbNav.OpenRecordset("Select * from [tblEZFields] Order by [FieldName];", dbOpenSnapshot)
    Set rsEZPhrases = g.dbNav.OpenRecordset("Select * from [tblEZPhrases];", dbOpenDynaset)
        
'---------------------------------------------------------------------------
'Line 3-n:  Read/process each record (Phrase)

    lIndex = lStart + 1
    Do Until lIndex > m.astrExportFile.Size - 1
    
        'Get data record...
        outputline = m.astrExportFile(lIndex)
        outputline = ReplaceChars(outputline, "||", Chr(13) & Chr(10))
        m.b = Split(outputline, vbTab)
        
        'Exit if the line retrieve is the next section header
        If UBound(m.b) = -1 Then Exit Do
        If Left(m.b(0), 1) = "[" Then Exit Do
        
        'Insert any tab keys back into text fields
        ReInsertTabKeys
        
        'Get Rule number from tblRules...
        rsRules.FindFirst "[RuleName]='" & m.b(0) & "'"
        If rsRules.NoMatch Then
            m.aMsg.Add "Could not import EZ Phrase for Rule: " & UCase(m.b(0)) & " because the Rule was not found."
            Err.Raise mError, , m.aMsg.Item(m.aMsg.Size - 1)
        End If
        
        'Get Field ID from tblEZFields...
        rsEZFields.FindFirst "[RuleID]=" & Str(rsRules!RuleID) & _
            " and [FieldName]='" & m.b(1) & "'"
        If rsEZFields.NoMatch Then
            m.aMsg.Add "Could not import EZ Phrase from EZ Field name: " & UCase(m.b(1)) & " from Rule: " & UCase(m.b(0)) & " because the Field was not found."
            Err.Raise mError, , m.aMsg.Item(m.aMsg.Size - 1)
        End If
        
        'Post the EZ Phrase
        ' NOTE: Need to make this more generic eventually!!
        With rsEZPhrases
            .AddNew
            !FID = rsEZFields!FID
            !Seq = CLng(m.b(4))
            !PrePhrase = Str(m.b(5))
            !NextPhrase = Str(m.b(6))
            !RuleType = Str(m.b(7))
            !ValueLength = CLng(m.b(8))
            !SpaceBefore = CBool(m.b(9))
            !SpaceAfter = CBool(m.b(10))
            .Update
        End With
        
        rsEZPhrases.Requery              'Refresh...
        
        lIndex = lIndex + 1
    Loop
        
ErrExit:
    Set rsRules = Nothing
    Set rsEZPhrases = Nothing
    Set rsEZFields = Nothing
    Exit Sub

ErrSection:
    RaiseError "cPackagedFile.ImportEZPhrases", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ImportFunctionParms
'' Description: Import the Function Parms table from the import file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ImportFunctionParms()
On Error GoTo ErrSection:
    
    Dim X As Long                       ' Index into a for loop
    Dim rsFunctions As Recordset        ' Recordset into the database
    Dim rsParms As Recordset            ' Recordset into the database
    Dim strOuputLine As String          ' Line from the input file
    Dim lIndex As Long                  ' Index into a for loop
    
    lIndex = FindSection("tblFunctionParms")
    
    Set rsFunctions = g.dbNav.OpenRecordset("Select * from [tblFunctions] " & _
        " Where [LibraryID]=" & Str(m.lLibraryID) & _
        " Order by [LibraryID],[CodedName];", dbOpenDynaset)
        
    Set rsParms = g.dbNav.OpenRecordset("Select * from [tblFunctionParms] " & _
        " Order by [FunctionID],[ParmText];", dbOpenDynaset)
    
    Do Until lIndex > m.astrExportFile.Size - 1
        'Get detail record to post...
        strOuputLine = m.astrExportFile(lIndex)
        strOuputLine = ReplaceChars(strOuputLine, "||", Chr(13) & Chr(10))
        m.b = Split(strOuputLine, vbTab)
        
        'Exit if the line retrieve is the next section header
        If Left(m.b(0), 1) = "[" Then Exit Do
        
        'Find the FunctionID
        rsFunctions.FindFirst "[CodedName]='" & m.b(0) & "'"
        If rsFunctions.NoMatch Then
            m.aMsg.Add "Function Parm '" & m.b(2) & "' could not be added because function '" & m.b(0) & "' does not exist"
        Else
            'Search tblfunctionParms with the Parm name
            With rsParms
                .FindFirst "[FunctionID]=" & Str(rsFunctions!FunctionID) & _
                        " and [ParmText]='" & m.b(2) & "'"
                If .NoMatch Then
                    .AddNew
                Else
                    .Edit
                End If
                !FunctionID = rsFunctions!FunctionID
                For X = 2 To .Fields.Count - 1
                    
                    'If nbr of fields in tables GREATER than the number
                    'in the text file then the text file is old.  Use the
                    'default value from the table
                    If X - 1 <= UBound(m.b) Then
                        .Fields(X).Value = MoveArrayValue(m.b(X - 1), X, rsParms)
                    Else
                        .Fields(X).Value = MoveArrayValue("", X, rsParms)
                    End If
                Next X
                .Update
            End With
        End If
        
        lIndex = lIndex + 1
    Loop
    
    'now delete all obsolete parms (flagged as ParmNbr = 255)
    Set rsParms = g.dbNav.OpenRecordset("Select * from [tblFunctionParms] " & _
            " WHERE [ParmNbr]=255;", dbOpenDynaset)
    Do While Not rsParms.EOF
        rsParms.Delete
        rsParms.MoveNext
    Loop
        
ErrExit:
    Set rsFunctions = Nothing
    Set rsParms = Nothing
    Exit Sub

ErrSection:
    RaiseError "cPackagedFile.ImportFunctionParms", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ImportFunctionRefs
'' Description: Import the Function Refs table from the import file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ImportFunctionRefs()
On Error GoTo ErrSection:
    
    Dim X As Long                       ' Index into a for loop
    Dim rsFunctions As Recordset        ' Recordset from the database
    Dim rsChildFunctions As Recordset   ' Recordset from the database
    Dim rsRefs As Recordset             ' Recordset from the database
    Dim strOuputLine As String          ' Line from the input file
    Dim lIndex As Long                  ' Index into a for loop
    
    If m.bBuiltin Or m.bGenesisManaged Then
        lIndex = FindSection("tblFunctionRefs")
        
        Set rsFunctions = g.dbNav.OpenRecordset("Select * from [tblFunctions] " & _
            " Where [LibraryID]=" & Str(m.lLibraryID) & _
            " Order by [LibraryID],[CodedName];", dbOpenDynaset)
        
        Set rsChildFunctions = g.dbNav.OpenRecordset("Select * from [tblFunctions] " & _
            " Order by [CodedName];", dbOpenDynaset)
            
        Set rsRefs = g.dbNav.OpenRecordset("Select * from [tblFunctionRefs] " & _
            " Order by [FunctionID],[FunctionIDRef];", dbOpenDynaset)
        
        Do Until lIndex > m.astrExportFile.Size - 1
        
            'Get detail record to post...
            strOuputLine = m.astrExportFile(lIndex)
            strOuputLine = ReplaceChars(strOuputLine, "||", Chr(13) & Chr(10))
            m.b = Split(strOuputLine, vbTab)
            
            'Exit if the line retrieve is the next section header
            If Left(m.b(0), 1) = "[" Then Exit Do
            
            'Find the FunctionID
            rsFunctions.FindFirst "[CodedName]='" & m.b(0) & "'"
            
            'Find the Sub FunctionID
            rsChildFunctions.FindFirst "[CodedName]='" & m.b(1) & "'"
            If rsChildFunctions.NoMatch Then
                m.aMsg.Add "Function: " & m.b(1) & " used by " & _
                    "Function: " & UCase(m.b(0)) & " was not found."
                If Not rsFunctions.NoMatch Then
                    With rsFunctions
                        .Edit
                        !Reverify = True
                        !CheckSum = BuildCheckSum(rsFunctions, "tblFunctions")
                        .Update
                    End With
                End If
            Else
                'Function found, build reference
                With rsRefs
                    .FindFirst "[FunctionID]=" & Str(rsFunctions!FunctionID) & _
                            " and [FunctionIDRef]=" & Str(rsChildFunctions!FunctionID)
                    If .NoMatch Then
                        .AddNew
                        !FunctionID = rsFunctions!FunctionID
                        !functionidref = rsChildFunctions!FunctionID
                        .Update
                    End If
                End With
            End If
            
            lIndex = lIndex + 1
        Loop
    End If
    
ErrExit:
    Set rsFunctions = Nothing
    Set rsChildFunctions = Nothing
    Set rsRefs = Nothing
    Exit Sub

ErrSection:
    RaiseError "cPackagedFile.ImportFunctionRefs", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ImportRules
'' Description: Import the Rules table from the import file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ImportRules()
On Error GoTo ErrSection:
    
    Dim Name As String
    Dim Fnd As Boolean
    Dim X As Long
    Dim rs As Recordset
    Dim outputline As String
    Dim lIndex As Long
    Dim lPos As Long
    
    Dim rsAllRules As Recordset         ' Recordset of all of the rules
    Dim rsSame As Recordset             ' Recordset of all rules with certain name
    Dim rsAllSystems As Recordset       ' Recordset of all the systems
    Dim strBuffer As String             ' Buffer from the import file
    Dim astrNames As New cGdArray       ' Array of rule names
    Dim bFound As Boolean               ' Was the element found?
    Dim bOK As Boolean                  ' Is everything OK?
    Dim bLocal As Boolean               ' Is the rule to import local?
    Dim lSysNumField As Long            ' Field number for the System Number
    Dim lSystemNumber As Long           ' System nubmer for the current rule
    
    lIndex = FindSection("tblRules")
    
    Set rsAllRules = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] " & _
        "ORDER BY [Name];", dbOpenDynaset)
    Set rsAllSystems = g.dbNav.OpenRecordset("SELECT * FROM [tblSystems] " & _
        "ORDER BY [SystemName];", dbOpenDynaset)
    
    Set astrNames = New cGdArray
    astrNames.Create eGDARRAY_Strings
    
    lSysNumField = rsAllRules.Fields("SystemNumber").OrdinalPosition - 1
    
    Do Until lIndex > m.astrExportFile.Size - 1
    
        'Get detail record to post...
        strBuffer = m.astrExportFile(lIndex)
        strBuffer = ReplaceChars(strBuffer, "||", vbCrLf)
        m.b = Split(strBuffer, vbTab)
        
        'Exit if the line retrieve is the next section header
        If Left(m.b(0), 1) = "[" Then Exit Do
        
        'Insert any tab keys back into text fields
        ReInsertTabKeys
        
        bFound = True
        If Str(m.b(lSysNumField)) = "0" Or Str(m.b(lSysNumField)) = "" Then
            bLocal = False
            lSystemNumber = 0
        Else
            bLocal = True
            rsAllSystems.FindFirst "[SystemName]='" & Str(m.b(lSysNumField)) & "'"
            If Not rsAllSystems.NoMatch Then
                lSystemNumber = rsAllSystems!SystemNumber
            Else
                lSystemNumber = -99999
            End If
        End If
        
        Set rsSame = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] " & _
                    "WHERE [Name]='" & Str(m.b(0)) & "';", dbOpenDynaset)
        
        With rsSame
            bOK = True
            bFound = False
            
            If Not .EOF Then
                .MoveFirst
                
                ' Loop through the rules with this name and change the Library ID
                ' to the Batting 800 library if the rule is currently Inner Circle
                Do While Not .EOF
                    If !LibraryID <> m.lLibraryID Then
                        If UCase(m.strPackName) = "BATTING 800" Then
                            If !LibraryID = m.lInnerCircleID Then
                                .Edit
                                !LibraryID = m.lLibraryID
                                !CheckSum = BuildCheckSum(rsSame, "tblRules")
                                .Update
                            End If
                        End If
                    End If
                    .MoveNext
                Loop
                
                ' Try to find the match...
                .MoveFirst
                .FindFirst "[SystemNumber]=" & Str(lSystemNumber)
                ' If there was no match on the RuleName/SystemNumber combo...
                If .NoMatch Then
                    ' If the importing rule is local...
                        'If bLocal Then
                        '    ' Look for a Shared rule with the same name...
                        '    .FindFirst "[SystemNumber]=0"
                        '    If Not .NoMatch Then bFound = (!LibraryID = m.lLibraryID)
                        'Else
                        '    ' Look for the first rule in this library with the same name...
                        '    .FindFirst "[LibraryID]=" & m.lLibraryID
                        '    If Not .NoMatch Then bFound = True
                        'End If
                        bFound = False
                    
                ' If there was a match on the RuleName/SystemNumber combo...
                Else
                    ' Found an exact match...
                    If !LibraryID = m.lLibraryID Then
                        bFound = True
                        
                    ' Found another shared rule in another library...
                    ElseIf lSystemNumber = 0 And !SystemNumber = 0 Then
                        ' If the Rule we are trying to import is a BuiltIn rule,
                        ' then try to rename the existing rule...
                        If m.bBuiltin Then
                            rsAllRules.FindFirst "[Name]='_" & m.b(0) & "'"
                            If rsAllRules.NoMatch Then
                                .Edit
                                !Name = "_" & m.b(0)
                                !CheckSum = BuildCheckSum(rsSame, "tblRules")
                                .Update
                            End If
                            bFound = False
                        Else
                            bOK = False
                        End If
                    End If
                End If
            End If
            
            If bOK = False Then
                InfBox "ERROR: Could not import rule '" & m.b(0) & "'", "!", , "Import Error"
                m.aMsg.Add "Couldn't add rule '" & UCase(m.b(0)) & "' because rule exists in another library"
            ElseIf lSystemNumber = -99999 Then
                InfBox "ERROR: Could not import rule '" & m.b(0) & "'", "!", , "Import Error"
                m.aMsg.Add "Couldn't add rule '" & UCase(m.b(0)) & "' because it's strategy does not exist"
            Else
                If Not bFound Then
                    .AddNew
                Else
                    .Edit
                End If
                
                For X = 1 To .Fields.Count - 1
                    If .Fields(X).Name <> "SystemNumber" Then
                        'If nbr of fields in tables GREATER than the number
                        'in the text file then the text file is old.  Use the
                        'default value from the table
                        If X - 1 <= UBound(m.b) Then
                            .Fields(X).Value = MoveArrayValue(m.b(X - 1), X, rsSame)
                        Else
                            .Fields(X).Value = MoveArrayValue("", X, rsSame)
                        End If
                    Else
                        If m.b(X - 1) = "0" Then
                            .Fields(X).Value = 0
                        Else
                            'If m.astrSystemName.BinarySearch(m.b(X - 1), lPos) = False Then
                            '    m.astrSystemName.Add m.b(X - 1), lPos
                            '    If m.astrSystemName.Size = 1 Then
                            '        m.alBogusID.Add -1, lPos
                            '    Else
                            '        m.alBogusID.Add m.alBogusID.MinValue - 1, lPos
                            '    End If
                            'End If
                            '.Fields(X).Value = m.alBogusID(lPos)
                            .Fields(X).Value = lSystemNumber
                        End If
                    End If
                Next X
                
                ' Need to re-encrypt appropriate fields
                EncryptField !PreviewRTF, NullChk(!PreviewRTF)
                EncryptField !CondRTF, NullChk(!CondRTF)
                EncryptField !CondCoded, NullChk(!CondCoded)
                EncryptField !PriceRTF, NullChk(!PriceRTF)
                EncryptField !PriceCoded, NullChk(!PriceCoded)
                EncryptField !Price2RTF, NullChk(!Price2RTF)
                EncryptField !Price2Coded, NullChk(!Price2Coded)
                EncryptField !Password, NullChk(!Password)
                
                !Reverify = True
                !LibraryID = m.lLibraryID
                !CheckSum = BuildCheckSum(rsSame, "tblRules")
                astrNames.Add !Name & vbTab & Str(!SystemNumber)
                .Update
            End If
        End With
        
        lIndex = lIndex + 1
    Loop
    
    astrNames.Sort
    
    ' Delete any existing functions in library that are NOT in the packaged
    ' library unless there is a Shared rule that is still being used by a system
    ' in another library
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] " & _
        "WHERE [LibraryID]=" & Str(m.lLibraryID) & " " & _
        "ORDER BY [LibraryID],[Name];", dbOpenDynaset)
    Do Until rs.EOF
        If Not astrNames.BinarySearch(rs!Name & vbTab & Str(rs!SystemNumber)) Then
            rs.Delete
        End If
        rs.MoveNext
    Loop
    
ErrExit:
    Set rs = Nothing
    Set astrNames = Nothing
    Exit Sub

ErrSection:
    If Err.Number = 3022 Then
        Err.Raise eDupItem, , "A rule name: " & _
            " named " & UCase(m.b(0)) & " already exists.  In order " & _
            "to install this library, rename your version of the " & _
            "rule to something else."
    Else
        RaiseError "cPackagedFile.ImportRules", eGDRaiseError_Raise, g.strAppPath
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ImportRuleParms
'' Description: Import the Rule Parms table from the import file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ImportRuleParms()
On Error GoTo ErrSection:
    
    Dim X As Long
    Dim rsRules As Recordset
    Dim rsParms As Recordset
    Dim rsSystems As Recordset
    Dim strOuputLine As String
    Dim lIndex As Long
    Dim lSystemNumber As Long
    Dim astrNames As New cGdArray
    
    lIndex = FindSection("tblRuleParms")
    astrNames.Create eGDARRAY_Strings
    
    Set rsRules = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] " & _
        "WHERE [LibraryID]=" & Str(m.lLibraryID) & " " & _
        "ORDER BY [Name],[SystemNumber];", dbOpenDynaset)
        
    Set rsParms = g.dbNav.OpenRecordset("SELECT * FROM [tblRuleParms] " & _
        "ORDER BY [RuleID],[ParmName];", dbOpenDynaset)
        
    Set rsSystems = g.dbNav.OpenRecordset("SELECT * FROM [tblSystems] " & _
        "ORDER BY [SystemName];", dbOpenDynaset)
    
    Do Until lIndex > m.astrExportFile.Size - 1
    
        'Get detail record to post...
        strOuputLine = m.astrExportFile(lIndex)
        strOuputLine = ReplaceChars(strOuputLine, "||", Chr(13) & Chr(10))
        m.b = Split(strOuputLine, vbTab)
        
        'Exit if the line retrieve is the next section header
        If Left(m.b(0), 1) = "[" Then Exit Do
        
        ' Find the RuleID using the RuleName/System Name pair...
        If Str(m.b(1)) = "0" Or Str(m.b(1)) = "" Then
            lSystemNumber = 0
        Else
            rsSystems.FindFirst "[SystemName]='" & Str(m.b(1)) & "'"
            If Not rsSystems.NoMatch Then
                lSystemNumber = rsSystems!SystemNumber
            Else
                lSystemNumber = -99999
            End If
        End If
        
        rsRules.FindFirst "[Name]='" & Str(m.b(0)) & "' AND [SystemNumber]=" & Str(lSystemNumber)
        If rsRules.NoMatch Then
            m.aMsg.Add "Couldn't add rule parm '" & UCase(m.b(3)) & "' of rule '" & m.b(0) & "'"
        Else
            With rsParms
                .FindFirst "[RuleID]=" & Str(rsRules!RuleID) & " AND [ParmName]='" & Str(m.b(3)) & "'"
                If .NoMatch Then
                    .AddNew
                Else
                    .Edit
                End If
                !RuleID = rsRules!RuleID
                For X = 2 To .Fields.Count - 1
                    'If nbr of fields in tables GREATER than the number
                    'in the text file then the text file is old.  Use the
                    'default value from the table
                    If X <= UBound(m.b) Then
                        .Fields(X).Value = MoveArrayValue(m.b(X), X, rsParms)
                    Else
                        .Fields(X).Value = MoveArrayValue("", X, rsParms)
                    End If
                Next X
                
                astrNames.Add Str(!RuleID) & vbTab & !ParmName
                .Update
            End With
        End If
        
        lIndex = lIndex + 1
    Loop
    
    astrNames.Sort
    
    Set rsParms = g.dbNav.OpenRecordset("SELECT tblRuleParms.*,tblRules.LibraryID " & _
        "FROM [tblRuleParms] INNER JOIN [tblRules] ON tblRuleParms.RuleID = tblRules.RuleID " & _
        "WHERE [LibraryID]=" & Str(m.lLibraryID) & " " & _
        "ORDER BY [tblRuleParms.RuleID],[ParmName];", dbOpenDynaset)
    Do While Not rsParms.EOF
        If Not astrNames.BinarySearch(Str(rsParms!RuleID) & vbTab & rsParms!ParmName) Then
            rsParms.Delete
        End If
        
        rsParms.MoveNext
    Loop
    
ErrExit:
    Set rsRules = Nothing
    Set rsParms = Nothing
    Set rsSystems = Nothing
    Exit Sub

ErrSection:
    RaiseError "cPackagedFile.ImportRuleParms", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ImportSystems
'' Description: Import the Systems table from the import file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ImportSystems()
On Error GoTo ErrSection:
    
    Dim Name As String
    Dim astrNames As cGdArray
    Dim Fnd As Boolean
    Dim X As Long
    Dim rsAllSystems As Recordset
    Dim outputline As String
    Dim bOK As Boolean
    Dim strMatchName As String
    Dim lIndex As Long
    Dim lBogusID As Long
    Dim lPos As Long
    
    lIndex = FindSection("tblSystems")
    
    Set rsAllSystems = g.dbNav.OpenRecordset("SELECT * FROM [tblSystems] " & _
        "ORDER BY [SystemName];", dbOpenDynaset)
    Set astrNames = New cGdArray
    astrNames.Create eGDARRAY_Strings
    
    Do Until lIndex > m.astrExportFile.Size - 1
        'Get detail record to post...
        outputline = m.astrExportFile(lIndex)
        outputline = ReplaceChars(outputline, "||", Chr(13) & Chr(10))
        m.b = Split(outputline, vbTab)
        
        'Exit if the line retrieve is the next section header
        If Left(m.b(0), 1) = "[" Then Exit Do
        
        ' Build a sorted array of System astrNames for the next section...
        If Not astrNames.BinarySearch(Str(m.b(0)), lPos) Then astrNames.Add Str(m.b(0)), lPos
        
        With rsAllSystems
            bOK = True
            'look for a matching record (always leave ourselves on the record to edit)
            strMatchName = m.b(0)
            .FindFirst "[SystemName]='" & strMatchName & "'"
            
            ' If there is no match, we are adding a new record
            If .NoMatch Then
                strMatchName = ""
            
            ' If found a "matching" record
            Else
                'see if in a different library
                If !LibraryID <> m.lLibraryID Then
                    'if not builtin, can't add
                    If Not m.bBuiltin And Not m.bGenesisManaged Then
                        bOK = False
                    Else
                        'else first try to rename existing system
                        .FindFirst "[SystemName]='_" & strMatchName & "'"
                        If .NoMatch Then
                            .FindFirst "[SystemName]='" & strMatchName & "'"
                            .Edit
                            !SystemName = "_" & strMatchName
                            !CheckSum = BuildCheckSum(rsAllSystems, "tblSystems")
                            .Update
                            strMatchName = ""
                        End If
                        .FindFirst "[SystemName]='" & strMatchName & "'"
                    End If
                End If
            End If
            
            If bOK = False Then
                InfBox "ERROR: Could not import strategy '" & m.b(0) & "'", "!", , "Import Error"
                m.aMsg.Add "Couldn't add strategy '" & UCase(m.b(0)) & "' because strategy exists in another library"
            Else
                'If strMatchName <> "" Then
                    'for systems, we want to delete the existing system and
                    're-add it (to make sure we've wiped out all SystemRule
                    'refs, SystemInputs refs, etc.)
                    '.Delete
                'End If
                If strMatchName = "" Then
                    .AddNew
                Else
                    .Edit
                End If
                
                For X = 1 To .Fields.Count - 1
                    
                    'If nbr of fields in tables GREATER than the number
                    'in the text file then the text file is old.  Use the
                    'default value from the table
                    If X - 1 <= UBound(m.b) Then
                        .Fields(X).Value = MoveArrayValue(m.b(X - 1), X, rsAllSystems)
                    Else
                        .Fields(X).Value = MoveArrayValue("", X, rsAllSystems)
                    End If
                    
                Next X
                
                ' Need to re-encrypt appropriate fields
                EncryptField !Password, NullChk(!Password)
                
                !LibraryID = m.lLibraryID
                !CheckSum = BuildCheckSum(rsAllSystems, "tblSystems")
                .Update
            End If
            
        End With
                    
        lIndex = lIndex + 1
    Loop
    
    ' Delete any existing systems in library that are NOT in the packaged
    ' library these are outdated functions and should be deleted.
    Set rsAllSystems = g.dbNav.OpenRecordset("SELECT * FROM [tblSystems] " & _
        "WHERE [LibraryID]=" & Str(m.lLibraryID) & " " & _
        "ORDER BY [SystemName];", dbOpenDynaset)
    Do Until rsAllSystems.EOF
        If Not astrNames.BinarySearch(rsAllSystems!SystemName) Then
            rsAllSystems.Delete
        End If
        
        rsAllSystems.MoveNext
    Loop
    
ErrExit:
    Set rsAllSystems = Nothing
    Set astrNames = Nothing
    Exit Sub

ErrSection:
    If Err.Number = 3022 Then
        Err.Raise eDupItem, , "A strategy: '" & _
            UCase(m.b(0)) & "' already exists.  In order " & _
            "to install this library, rename your version of " & _
            "the strategy to something else."
    Else
        RaiseError "cPackagedFile.ImportSystems", eGDRaiseError_Raise, g.strAppPath
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ImportSystemRules
'' Description: Import the System Rules table from the import file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ImportSystemRules()
On Error GoTo ErrSection:
    
    Dim X As Long
    Dim rsSystems As Recordset
    Dim rsRules As Recordset
    Dim rsSystemRules As Recordset
    Dim strOuputLine As String
    Dim lIndex As Long
    Dim lSystemNumber As Long
    Dim strSystemName As String
    Dim strRuleName As String
    Dim astrTemp As New cGdArray
    Dim strLink As String
    Dim astrNames As New cGdArray
    
    lIndex = FindSection("tblSystemRules")
    astrNames.Create eGDARRAY_Strings
    
    Set rsSystems = g.dbNav.OpenRecordset("SELECT * FROM [tblSystems] " & _
        "WHERE [LibraryID]=" & Str(m.lLibraryID) & " " & _
        "ORDER BY [SystemName];", dbOpenDynaset)
        
    Set rsRules = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] " & _
        "ORDER BY [Name],[SystemNumber];", dbOpenDynaset)
        
    Set rsSystemRules = g.dbNav.OpenRecordset("SELECT * FROM [tblSystemRules] " & _
        "ORDER BY [SystemNumber],[RuleID];", dbOpenDynaset)
    
    Do Until lIndex > m.astrExportFile.Size - 1
        'Get detail record to post...
        strOuputLine = m.astrExportFile(lIndex)
        strOuputLine = ReplaceChars(strOuputLine, "||", Chr(13) & Chr(10))
        m.b = Split(strOuputLine, vbTab)
        
        'Exit if the line retrieve is the next section header
        If Left(m.b(0), 1) = "[" Then Exit Do
        
        'Find the System Number
        lSystemNumber = 0
        rsSystems.FindFirst "[SystemName]='" & m.b(0) & "'"
        If rsSystems.NoMatch Then
            m.aMsg.Add "Strategy: '" & m.b(0) & "' was not found."
        Else
            lSystemNumber = rsSystems!SystemNumber
        
            'Find the RuleID
            rsRules.FindFirst "[Name]='" & m.b(1) & "' AND [SystemNumber]=" & Str(lSystemNumber)
            If rsRules.NoMatch Then
                rsRules.FindFirst "[Name]='" & m.b(1) & "' AND [SystemNumber]=0"
            End If
            
            If rsRules.NoMatch Then
                m.aMsg.Add "Rule: " & m.b(1) & " of Strategy: " & m.b(0) & " was not found."
                If Not rsSystems.NoMatch Then
                    With rsSystems
                        .Edit
                        !Reverify = True
                        !CheckSum = BuildCheckSum(rsSystems, "tblSystems")
                        .Update
                    End With
                End If
            Else
                'Search for Rule/Update
                With rsSystemRules
                    .FindFirst "[SystemNumber]=" & Str(rsSystems!SystemNumber) & " and [RuleID]=" & Str(rsRules!RuleID)
                    If .NoMatch Then
                        .AddNew
                    Else
                        .Edit
                    End If
                    !SystemNumber = rsSystems!SystemNumber
                    !RuleID = rsRules!RuleID
                    For X = 2 To .Fields.Count - 1
                        
                        'If nbr of fields in tables GREATER than the number
                        'in the text file then the text file is old.  Use the
                        'default value from the table
                        If X - 1 <= UBound(m.b) Then
                            .Fields(X).Value = MoveArrayValue(m.b(X), X, rsSystemRules)
                        Else
                            .Fields(X).Value = MoveArrayValue("", X, rsSystemRules)
                        End If
                        
                    Next X
                    
                    ' Change Linked Entries back to Rule ID's from Rule Names if applicable
                    If NullChk(!LinkedRules) <> "" Then
                        Set astrTemp = New cGdArray
                        astrTemp.Create eGDARRAY_Strings
                        astrTemp.SplitFields NullChk(!LinkedRules), ","
                        
                        For X = 0 To astrTemp.Size - 1
                            strRuleName = Parse(astrTemp(X), "'", 1)
                            strSystemName = Parse(astrTemp(X), "'", 2)
                            
                            If strSystemName <> "0" Then
                                rsSystems.FindFirst "[SystemName]='" & strSystemName & "'"
                                If Not rsSystems.NoMatch Then
                                    strSystemName = Str(rsSystems!SystemNumber)
                                Else
                                    strSystemName = ""
                                End If
                            End If
                            If strRuleName <> "" And strSystemName <> "" Then
                                rsRules.FindFirst "[Name]='" & strRuleName & "' AND [SystemNumber]=" & strSystemName
                                If Not rsRules.NoMatch Then
                                    astrTemp(X) = Str(rsRules!RuleID)
                                Else
                                    astrTemp(X) = ""
                                End If
                            Else
                                astrTemp(X) = ""
                            End If
                        Next X
                        
                        strLink = astrTemp.JoinFields(",")
                        If Len(Trim(strLink)) <> 0 Then
                            If Left(strLink, 1) <> "," Then strLink = "," & strLink
                            If Right(strLink, 1) <> "," Then strLink = strLink & ","
                        End If
                        !LinkedRules = strLink
                    End If
                    
                    !CheckSum = BuildCheckSum(rsSystemRules, "tblSystemRules")
                    
                    astrNames.Add Str(!SystemNumber) & vbTab & Str(!RuleID)
                    .Update
                End With
            End If
        End If
        
        lIndex = lIndex + 1
    Loop
    
    astrNames.Sort
    
    Set rsSystems = Nothing
    Set rsRules = Nothing
    Set rsSystemRules = g.dbNav.OpenRecordset("SELECT tblSystemRules.*, tblSystems.LibraryID " & _
        "FROM [tblSystemRules] INNER JOIN [tblSystems] ON tblSystemRules.SystemNumber = tblSystems.SystemNumber " & _
        "WHERE [LibraryID]=" & Str(m.lLibraryID) & " " & _
        "ORDER BY [tblSystemRules.SystemNumber],[RuleID];", dbOpenDynaset)
    Do While Not rsSystemRules.EOF
        If Not astrNames.BinarySearch(Str(rsSystemRules!SystemNumber) & vbTab & Str(rsSystemRules!RuleID)) Then
            rsSystemRules.Delete
        End If
        
        rsSystemRules.MoveNext
    Loop
    
ErrExit:
    Set rsSystems = Nothing
    Set rsRules = Nothing
    Set rsSystemRules = Nothing
    Exit Sub

ErrSection:
    RaiseError "cPackagedFile.ImportSystemRules", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ImportSystemParms
'' Description: Import the System Parms table from the import file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ImportSystemParms()
On Error GoTo ErrSection:
    
    Dim X As Long
    Dim rsSystems As Recordset
    Dim rsRules As Recordset
    Dim rsRuleParms As Recordset
    Dim rsSystemParms As Recordset
    Dim strOuputLine As String
    Dim lIndex As Long
    Dim astrNames As New cGdArray
    
    lIndex = FindSection("tblSystemParms")
    astrNames.Create eGDARRAY_Strings
    
    Set rsSystems = g.dbNav.OpenRecordset("SELECT * FROM [tblSystems] " & _
        "WHERE [LibraryID]=" & Str(m.lLibraryID) & " " & _
        "ORDER BY [SystemName];", dbOpenDynaset)
        
    Set rsRules = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] " & _
        "ORDER BY [Name],[SystemNumber];", dbOpenDynaset)
        
    Set rsRuleParms = g.dbNav.OpenRecordset("SELECT * FROM [tblRuleParms] " & _
        "ORDER BY [RuleID],[ParmName];", dbOpenSnapshot)
        
    Set rsSystemParms = g.dbNav.OpenRecordset("SELECT * FROM [tblSystemParms] " & _
        "ORDER BY [SystemNumber],[ParmID];", dbOpenDynaset)
    
    Do Until lIndex > m.astrExportFile.Size - 1
        'Get detail record to post...
        strOuputLine = m.astrExportFile(lIndex)
        strOuputLine = ReplaceChars(strOuputLine, "||", Chr(13) & Chr(10))
        m.b = Split(strOuputLine, vbTab)
        
        'Exit if the line retrieve is the next section header
        If Left(m.b(0), 1) = "[" Then Exit Do
        
        'Find the System Number
        rsSystems.FindFirst "[SystemName]='" & m.b(0) & "'"
        If rsSystems.NoMatch Then
            m.aMsg.Add "Strategy: '" & m.b(0) & "' was not found."
        Else
            'Search for the Rule Name
            rsRules.FindFirst "[Name]='" & m.b(1) & "' AND [SystemNumber]=" & Str(rsSystems!SystemNumber)
            If rsRules.NoMatch Then
                rsRules.FindFirst "[Name]='" & m.b(1) & "' AND [SystemNumber]=0"
            End If
            
            If rsRules.NoMatch Then
                m.aMsg.Add "Rule: " & UCase(m.b(1)) & " of " & _
                    "Strategy: " & UCase(m.b(0)) & " was not found."
                If Not rsSystems.NoMatch Then
                    With rsSystems
                        .Edit
                        !Reverify = True
                        !CheckSum = BuildCheckSum(rsSystems, "tblSystems")
                        .Update
                    End With
                End If
            Else
                'Search for the Input Name for rule
                rsRuleParms.FindFirst "[RuleID]=" & Str(rsRules!RuleID) & _
                    " and [ParmName]='" & m.b(2) & "'"
                If rsRuleParms.NoMatch Then
                    m.aMsg.Add "Input: " & UCase(m.b(2)) & " of " & _
                        "Rule: " & m.b(1) & " of Strategy: " & UCase(m.b(0)) & " was not " & _
                        "found."
                    With rsRules
                        .Edit
                        !Reverify = True
                        !CheckSum = BuildCheckSum(rsRules, "tblRules")
                        .Update
                    End With
                    With rsSystems
                        .Edit
                        !Reverify = True
                        !CheckSum = BuildCheckSum(rsSystems, "tblSystems")
                        .Update
                    End With
                Else
                    'Search Import the system rule
                    With rsSystemParms
                        .FindFirst "[SystemNumber]=" & Str(rsSystems!SystemNumber) & _
                                " and [ParmID]=" & Str(rsRuleParms!ParmID)
                        If .NoMatch Then
                            .AddNew
                        Else
                            .Edit
                        End If
                        !SystemNumber = rsSystems!SystemNumber
                        !ParmID = rsRuleParms!ParmID
                        For X = 2 To .Fields.Count - 1
                            
                            'If nbr of fields in tables GREATER than the number
                            'in the text file then the text file is old.  Use the
                            'default value from the table
                            If X - 1 <= UBound(m.b) Then
                                .Fields(X).Value = MoveArrayValue(m.b(X + 1), X, rsSystemParms)
                            Else
                                .Fields(X).Value = MoveArrayValue("", X, rsSystemParms)
                            End If
                            
                        Next X
                        
                        astrNames.Add Str(!SystemNumber) & vbTab & Str(!ParmID)
                        .Update
                    End With
                End If
            End If
        End If
        
        lIndex = lIndex + 1
    Loop
    
    astrNames.Sort
    
    Set rsSystemParms = g.dbNav.OpenRecordset("SELECT tblSystemParms.*, tblSystems.LibraryID " & _
        "FROM [tblSystemParms] INNER JOIN [tblSystems] ON tblSystemParms.SystemNumber = tblSystems.SystemNumber " & _
        "WHERE [LibraryID]=" & Str(m.lLibraryID) & " " & _
        "ORDER BY [tblSystemParms.SystemNumber],[ParmID];", dbOpenDynaset)
    Do While Not rsSystemParms.EOF
        If Not astrNames.BinarySearch(Str(rsSystemParms!SystemNumber) & vbTab & Str(rsSystemParms!ParmID)) Then
            rsSystemParms.Delete
        End If
        
        rsSystemParms.MoveNext
    Loop

ErrExit:
    Set rsSystems = Nothing
    Set rsRules = Nothing
    Set rsRuleParms = Nothing
    Set rsSystemParms = Nothing
    Exit Sub

ErrSection:
    RaiseError "cPackagedFile.ImportSystemParms", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ImportSystemSecurities
'' Description: Import the System Securities table from the import file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ImportSystemSecurities()
On Error GoTo ErrSection:
    
    Dim X As Long
    Dim rsSystems As Recordset
    Dim rsRules As Recordset
    Dim rsRuleParms As Recordset
    Dim rsSystemSecurities As Recordset
    Dim strOuputLine As String
    Dim lIndex As Long
    Dim astrNames As New cGdArray
    
    lIndex = FindSection("tblSystemSecurities")
    astrNames.Create eGDARRAY_Strings
    
    Set rsSystems = g.dbNav.OpenRecordset("SELECT * FROM [tblSystems] " & _
        "WHERE [LibraryID]=" & Str(m.lLibraryID) & " " & _
        "ORDER BY [SystemName];", dbOpenDynaset)
        
    Set rsRules = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] " & _
        "ORDER BY [Name],[SystemNumber];", dbOpenDynaset)
        
    Set rsRuleParms = g.dbNav.OpenRecordset("SELECT * FROM [tblRuleParms] " & _
        "ORDER BY [RuleID],[ParmName];", dbOpenDynaset)
        
    Set rsSystemSecurities = g.dbNav.OpenRecordset("SELECT * FROM [tblSystemSecurities] " & _
        "ORDER BY [SystemNumber],[ParmID];", dbOpenDynaset)
    
    Do Until lIndex > m.astrExportFile.Size - 1
        'Get detail record to post...
        strOuputLine = m.astrExportFile(lIndex)
        If Len(strOuputLine) = 0 Then Exit Do
        strOuputLine = ReplaceChars(strOuputLine, "||", Chr(13) & Chr(10))
        m.b = Split(strOuputLine, vbTab)
        
        'Exit if the line retrieve is the next section header
        If Left(m.b(0), 1) = "[" Or Len(m.b(0)) = 0 Then Exit Do
        
        'Find the System Number
        rsSystems.FindFirst "[SystemName]='" & m.b(0) & "'"
        If rsSystems.NoMatch Then
            m.aMsg.Add "Strategy: '" & m.b(0) & "' does not exist."
        Else
            'Search for the Rule Name
            rsRules.FindFirst "[Name]='" & m.b(1) & "' AND [SystemNumber]=" & Str(rsSystems!SystemNumber)
            If rsRules.NoMatch Then
                rsRules.FindFirst "[Name]='" & m.b(1) & "' AND [SystemNumber]=0"
            End If
            
            If rsRules.NoMatch Then
                m.aMsg.Add "Rule: " & m.b(1) & " of " & _
                    "Strategy: " & m.b(0) & " was not found."
                If Not rsSystems.NoMatch Then
                    With rsSystems
                        .Edit
                        !Reverify = True
                        !CheckSum = BuildCheckSum(rsSystems, "tblSystems")
                        .Update
                    End With
                End If
            Else
                'Search for the Input Name in the Rule Inputs table...
                rsRuleParms.FindFirst "[RuleID]=" & Str(rsRules!RuleID) & _
                    " and [ParmName]='" & m.b(2) & "'"
                If rsRuleParms.NoMatch Then
                    m.aMsg.Add "Input: " & m.b(2) & " of " & _
                        "Rule: " & m.b(1) & " of Strategy: " & m.b(0) & " was not " & _
                        "found."
                    With rsRules
                        .Edit
                        !Reverify = True
                        !CheckSum = BuildCheckSum(rsRules, "tblRules")
                        .Update
                    End With
                    With rsSystems
                        .Edit
                        !Reverify = True
                        !CheckSum = BuildCheckSum(rsSystems, "tblSystems")
                        .Update
                    End With
                Else
                    'Search/Update the System Security input...
                    With rsSystemSecurities
                        .FindFirst "[SystemNumber]=" & Str(rsSystems!SystemNumber) & _
                                " and [ParmID]=" & Str(rsRuleParms!ParmID)
                        If .NoMatch Then
                            .AddNew
                        Else
                            .Edit
                        End If
                        !SystemNumber = rsSystems!SystemNumber
                        !ParmID = rsRuleParms!ParmID
                        For X = 2 To .Fields.Count - 1
                            
                            'If nbr of fields in tables GREATER than the number
                            'in the text file then the text file is old.  Use the
                            'default value from the table
                            If X + 1 <= UBound(m.b) Then
                                .Fields(X).Value = MoveArrayValue(m.b(X + 1), X, rsSystemSecurities)
                            Else
                                .Fields(X).Value = MoveArrayValue("", X, rsSystemSecurities)
                            End If
                            
                        Next X
                        
                        astrNames.Add Str(!SystemNumber) & vbTab & Str(!ParmID)
                        .Update
                    End With
                End If
            End If
        End If
        
        lIndex = lIndex + 1
    Loop
    
    astrNames.Sort
    
    Set rsSystemSecurities = g.dbNav.OpenRecordset("SELECT tblSystemSecurities.*, tblSystems.LibraryID " & _
        "FROM [tblSystemSecurities] INNER JOIN [tblSystems] ON tblSystemSecurities.SystemNumber = tblSystems.SystemNumber " & _
        "WHERE [LibraryID]=" & Str(m.lLibraryID) & " " & _
        "ORDER BY [tblSystemSecurities.SystemNumber],[ParmID];", dbOpenDynaset)
    Do While Not rsSystemSecurities.EOF
        If Not astrNames.BinarySearch(Str(rsSystemSecurities!SystemNumber) & vbTab & Str(rsSystemSecurities!ParmID)) Then
            rsSystemSecurities.Delete
        End If
    
        rsSystemSecurities.MoveNext
    Loop
        
ErrExit:
    Set rsSystems = Nothing
    Set rsRules = Nothing
    Set rsRuleParms = Nothing
    Set rsSystemSecurities = Nothing
    Exit Sub

ErrSection:
    RaiseError "cPackagedFile.ImportSystemSecurities", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ImportMMSystemInputs
'' Description: Import the MMSystem Inputs table from the import file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ImportMMSystemInputs()
On Error GoTo ErrSection:

    Dim X As Long
    Dim rsSystems As Recordset
    Dim rsRules As Recordset
    Dim rsRuleParms As Recordset
    Dim rsMMSystemInputs As Recordset
    Dim strOuputLine As String
    Dim lIndex As Long
    
    lIndex = FindSection("tblSystemMMInputs")
    
    Set rsSystems = g.dbNav.OpenRecordset("Select * from [tblSystems] " & _
        " Where [LibraryID]=" & Str(m.lLibraryID) & _
        " Order by [LibraryID],[SystemName];", dbOpenSnapshot)
        
    Set rsRules = g.dbNav.OpenRecordset("Select * from [tblRules] " & _
        " Order by [Name];", dbOpenSnapshot)
        
    Set rsRuleParms = g.dbNav.OpenRecordset("Select * from [tblRuleParms] " & _
        " Order by [RuleID],[ParmName];", dbOpenSnapshot)
        
    Set rsMMSystemInputs = g.dbNav.OpenRecordset("Select * from [tblMMSystemInputs] " & _
        " Order by [SystemNumber];", dbOpenDynaset)
    
    Do Until lIndex > m.astrExportFile.Size - 1
        'Get detail record to post...
        strOuputLine = m.astrExportFile(lIndex)
        strOuputLine = ReplaceChars(strOuputLine, "||", Chr(13) & Chr(10))
        m.b = Split(strOuputLine, vbTab)
        
        'Exit if the line retrieve is the next section header
        If Left(m.b(0), 1) = "[" Then Exit Do
        
        'Find the System Number
        rsSystems.FindFirst "[SystemName]='" & m.b(0) & "'"
        
        'Find the Rule Name
        rsRules.FindFirst "[Name]='" & m.b(1) & "'"
        If rsRules.NoMatch Then
            Err.Raise vbObjectError + 1000, , "Money Management Rule: " & UCase(m.b(1)) & " was " & _
                "not found.  It is used by strategy: " & UCase(m.b(0)) & _
                ".  Make sure this rule exists and then retry the import."
        End If
        
        'Find the ParmName for rule
        rsRuleParms.FindFirst "[RuleID]=" & Str(rsRules!RuleID) & _
            " and [ParmName]='" & m.b(2) & "'"
        If rsRuleParms.NoMatch Then
            Err.Raise vbObjectError + 1000, , "Input: " & UCase(m.b(2)) & " of " & _
                "Money Management Rule: " & m.b(1) & " of Strategy: " & UCase(m.b(0)) & " was not " & _
                "found.  Make sure this input exists and then retry " & _
                "the import."
        End If
        
        'Search Import the system rule
        With rsMMSystemInputs
            .FindFirst "[SystemNumber]=" & Str(rsSystems!SystemNumber) & _
                    " and [ParmID]=" & Str(rsRuleParms!ParmID)
            If .NoMatch Then
                .AddNew
            Else
                .Edit
            End If
            !SystemNumber = rsSystems!SystemNumber
            !ParmID = rsRuleParms!ParmID
            For X = 2 To .Fields.Count - 1
                
                'If nbr of fields in tables GREATER than the number
                'in the text file then the text file is old.  Use the
                'default value from the table
                If X + 1 <= UBound(m.b) Then
                    .Fields(X).Value = MoveArrayValue(m.b(X + 1), X, rsMMSystemInputs)
                Else
                    .Fields(X).Value = MoveArrayValue("", X, rsMMSystemInputs)
                End If
                
            Next X
            .Update
        End With
        
        lIndex = lIndex + 1
    Loop
    
ErrExit:
    Set rsSystems = Nothing
    Set rsRules = Nothing
    Set rsRuleParms = Nothing
    Set rsMMSystemInputs = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "cPackagedFile.ImportMMSystemInputs", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ImportSystemMMInputs
'' Description: Import the System MMInputs table from the import file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ImportSystemMMInputs()
On Error GoTo ErrSection:

    Dim X As Long
    Dim rsSystems As Recordset
    Dim rsRules As Recordset
    Dim rsRuleParms As Recordset
    Dim rsSignalMMInputs As Recordset
    Dim strOuputLine As String
    Dim lIndex As Long
    
    lIndex = FindSection("tblMMSignalInputs")
    
    Set rsSystems = g.dbNav.OpenRecordset("Select * from [tblSystems] " & _
        " Where [LibraryID]=" & Str(m.lLibraryID) & _
        " Order by [LibraryID],[SystemName];", dbOpenSnapshot)
        
    Set rsRules = g.dbNav.OpenRecordset("Select * from [tblRules] " & _
        " Order by [Name];", dbOpenSnapshot)
        
    Set rsRuleParms = g.dbNav.OpenRecordset("Select * from [tblRuleParms] " & _
        " Order by [RuleID],[ParmName];", dbOpenSnapshot)
        
    Set rsSignalMMInputs = g.dbNav.OpenRecordset("Select * from [tblMMSignalInputs] " & _
        " Order by [SystemNumber];", dbOpenDynaset)
    
    Do Until lIndex > m.astrExportFile.Size - 1
        'Get detail record to post...
        strOuputLine = m.astrExportFile(lIndex)
        strOuputLine = ReplaceChars(strOuputLine, "||", Chr(13) & Chr(10))
        m.b = Split(strOuputLine, vbTab)
        
        'Exit if the line retrieve is the next section header
        If Left(m.b(0), 1) = "[" Then Exit Do
        
        'Find the System Number
        rsSystems.FindFirst "[SystemName]='" & m.b(0) & "'"
        
        'Find the Rule Name
        rsRules.FindFirst "[Name]='" & m.b(1) & "'"
        If rsRules.NoMatch Then
            Err.Raise vbObjectError + 1000, , "Money Management Rule: " & UCase(m.b(1)) & " was " & _
                "not found.  It is used by strategy: " & UCase(m.b(0)) & _
                ".  Make sure this rule exists and then retry the import."
        End If
        
        'Find the ParmName for rule
        rsRuleParms.FindFirst "[RuleID]=" & Str(rsRules!RuleID) & _
            " and [ParmName]='" & m.b(2) & "'"
        If rsRuleParms.NoMatch Then
            Err.Raise vbObjectError + 1000, , "Input: " & UCase(m.b(2)) & " of " & _
                "Money Management Rule: " & m.b(1) & " of Strategy: " & UCase(m.b(0)) & " was not " & _
                "found.  Make sure this input exists and then retry " & _
                "the import."
        End If
        
        'Search Import the system rule
        With rsSignalMMInputs
            .FindFirst "[SystemNumber]=" & Str(rsSystems!SystemNumber) & _
                    " and [SignalID]= " & Str(rsRules!RuleID) & _
                    " and [ParmID]=" & Str(rsRuleParms!ParmID)
            If .NoMatch Then
                .AddNew
            Else
                .Edit
            End If
            !SystemNumber = rsSystems!SystemNumber
            !SignalID = rsRules!RuleID
            !ParmID = rsRuleParms!ParmID
            For X = 3 To .Fields.Count - 1
                
                'If nbr of fields in tables GREATER than the number
                'in the text file then the text file is old.  Use the
                'default value from the table
                If X + 1 <= UBound(m.b) Then
                    .Fields(X).Value = MoveArrayValue(m.b(X + 1), X, rsSignalMMInputs)
                Else
                    .Fields(X).Value = MoveArrayValue("", X, rsSignalMMInputs)
                End If
                
            Next X
            .Update
        End With
        
        lIndex = lIndex + 1
    Loop
    
ErrExit:
    Set rsSystems = Nothing
    Set rsRules = Nothing
    Set rsRuleParms = Nothing
    Set rsSignalMMInputs = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "cPackagedFile.ImportSystemMMInputs", , g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ImportMM
'' Description: Import the Money Management table from the import file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ImportMM()
On Error GoTo ErrSection:
    
    Dim Name As String
    Dim Names As cGdArray
    Dim bFound As Boolean
    Dim X As Long
    Dim rs As Recordset
    Dim strOuputLine As String
    Dim lIndex As Long
    
    lIndex = FindSection("tblMM")
    
    Set rs = g.dbNav.OpenRecordset("Select * from [tblMM] " & _
        "Where [LibraryID]=" & Str(m.lLibraryID) & _
        " Order by [LibraryID],[Name];", dbOpenDynaset)
    Set Names = New cGdArray
    m.vRetVal = Names.Create(eGDARRAY_Strings)
    
    Do Until lIndex > m.astrExportFile.Size - 1
        'Get detail record to post...
        strOuputLine = m.astrExportFile(lIndex)
        strOuputLine = ReplaceChars(strOuputLine, "||", Chr(13) & Chr(10))
        m.b = Split(strOuputLine, vbTab)
        
        'Exit if the line retrieve is the next section header
        If Left(m.b(0), 1) = "[" Then Exit Do
        
        'Save function name for comparison in next section.
        Names.Add Str(m.b(0))
        
        With rs
            .FindFirst "[Name]='" & m.b(0) & "'"
            If .NoMatch Then
                .AddNew
            Else
                .Edit
            End If
            For X = 1 To .Fields.Count - 1
                
                'If nbr of fields in tables GREATER than the number
                'in the text file then the text file is old.  Use the
                'default value from the table
                If X - 1 <= UBound(m.b) Then
                    .Fields(X).Value = MoveArrayValue(m.b(X - 1), X, rs)
                Else
                    .Fields(X).Value = MoveArrayValue("", X, rs)
                End If
                
            Next X
            !LibraryID = m.lLibraryID
            .Update
        End With
        
        lIndex = lIndex + 1
    Loop
    
    'Delete any existing functions in library that are NOT in the packaged
    'library these are outdated functions and should be deleted.
    Set rs = g.dbNav.OpenRecordset("Select * from [tblMM] " & _
        "Where [LibraryID]=" & Str(m.lLibraryID) & _
        " Order by [LibraryID],[Name];", dbOpenDynaset)
    Do Until rs.EOF
        bFound = False
        For X = 0 To Names.Size - 1
            If Names.Item(X) = rs.Fields.Item(1).Value Then
                bFound = True
                Exit For
            End If
        Next X
        If Not bFound Then
            rs.Delete
        End If
        rs.MoveNext
    Loop
    
ErrExit:
    Set rs = Nothing
    Set Names = Nothing
    Exit Sub

ErrSection:
    If Err.Number = 3022 Then
        Err.Raise vbObjectError + 1000, , "A Money Management strategy: " & _
            " named " & UCase(m.b(0)) & " already exists.  In order " & _
            "to install this library, rename your version of " & _
            "the strategy to something else."
    Else
        RaiseError "cPackagedFile.ImportMM", eGDRaiseError_Raise, g.strAppPath
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ImportMMRules
'' Description: Import the Money Management Rules table from the import file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ImportMMRules()
On Error GoTo ErrSection:
    
    Dim X As Long
    Dim rsMM As Recordset
    Dim rsRules As Recordset
    Dim rsMMRules As Recordset
    Dim strOuputLine As String
    Dim lIndex As Long
    
    lIndex = FindSection("tblMMRules")
    
    Set rsMM = g.dbNav.OpenRecordset("Select * from [tblMM] " & _
        " Where [LibraryID]=" & Str(m.lLibraryID) & _
        " Order by [LibraryID],[Name];", dbOpenSnapshot)
        
    Set rsRules = g.dbNav.OpenRecordset("Select * from [tblRules] " & _
        " Order by [Name];", dbOpenSnapshot)
        
    Set rsMMRules = g.dbNav.OpenRecordset("Select * from [tblMMRules] " & _
        " Order by [MMID],[RuleID];", dbOpenDynaset)
    
    Do Until lIndex > m.astrExportFile.Size - 1
        'Get detail record to post...
        strOuputLine = m.astrExportFile(lIndex)
        strOuputLine = ReplaceChars(strOuputLine, "||", Chr(13) & Chr(10))
        m.b = Split(strOuputLine, vbTab)
        
        'Exit if the line retrieve is the next section header
        If Left(m.b(0), 1) = "[" Then Exit Do
        
        'Find the MM Number
        rsMM.FindFirst "[Name]='" & m.b(0) & "'"
        
        'Find the RuleID
        rsRules.FindFirst "[Name]='" & m.b(1) & "'"
        If rsRules.NoMatch Then
            Err.Raise vbObjectError + 1000, , "Rule: " & m.b(1) & " was " & _
                "not found.  It is used by Money Management Strategy: " & m.b(0) & _
                ".  Make sure this rule exists and then retry the import."
        End If

        'Search Import the MM rule
        With rsMMRules
            .FindFirst "[MMid]=" & Str(rsMM!MMID) & _
                    " and [RuleID]=" & Str(rsRules!RuleID)
            If .NoMatch Then
                .AddNew
            Else
                .Edit
            End If
            !MMID = rsMM!MMID
            !RuleID = rsRules!RuleID
            For X = 2 To .Fields.Count - 1
                
                'If nbr of fields in tables GREATER than the number
                'in the text file then the text file is old.  Use the
                'default value from the table
                If X <= UBound(m.b) Then
                    .Fields(X).Value = MoveArrayValue(m.b(X), X, rsMMRules)
                Else
                    .Fields(X).Value = MoveArrayValue("", X, rsMMRules)
                End If
                
            Next X
            .Update
        End With
        
        lIndex = lIndex + 1
    Loop

ErrExit:
    Set rsMM = Nothing
    Set rsRules = Nothing
    Set rsMMRules = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "cPackagedFile.ImportMMRules", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ImportMMInputs
'' Description: Import the Money Management Inputs table from the import file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ImportMMInputs()
On Error GoTo ErrSection:
    
    Dim X As Long
    Dim rsMM As Recordset
    Dim rsRules As Recordset
    Dim rsRuleParms As Recordset
    Dim rsMMInputs As Recordset
    Dim strOuputLine As String
    Dim lIndex As Long
    
    lIndex = FindSection("tblMMInputs")
    
    Set rsMM = g.dbNav.OpenRecordset("Select * from [tblMM] " & _
        " Where [LibraryID]=" & Str(m.lLibraryID) & _
        " Order by [LibraryID],[Name];", dbOpenSnapshot)
        
    Set rsRules = g.dbNav.OpenRecordset("Select * from [tblRules] " & _
        " Order by [Name];", dbOpenSnapshot)
        
    Set rsRuleParms = g.dbNav.OpenRecordset("Select * from [tblRuleParms] " & _
        " Order by [RuleID],[ParmName];", dbOpenSnapshot)
        
    Set rsMMInputs = g.dbNav.OpenRecordset("Select * from [tblMMInputs] " & _
        " Order by [MMid];", dbOpenDynaset)
    
    Do Until lIndex > m.astrExportFile.Size - 1
        'Get detail record to post...
        strOuputLine = m.astrExportFile(lIndex)
        strOuputLine = ReplaceChars(strOuputLine, "||", Chr(13) & Chr(10))
        m.b = Split(strOuputLine, vbTab)
        
        'Exit if the line retrieve is the next section header
        If Left(m.b(0), 1) = "[" Then Exit Do
        
        'Find the MM Number
        rsMM.FindFirst "[Name]='" & m.b(0) & "'"
        
        'Find the Rule Name
        rsRules.FindFirst "[Name]='" & m.b(1) & "'"
        If rsRules.NoMatch Then
            Err.Raise vbObjectError + 1000, , "Rule: " & m.b(1) & " was " & _
                "expected to be found.  It is used by MM: " & m.b(0) & _
                ".  Make sure this rule exists and then retry the import."
        End If
        
        'Find the ParmName for rule
        rsRuleParms.FindFirst "[RuleID]=" & Str(rsRules!RuleID) & _
            " and [ParmName]='" & m.b(2) & "'"
        If rsRuleParms.NoMatch Then
            Err.Raise vbObjectError + 1000, , "Input: " & m.b(2) & " of " & _
                "rule: " & m.b(1) & " of MM: " & m.b(0) & " was not " & _
                "found.  Make sure this input exists and then retry " & _
                "the import."
        End If
        
        'Search Import the MM rule
        With rsMMInputs
            .FindFirst "[MMid]=" & Str(rsMM!MMID) & _
                    " and [ParmID]=" & Str(rsRuleParms!ParmID)
            If .NoMatch Then
                .AddNew
            Else
                .Edit
            End If
            !MMID = rsMM!MMID
            !ParmID = rsRuleParms!ParmID
            For X = 2 To .Fields.Count - 1
                
                'If nbr of fields in tables GREATER than the number
                'in the text file then the text file is old.  Use the
                'default value from the table
                If X + 1 <= UBound(m.b) Then
                    .Fields(X).Value = MoveArrayValue(m.b(X + 1), X, rsMMInputs)
                Else
                    .Fields(X).Value = MoveArrayValue("", X, rsMMInputs)
                End If
                
            Next X
            .Update
        End With
        
        lIndex = lIndex + 1
    Loop
    
ErrExit:
    Set rsMM = Nothing
    Set rsRules = Nothing
    Set rsRuleParms = Nothing
    Set rsMMInputs = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "cPackagedFile.ImportMMInputs", eGDRaiseError_Raise, g.strAppPath

End Sub

Private Sub ImportFiles()
On Error GoTo ErrSection:

    Dim X As Long                       ' Index into a for loop
    Dim rsFiles As Recordset            ' Recordset into the database
    Dim strOuputLine As String          ' Line from the input file
    Dim lIndex As Long                  ' Index into a for loop
    
    lIndex = FindSection("tblFiles")
    
    Set rsFiles = g.dbNav.OpenRecordset("SELECT * FROM [tblFiles] " & _
        "WHERE [LibraryID]=" & Str(m.lLibraryID) & " " & _
        "ORDER BY [LibraryID],[FileName];", dbOpenDynaset)
    Do While Not rsFiles.EOF
        rsFiles.Delete
        rsFiles.MoveNext
    Loop
        
    Do Until lIndex > m.astrExportFile.Size - 1
        ' Get detail record to post...
        strOuputLine = m.astrExportFile(lIndex)
        strOuputLine = ReplaceChars(strOuputLine, "||", Chr(13) & Chr(10))
        m.b = Split(strOuputLine, vbTab)
        
        ' Exit if the line retrieve is the next section header...
        If Left(m.b(0), 1) = "[" Then Exit Do
        
        With rsFiles
            .AddNew
            
            For X = 0 To .Fields.Count - 1
                
                'If nbr of fields in tables GREATER than the number
                'in the text file then the text file is old.  Use the
                'default value from the table
                If X <= UBound(m.b) Then
                    .Fields(X).Value = MoveArrayValue(m.b(X), X, rsFiles)
                Else
                    .Fields(X).Value = MoveArrayValue("", X, rsFiles)
                End If
            Next X
            
            !LibraryID = m.lLibraryID
            .Update
        End With
        
        lIndex = lIndex + 1
    Loop
    
ErrExit:
    Set rsFiles = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "cPackagedFile.ImportFiles", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FindSection
'' Description: Find the appropriate section in the input file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FindSection(ByVal strSectionName As String) As Long
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    FindSection = m.astrExportFile.Size + 1
    For lIndex = 0 To m.astrExportFile.Size - 1
        m.b = Split(m.astrExportFile(lIndex), vbTab)
        If m.b(0) = "[" & strSectionName & "]" Then
            FindSection = lIndex + 1
            Exit For
        End If
    Next lIndex

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPackagedFile.FindSection", eGDRaiseError_Raise, g.strAppPath
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetPackagedLibraryInfo
'' Description: Get Library information from the GLB file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetPackagedLibraryInfo()
On Error GoTo ErrSection:
    
    Dim strBuffer As String             ' Buffer from the input array
    Dim lIndex As Long                  ' Index into the array

    ' Try to load the Library File without doing anything with any attached DLLs
    If LoadLibraryFile(True, False) = False Then Exit Sub
    
    ' Retrieve packaged library line "[PackagedLib]"
    lIndex = FindSection("PackagedLib")
    
    ' Get packaging library Information
    strBuffer = m.astrExportFile(lIndex)
    m.b = Split(strBuffer, vbTab)
    m.strPackName = Str(m.b(1))
    m.strPackDesc = Str(m.b(2))
    m.strPackDLLName = Str(m.b(3))
    m.strPackVersion = Str(m.b(4))
    m.dPackLastModified = GetDate(m.b(5))
    m.strPackAuthor = Str(m.b(6))
    m.bPackSecurityLevel = Str(m.b(7))
    m.bPackCannotDelete = Str(m.b(8))
    m.strPackPassword = Str(m.b(9))
    m.strPackRtfFileName = Str(m.b(10))
    m.strPackRtfEMail = Str(m.b(11))
    m.strPackRtfPhone = Str(m.b(12))
    m.strPackRtfWebSite = Str(m.b(13))
    m.strPackRequiredMod = Str(m.b(17))
        
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cPackagedFile.GetPackagedLibaryInfo", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ValidatePackagedLibrary
'' Description: Validate the library against one that they may already have,
''              and validate the libraries that this one depends on
'' Inputs:      Whether or not to show any messages
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ValidatePackagedLibrary(Optional ByVal bSkipMsg As Boolean = False)
On Error GoTo ErrSection:
    
    Dim rs As Recordset
    Dim PackLine As String
    Dim Y As Integer
    Dim Response As Variant
    Dim b As Variant
    Dim lIndex As Long
    Dim strResponse As String
    
    ' Get the Library information from the GLB file
    GetPackagedLibraryInfo
    
    ' Get list of installed librarys on users machine.
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblLibrarys] ORDER BY [LibraryName];", dbOpenSnapshot)
    
    ' See if library being installed already exists on users machine...
    rs.FindFirst "[LibraryName]='" & m.strPackName & "'"
    If Not rs.NoMatch And (bSkipMsg = False) Then
        
        ' Installing library is OLDER than currently installed library
        'If rs!LastModified > m.dPackLastModified Then
        If CLng(rs!Version) > CLng(m.strPackVersion) Then
            strResponse = InfBox("You are attempting to install version " & m.strPackVersion & _
                    " of|" & UCase(m.strPackName) & ".|You currently have version " & Str(rs!Version) & _
                    " on your machine.||This could cause undesired results.|Do you want " & _
                    "to continue?|", , "-Yes|+No", "Confirmation")
            If strResponse = "N" Then Err.Raise eLibError, , "Library was not imported."
        
        ' Installing library is NEWER than currently installed version.
        Else
            strResponse = InfBox("You are about to overwrite version " & Str(rs!Version) & _
                    " of|" & UCase(m.strPackName) & "| with version " & m.strPackVersion & _
                    ".||Do you want to continue?", , "+Yes|-No", "Confirmation")
            If strResponse = "N" Then Err.Raise eLibError, , "Library was not imported."
        End If
    End If
    
    ' Skip Dependent library header
    lIndex = FindSection("DepLibs")
    
    ' Get dependent library Information
    Set m.aDepLibs = New cGdArray
    Set m.aDepLibsDate = New cGdArray
    Set m.aDepLibsVers = New cGdArray
    m.vRetVal = m.aDepLibs.Create(eGDARRAY_Strings)
    m.vRetVal = m.aDepLibsDate.Create(eGDARRAY_Strings)
    m.vRetVal = m.aDepLibsVers.Create(eGDARRAY_Strings)
    
    Do Until lIndex > m.astrExportFile.Size - 1
        PackLine = m.astrExportFile(lIndex)
        b = Split(PackLine, vbTab)
        
        ' Exit when the next section is encountered: "[tblFunctions]"
        If Left(b(0), 1) = "[" Then Exit Do
        m.aDepLibs.Add Str(b(0))
        m.aDepLibsDate.Add Str(b(1))
        m.aDepLibsVers.Add Str(b(2))
        
        lIndex = lIndex + 1
    Loop
    
    ' Verify that all dependent libraries are up to date.
    For Y = 0 To m.aDepLibs.Size - 1
        
        rs.FindFirst "[LibraryName]='" & m.aDepLibs.Item(Y) & "'"
        If rs.NoMatch Then
            
            'Don't allow user to install this library if the
            'dependent library is not found.
            'The reason for this is that rules are stored as copies in
            'Portfolios and Models.  This means a developer of a library
            'could create a model using rules from a different developer
            'that he may have purchased.  If we allow the library to import,
            'the user can use the model or portfolio as is without having
            'to purchase the dependent library.
            If g.CalledFrom = PortfolioNavigator Then
                Err.Raise eLibError, , "You need to install the following library before continuing: " & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
                    "Library: " & UCase(m.aDepLibs.Item(Y)) & Chr(13) & Chr(10) & _
                    "Version: " & UCase(m.aDepLibsVers.Item(Y)) & Chr(13) & Chr(10) & _
                    "LastMod: " & CDate(m.aDepLibsDate.Item(Y)) & Chr(13) & Chr(10) & Chr(13) & Chr(10) & _
                    "Please contact the library developer."
            Else
                'Dependent library not found (System Navigator)
                strResponse = InfBox("Version " & m.strPackVersion & " of|" & UCase(m.strPackName) & _
                        "|requires the library|" & UCase(m.aDepLibs.Item(Y)) & "|which was " & _
                        "not found on your system.||Importing this library could cause " & _
                        "undesired results.|Do you want to continue?|", , "-Yes|+No", "Confirmation")
                If strResponse = "N" Then Err.Raise eLibError, , "Library was not imported."
            End If
        Else
            'Version being installed is OLDER than version currently
            'installed on users machine
            'If Int(CDate(m.aDepLibsDate.Item(Y))) < Int(rs!LastModified) Then
            If CLng(m.aDepLibsVers.Item(Y)) < CLng(rs!Version) Then
                strResponse = InfBox("You currently have version " & rs!Version & " of|" & _
                        UCase(m.aDepLibs.Item(Y)) & ".|The library " & UCase(m.strPackName) & "|was " & _
                        "created using version " & m.aDepLibsVers.Item(Y) & ".||If " & _
                        "you encounter any problems importing this library, please " & _
                        "contact the developer of|" & UCase(m.strPackName), , , "Confirmation")
            
            'Version being installed is NEWER than the installed version.
            'If Int(CDate(m.aDepLibsDate.Item(Y))) > Int(rs!LastModified) Then
            ElseIf CLng(m.aDepLibsVers.Item(Y)) > CLng(rs!Version) Then
                strResponse = InfBox("You currently have version " & rs!Version & " of|" & _
                        UCase(m.aDepLibs.Item(Y)) & ".|The library " & UCase(m.strPackName) & "|was " & _
                        "created using version " & m.aDepLibsVers.Item(Y) & ".||You " & _
                        "should get an upgrade for|" & UCase(m.aDepLibs.Item(Y)) & "|before " & _
                        "importing this library.|Do you want to continue?|.", , "-Yes|+No", "Confirmation")
                If strResponse = "N" Then Err.Raise eLibError, , "Library was not imported."
            End If
        End If
    Next Y
    
ErrExit:
    Set rs = Nothing
    Exit Sub

ErrSection:
    RaiseError "cPackagedFile.ValidatePackagedLibrary", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitializeProgressBar
'' Description: Intialize the progress bar
'' Inputs:      Max Step for the Progress Bar
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitializeProgressBar(ByVal lMaxStep As Long)
On Error GoTo ErrSection:

    Set m.StatusBar = New cStatusBar
    m.StatusBar.StatusBar = m.vsStatusBar
    m.vsStatusBar.Visible = True
    m.vsStatusBar.ForeColor = vbBlack
    m.ctlStatusMsg.Visible = True
    m.lStep = 0
    m.lMaxStep = lMaxStep
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPackagedFile.InitializeProgressBar", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateMeter
'' Description: Update the progress bar
'' Inputs:      Message to put with the progress bar
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UpdateMeter(ByVal strMessage As String)
On Error GoTo ErrSection:

    m.lStep = m.lStep + 1
    If m.StatusBar.Value <> Int((m.lStep / m.lMaxStep) * 100) Then
        m.StatusBar.Value = Int((m.lStep / m.lMaxStep) * 100)
        m.StatusBar.Refresh
    End If
    m.ctlStatusMsg.Caption = strMessage
    m.ctlStatusMsg.Refresh

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cPackagedFile.UpdateMeter", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ReverifyFunctions
'' Description: Reverify all of the coded text functions in the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ReverifyFunctions()
On Error GoTo ErrSection:
    
    Dim rs As Recordset
    Dim Expr As Object
    Dim RTFText As String
    Dim EnglishText As String
    Dim F As Object
    Dim Functions As Object
    Dim Coloring As Object
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    Dim RichText As String
    Dim lPrevNumFunctions As Long
    Dim strSQL As String
    Dim Verifying As Boolean
    Dim TempInputs As Object
    Dim astrTemp As New cGdArray
    Dim dLastDoEvents As Double
    
    astrTemp.Create eGDARRAY_Longs
    astrTemp.Size = 0&
    
    ''RefreshDatabaseConnection
    Set g.CommonBridge = CreateObject(g.strCommonDLL & "cCommonBridge")
    With g.CommonBridge
        .AppPath = g.strAppPath
        .CustomerID = g.lLCD
        .dbNavRef = g.dbNav
    End With
    
    'Only select codedtext functions (not in the builtin libraries).
    'Set all functions to Reverify...
If UsingNewVerifyMode Then
    ' get functions which are not in Genesis-managed libraries
    Set rs = g.dbNav.OpenRecordset("SELECT tblFunctions.* " & _
        "FROM tblLibrarys INNER JOIN tblFunctions ON tblLibrarys.LibraryID = tblFunctions.LibraryID " & _
        "WHERE (tblLibrarys.Builtin=False) And (tblLibrarys.DLLSecurityCode >= 0);", dbOpenDynaset)
Else
    Set rs = g.dbNav.OpenRecordset("SELECT tblFunctions.* " & _
        "FROM tblLibrarys INNER JOIN tblFunctions ON tblLibrarys.LibraryID = tblFunctions.LibraryID " & _
        "WHERE (tblLibrarys.Builtin=False);", dbOpenDynaset)
End If
    ValidateCheckSums rs, "tblFunctions"
    If rs.BOF And rs.EOF Then Exit Sub
    Do Until rs.EOF
        If rs!CheckSum <> 0.5 Then
            rs.Edit
            rs!Reverify = True
            rs!CheckSum = BuildCheckSum(rs, "tblFunctions")
            rs.Update
        End If
        
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    Set TempInputs = CreateObject(g.strCommonDLL & "cInputs")
    Set Coloring = CreateObject(g.strTradeSenseOCX & "cColoring")

    'Loop n times in order to catch all nested functions
    lPrevNumFunctions = -1
    For X = 1 To 9999
        DoEvents
        dLastDoEvents = gdTickCount
        
        'Load all functions in tables into collection
        'Set Functions = New cFunctions
        Set Functions = CreateObject(g.strCommonDLL & "cFunctions")
        Functions.Load
        
        'Only select codedtext functions (not in the builtin libraries), only
        'select if functions need verifying...
        Set rs = g.dbNav.OpenRecordset("SELECT tblFunctions.* " & _
            "FROM tblLibrarys " & _
            "INNER JOIN tblFunctions " & _
            "ON tblLibrarys.LibraryID = tblFunctions.LibraryID " & _
            "WHERE (tblFunctions.Reverify=True) ORDER BY [FunctionName];", dbOpenDynaset)
        ValidateCheckSums rs, "tblFunctions"
        If rs.BOF And rs.EOF Then Exit For
        
        'If same # of functions to verify as last time through loop, then no need to continue
        rs.MoveLast
        If rs.RecordCount = lPrevNumFunctions Then Exit For
        lPrevNumFunctions = rs.RecordCount
    
        'Initialize progress meter...
        rs.MoveLast
        InitializeProgressBar rs.RecordCount
        rs.MoveFirst
            
        Set F = CreateObject(g.strCommonDLL & "cFunction")
        Do Until rs.EOF
            If rs!CheckSum <> 0.5 Then
                'Create instances to codedtext generator
                Set Expr = CreateObject(g.strTradeSenseOCX & "cExpression")
                With Expr
                    If g.CalledFrom = PortfolioNavigator Then
                        .PortfolioNavigator = True
                    Else
                        .PortfolioNavigator = False
                    End If
                    .Functions = Functions
                End With
            
                With F
                    .FunctionID = rs!FunctionID
                    If g.CalledFrom = PortfolioNavigator Then
                        .Load rs!FunctionID, True
                        Set TempInputs = .Inputs
                        
                        UpdateMeter "Pass " & Str(X) & ":  " & .FunctionName
                        If .ImplementationTypeID = 2 Then
                            EnglishText = .EditText
                            RichText = Coloring.Color(EnglishText)
                            Verifying = True
                            Expr.ValidateFunctionRule EnglishText
                            Verifying = False
                            .FunctionIDs = Expr.GetFIDs
                            .EditText = Expr.EditText
                            .Preview = Expr.Preview
                            .CodedText = Expr.CodedText
                            .Inputs = Expr.Inputs
                            For Z = 1 To .Inputs.Count
                                For Y = 1 To TempInputs.Count
                                    If TempInputs.Item(Y).ParmName = .Inputs.Item(Z).ParmName Then
                                        .Inputs.Item(Z).ParmSeq = TempInputs.Item(Y).ParmSeq
                                        .Inputs.Item(Z).ParmTypeID = TempInputs.Item(Y).ParmTypeID
                                        .Inputs.Item(Z).DefaultValue = TempInputs.Item(Y).DefaultValue
                                        .Inputs.Item(Z).ToValue = TempInputs.Item(Y).ToValue
                                        .Inputs.Item(Z).FromValue = TempInputs.Item(Y).FromValue
                                        .Inputs.Item(Z).Required = TempInputs.Item(Y).Required
                                        .Inputs.Item(Z).ParmDesc = TempInputs.Item(Y).ParmDesc
                                        Exit For
                                    End If
                                Next Y
                            Next Z
                        End If
                    Else
                        .Load
                        Set TempInputs = .Inputs
                        
                        'Update progress meter...
                        UpdateMeter "Pass " & Str(X) & ":  " & .FunctionName
                            
                        'Reverify codedtext
                        If .ImplementationTypeID = 2 Then
                            EnglishText = .Formatted
                            RichText = F.GetRTF(EnglishText)
                            Verifying = True
                            frmRTB.rtbAdvanced.TextRTF = RichText
                            Expr.ValidateFunctionRule frmRTB.rtbAdvanced.Text
                            Verifying = False
                            .Reverify = False
                            .FunctionIDs = Expr.GetFIDs
                            .Formatted = Expr.EditText
                            .FormattedWithFillWords = Expr.Preview
                            .CodedText = Expr.CodedText
                            .LateCalculating = Expr.LateCondition
                            .Inputs = Expr.Inputs
                            For Z = 1 To .Inputs.Count
                                For Y = 1 To TempInputs.Count
                                    If TempInputs.Item(Y).ParmName = .Inputs.Item(Z).ParmName Then
                                        .Inputs.Item(Z).ParmSeq = TempInputs.Item(Y).ParmSeq
                                        .Inputs.Item(Z).ParmTypeID = TempInputs.Item(Y).ParmTypeID
                                        .Inputs.Item(Z).DefaultValue = TempInputs.Item(Y).DefaultValue
                                        .Inputs.Item(Z).ToValue = TempInputs.Item(Y).ToValue
                                        .Inputs.Item(Z).FromValue = TempInputs.Item(Y).FromValue
                                        .Inputs.Item(Z).Required = TempInputs.Item(Y).Required
                                        .Inputs.Item(Z).ParmDesc = TempInputs.Item(Y).ParmDesc
                                        Exit For
                                    End If
                                Next Y
                            Next Z
                        Else
                            .FunctionIDs = astrTemp.ArrayHandle
                        End If
                    End If
                    
                    .Reverify = False
                    
                    Verifying = True
                    .Save
                    Verifying = False
                    
                    ' do this every few seconds so form will refresh
                    If gdTickCount > dLastDoEvents + 2000 Then
                        DoEvents
                        dLastDoEvents = gdTickCount
                    End If
                End With
            End If
NextFunction:
            rs.MoveNext
        Loop
        Set rs = Nothing
    Next X
    
ErrExit:
    Set F = Nothing
    Set Expr = Nothing
    Set TempInputs = Nothing
    Set Coloring = Nothing
    Set Functions = Nothing
    Set rs = Nothing
    Exit Sub

ErrSection:
    If (Err.Number < 0 Or UCase(Left(Err.Source, 5)) = "CLASS") And Verifying Then
        Resume NextFunction:
    Else
        RaiseError "cPackagedFile.ReverifyFunctions", eGDRaiseError_Raise, g.strAppPath
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ReverifyRules
'' Description: Reverify all of the coded text rules in the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ReverifyRules()
On Error GoTo ErrSection:
    
    Dim rs As Recordset
    Dim rs2 As Recordset
    Dim Expr As Object
    Dim Rule As Object
    Dim Functions As Object
    Dim X As Integer
    Dim Y As Integer
    Dim AdvancedRule As RichTextBox
    Dim Coloring As Object
    Dim Condition As String
    Dim EntryPrice As String
    Dim StopLimitPrice As String
    Dim OrderTypeInd As Byte
    Dim Buy As Byte
    Dim Verifying As Boolean
    Dim strRule As String
    Dim strOrderType As String
    Dim TempInputs As Object
    Dim bError As Boolean
    Dim dLastDoEvents As Double
    
    ''RefreshDatabaseConnection
    Set g.CommonBridge = CreateObject(g.strCommonDLL & "cCommonBridge")
    With g.CommonBridge
        .AppPath = g.strAppPath
        .CustomerID = g.lLCD
        .dbNavRef = g.dbNav
    End With
    
    'Load all functions...
    'Set Functions = New cFunctions
    Set TempInputs = CreateObject(g.strCommonDLL & "cInputs")
    Set Functions = CreateObject(g.strCommonDLL & "cFunctions")
    Functions.Load
    
    'Create instances to codedtext generator
    Set Coloring = CreateObject(g.strTradeSenseOCX & "cColoring")

    If m.bABuiltinWasImported Then
        Set rs = g.dbNav.OpenRecordset("SELECT tblRules.* " & _
            "FROM tblLibrarys INNER JOIN tblRules ON tblLibrarys.LibraryID = tblRules.LibraryID " & _
            "ORDER BY [Name];", dbOpenDynaset)
    Else
        Set rs = g.dbNav.OpenRecordset("SELECT tblRules.* " & _
            "FROM tblLibrarys INNER JOIN tblRules ON tblLibrarys.LibraryID = tblRules.LibraryID " & _
            "WHERE (((tblLibrarys.Builtin)=False))" & _
            "ORDER BY [Name];", dbOpenDynaset)
    End If
    ValidateCheckSums rs, "tblRules"
    
    'Initialize progress meter...
    rs.MoveLast
    InitializeProgressBar rs.RecordCount
    rs.MoveFirst
    
    Set Rule = CreateObject(g.strCommonDLL & "cRule")
    
    Do Until rs.EOF
        If rs!CheckSum <> 0.5 Then
            'Update progress meter...
            If g.CalledFrom = SystemNavigator Then
                UpdateMeter rs!Name
            Else
                UpdateMeter rs!RuleName
            End If
            
            ' do this every few seconds so form will refresh
            If gdTickCount > dLastDoEvents + 2000 Then
                DoEvents
                dLastDoEvents = gdTickCount
            End If
                        
            'Create instances to codedtext generator
            ' DAJ 05/08/2012: Need to do this every time through the loop to clear out the
            ' expression's array of macros.  On reverify, any rule that had a macro with the
            ' same name as a macro in another rule would get expanded with the coded text
            ' of the macro from the previous rule...
            Set Expr = CreateObject(g.strTradeSenseOCX & "cExpression")
            With Expr
                If g.CalledFrom = SystemNavigator Then
                    .PortfolioNavigator = False
                Else
                    .PortfolioNavigator = True
                End If
                .Functions = Functions
            End With
            
            If g.CalledFrom = PortfolioNavigator Then
                With Rule
                    .Functions = Functions
                    .RuleID = rs!RuleID
                    .Load rs!RuleID, True
                    Set TempInputs = .Inputs
                    .Reverify = False
                    
                    'Save Default value of inputs
                    'Set TempInputs = CreateObject(g.strCommonDLL & "cInputs")
                    'For X = 1 To .Inputs.Count
                    '    With TempInputs.Item(X)
                    '        .ParmName = .Inputs.Item(X).ParmName
                    '        .DefaultValue = .Inputs.Item(X).DefaultValue
                    '    End With
                    'Next X
                    
                    'Move rule to Rich text box to get english version of text
                    'which is then reverified...
                    frmRTB.rtbAdvanced = Coloring.Color(.EditText)
                    Verifying = True
                    Expr.ValidateRuleAdv frmRTB.rtbAdvanced.Text
                    Verifying = False
                    Unload frmRTB
    
                    'Save verified results...
                    .FunctionIDs = Expr.GetFIDs
                    .Preview = Expr.Preview
                    .CodedText = Expr.CodedText
                    
                    'Restore input default values
    
                    .Inputs = Expr.Inputs
                    For X = 1 To .Inputs.Count
                        For Y = 1 To TempInputs.Count
                            If TempInputs.Item(Y).ParmName = .Inputs.Item(X).ParmName Then
                                .Inputs.Item(X).DefaultValue = TempInputs.Item(Y).DefaultValue
                                Exit For
                            End If
                        Next Y
                    Next X
                    
                    'Set on special flag to force validation in cRule to be
                    'bypassed:  It normally checks to make sure the last mode date
                    'of the rule is greater than each of it's functions.  But
                    'since they were reverified before this routine was called,
                    'they are always more recent.
                    .ImportLibraryMode = True
                    .Save
                End With
            Else
                'Get Rule...
                'Set Rule = New cRule
                With Rule
                    .RuleID = rs!RuleID
                    .Load
                    Set TempInputs = .Inputs
                    
                    ' Only reverify those rules that we have to
                    If .Reverify = True Or m.bABuiltinWasImported Then
                        'Put together pieces of rule into advanced rule format and verify
                        frmRTB.rtbAdvanced = .GetRTF(EZToAdvanced(.BuySell, .OrderPlacement, _
                                                    .Cond, .Price, .Price2RTF))
                        Verifying = True
                        Expr.ValidateRule frmRTB.rtbAdvanced.Text
                        Verifying = False
                        Unload frmRTB
                
                        'Parse Editor codedtext (advanced format) for tblRules...
                        AdvancedToEZ Expr.EditText, Buy, OrderTypeInd, _
                            Condition, EntryPrice, StopLimitPrice, strOrderType
                            
                        'Save results...
                        .LateCondition = Expr.LateCondition
                        .LateAction = Expr.LateAction
                        .Cond = Condition
                        .Price = EntryPrice
                        .Price2RTF = StopLimitPrice
                        .OrderPlacement = strOrderType 'OrderTypeInd
                        
                        'Parse Engine Codedtext for tblRules...
                        AdvancedToEZ Expr.CodedText, Buy, OrderTypeInd, Condition, _
                            EntryPrice, StopLimitPrice, strOrderType
                    
                        'Save results...
                        .CondCoded = Condition
                        If EntryPrice = "" Then
                            .PriceCoded = "N/A"
                        Else
                            .PriceCoded = EntryPrice
                        End If
                    
                        If StopLimitPrice = "" Then
                            .Price2Coded = "N/A"
                        Else
                            .Price2Coded = StopLimitPrice
                        End If
                        
                        'Returns the function ID's used in the entire rule...
                        .CondRefs = Expr.GetFIDs
                    
                        .Inputs = Expr.Inputs
                        For X = 1 To .Inputs.Count
                            For Y = 1 To TempInputs.Count
                                If TempInputs.Item(Y).ParmName = .Inputs.Item(X).ParmName Then
                                    .Inputs.Item(X).DefaultValue = TempInputs.Item(Y).DefaultValue
                                    .Inputs.Item(X).ToValue = TempInputs.Item(Y).ToValue
                                    .Inputs.Item(X).FromValue = TempInputs.Item(Y).FromValue
                                    Exit For
                                End If
                            Next Y
                        Next X
                        
                        .Reverify = False
                        .Save
                    End If
                End With
            End If
        End If
        bError = False

NextRule:
        If bError Then
            rs.Edit
            rs!Reverify = True
            rs!CheckSum = BuildCheckSum(rs, "tblRules")
            rs.Update
        End If
        rs.MoveNext
    Loop
    
ErrExit:
    Set Rule = Nothing
    Set Expr = Nothing
    Set Functions = Nothing
    Set TempInputs = Nothing
    Set rs = Nothing
    ''CloseDb
    Exit Sub

ErrSection:
    If Err.Number < 0 And Verifying = True Then
        bError = True
        Resume NextRule:
    Else
        RaiseError "cPackagedFile.ReverifyRules", eGDRaiseError_Raise, g.strAppPath
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ReverifySystems
'' Description: System is verified if each of its rules are verified
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ReverifySystems()
On Error GoTo ErrSection:

    Dim rsSystems As Recordset          ' Recordset of systems from the database
    Dim rsRules As Recordset            ' Recordset of rules from each system
    Dim bReverify As Boolean            ' Does the system need to be reverified?
    
    Set rsSystems = g.dbNav.OpenRecordset("SELECT * FROM [tblSystems] " & _
            "ORDER BY [SystemNumber];", dbOpenDynaset)
    
    'Initialize progress meter...
    rsSystems.MoveLast
    InitializeProgressBar rsSystems.RecordCount
    rsSystems.MoveFirst
    
    Do While Not rsSystems.EOF
        UpdateMeter "Verifying Strategy: " & UCase(rsSystems!SystemName)
        
        If rsSystems!CheckSum <> 0.5 Then
            bReverify = False
            
            Set rsRules = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] " & _
                    "WHERE [SystemNumber]=" & Str(rsSystems!SystemNumber) & " " & _
                    "ORDER BY [RuleID];", dbOpenDynaset)
            Do While Not rsRules.EOF
                If rsRules!Reverify = True Then
                    bReverify = True
                    Exit Do
                End If
                rsRules.MoveNext
            Loop
            
            rsSystems.Edit
            rsSystems!Reverify = bReverify
            rsSystems!CheckSum = BuildCheckSum(rsSystems, "tblSystems")
            rsSystems.Update
        End If
        
        rsSystems.MoveNext
    Loop

ErrExit:
    Set rsSystems = Nothing
    Set rsRules = Nothing
    Exit Sub
    
ErrSection:
    Set rsSystems = Nothing
    Set rsRules = Nothing
    RaiseError "cPackagedFile.ReverifySystems", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ReverifyPortRules
'' Description: Reverify all of the portfolio rules in the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ReverifyPortRules()
On Error GoTo ErrSection:
    
    Dim Portfolios As Object
    Dim Expr As Object
    Dim Rules As Object
    Dim Functions As Object
    Dim X As Integer
    Dim Y As Integer
    Dim a As Long
    Dim b As Long
    Dim Coloring As Object
    Dim Verifying As Boolean
    Dim TempInputs As Object
    Dim Cnt As Long
    
    ''RefreshDatabaseConnection
    Set g.CommonBridge = CreateObject(g.strCommonDLL & "cCommonBridge")
    With g.CommonBridge
        .AppPath = g.strAppPath
        .CustomerID = g.lLCD
        .dbNavRef = g.dbNav
    End With
    
    'Load all functions...
    'Set Functions = New cFunctions
    Set TempInputs = CreateObject(g.strCommonDLL & "cInputs")
    Set Functions = CreateObject(g.strCommonDLL & "cFunctions")
    Functions.Load
    
    'Create instances to codedtext generator
    Set Expr = CreateObject(g.strTradeSenseOCX & "cExpression")
    With Expr
        .PortfolioNavigator = True
        .Functions = Functions
    End With
    Set Coloring = CreateObject(g.strTradeSenseOCX & "cColoring")
    
    Set Portfolios = CreateObject(g.strCommonDLL & "cPortfolios")
    Portfolios.Load
    For X = 1 To Portfolios.Count
        With Portfolios.Item(X)
            .Load .PortfolioNumber, False, False, True, True
            For Y = 1 To .Rules.Count
                Cnt = Cnt + 1
            Next Y
        End With
    Next X
    InitializeProgressBar CInt(Cnt)
    
    For a = 1 To Portfolios.Count
        For b = 1 To Portfolios.Item(a).Rules.Count
            With Portfolios.Item(a).Rules.Item(b)
                UpdateMeter .RuleName
            
                .Functions = Functions
                Set TempInputs = .Inputs
                .Reverify = True
                
                'Move rule to Rich text box to get english version of text
                'which is then reverified...
                frmRTB.rtbAdvanced = Coloring.Color(.EditText)
                Verifying = True
                Expr.ValidateRuleAdv frmRTB.rtbAdvanced.Text
                Verifying = False
                .Reverify = False
                Unload frmRTB

                'Save verified results...
                .FunctionIDs = Expr.GetFIDs
                .Preview = Expr.Preview
                .CodedText = Expr.CodedText
                .Inputs = Expr.Inputs
                For X = 1 To .Inputs.Count
                    For Y = 1 To TempInputs.Count
                        If TempInputs.Item(Y).ParmName = .Inputs.Item(X).ParmName Then
                            .Inputs.Item(X).DefaultValue = TempInputs.Item(Y).DefaultValue
                            .Inputs.Item(X).ToValue = TempInputs.Item(Y).ToValue
                            .Inputs.Item(X).FromValue = TempInputs.Item(Y).FromValue
                            Exit For
                        End If
                    Next Y
                Next X
            End With
NextRule:
        Next b
        Portfolios.Item(a).functionsglobal = Functions
        Portfolios.Item(a).SaveRules True
    Next a
    
ErrExit:
    Set Portfolios = Nothing
    Set Rules = Nothing
    Set Expr = Nothing
    Set Functions = Nothing
    Set TempInputs = Nothing
    ''CloseDb
    Exit Sub
    
ErrSection:
    If (Err.Number < 0 Or Left(Err.Source, 5) = "Class") And Verifying Then
        RaiseError "cPackagedFile.ReverifyPortRules", eGDRaiseError_Show, g.strAppPath
        Resume NextRule
    Else
        RaiseError "cPackagedFile.ReverifyPortRules", eGDRaiseError_Raise, g.strAppPath
    End If
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ReverifyModelRules
'' Description: Reverify all of the model rules in the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ReverifyModelRules()
On Error GoTo ErrSection:
    
    Dim Models As Object
    Dim Expr As Object
    Dim Rules As Object
    Dim Functions As Object
    Dim X As Integer
    Dim Y As Integer
    Dim a As Long
    Dim b As Long
    Dim Coloring As Object
    Dim Verifying As Boolean
    Dim TempInputs As Object
    Dim Cnt As Long
    
    ''RefreshDatabaseConnection
    Set g.CommonBridge = CreateObject(g.strCommonDLL & "cCommonBridge")
    With g.CommonBridge
        .AppPath = g.strAppPath
        .CustomerID = g.lLCD
        .dbNavRef = g.dbNav
    End With
    
    'Load all functions...
    'Set Functions = New cFunctions
    Set TempInputs = CreateObject(g.strCommonDLL & "cInputs")
    Set Functions = CreateObject(g.strCommonDLL & "cFunctions")
    Functions.Load
    
    'Create instances to codedtext generator
    Set Expr = CreateObject(g.strTradeSenseOCX & "cExpression")
    With Expr
        .PortfolioNavigator = True
        .Functions = Functions
    End With
    Set Coloring = CreateObject(g.strTradeSenseOCX & "cColoring")
    
    Set Models = CreateObject(g.strCommonDLL & "cModels")
    Models.Load
    For X = 1 To Models.Count
        With Models.Item(X)
            .Load .ModelNumber, True, True
            For Y = 1 To .Rules.Count
                Cnt = Cnt + 1
            Next Y
        End With
    Next X
    InitializeProgressBar CInt(Cnt)
    
    For a = 1 To Models.Count
        For b = 1 To Models.Item(a).Rules.Count
            With Models.Item(a).Rules.Item(b)
                UpdateMeter .RuleName
            
                .Functions = Functions
                Set TempInputs = .Inputs
                .Reverify = True
                
                'Move rule to Rich text box to get english version of text
                'which is then reverified...
                frmRTB.rtbAdvanced = Coloring.Color(.EditText)
                Verifying = True
                Expr.ValidateRuleAdv frmRTB.rtbAdvanced.Text
                Verifying = False
                .Reverify = False
                Unload frmRTB

                'Save verified results...
                .FunctionIDs = Expr.GetFIDs
                .Preview = Expr.Preview
                .CodedText = Expr.CodedText
                .Inputs = Expr.Inputs
                For X = 1 To .Inputs.Count
                    For Y = 1 To TempInputs.Count
                        If TempInputs.Item(Y).ParmName = .Inputs.Item(X).ParmName Then
                            .Inputs.Item(X).DefaultValue = TempInputs.Item(Y).DefaultValue
                            .Inputs.Item(X).ToValue = TempInputs.Item(Y).ToValue
                            .Inputs.Item(X).FromValue = TempInputs.Item(Y).FromValue
                            Exit For
                        End If
                    Next Y
                Next X
            End With
NextRule:
        Next b
        Models.Item(a).functionsglobal = Functions
        Models.Item(a).SaveRules True
    Next a
    
ErrExit:
    Set Models = Nothing
    Set Rules = Nothing
    Set Expr = Nothing
    Set Functions = Nothing
    Set TempInputs = Nothing
    ''CloseDb
    Exit Sub
    
ErrSection:
    If (Err.Number < 0 Or Left(Err.Source, 5) = "Class") And Verifying Then
        RaiseError "cPackagedFile.ReverifyModelRules", eGDRaiseError_Show, g.strAppPath
        Resume NextRule
    Else
        RaiseError "cPackagedFile.ReverifyModelRules", eGDRaiseError_Raise, g.strAppPath
    End If
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AdvancedToEZ
'' Description: Parse Coded Text into individual parts
'' Inputs:      Coded Text, Buy/Sell, Order Type, Condition, Entry Price,
''              Stop Limit Price, Order Type String
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Codedtext in Advanced Format is passed here.  It is parsed and returns
'Condition, Order Type, Entry Price and Stop/Limit price
Private Sub AdvancedToEZ(pCodedText As String, _
    pBuy As Byte, pOrderType As Byte, _
    pCondition As String, pEntryPrice As String, pStopLimitPrice As String, _
    pstrOrderType As String)
On Error GoTo ErrSection:
    
    Dim Parens As Integer
    Dim Commas As Integer
    Dim X As Long
    Dim curPos As Long
    Dim IfPos As Long
    Dim EnterPos As Long
    Dim EntryPos As Long
    Dim CommaPos1 As Long
    Dim CommaPos2 As Long
    Dim RParenPos As Long
    Dim ParmStart As Long
    Dim OrderType As String
    Dim wrkText As String
    Dim ThenPos As Long
    
    'System Navigator tokens
    Const cIf = 24
    Const cThen = 35
    Const cEnter = 80
    Const cFLParen = 16
    Const cFRParen = 17
    Const cComma = 22
    
    'Condition...
    IfPos = InStr(1, pCodedText, "~" & Format(cIf, "00"))
    ThenPos = InStr(IfPos + 1, pCodedText, "~" & Format(cThen, "00"))
    'EnterPos = InStr(IfPos + 1, pCodedText, "~" & Format(cEnter, "00"))
    EnterPos = InStr(ThenPos + 1, pCodedText, "~" & Format(cEnter, "00"))
    If IfPos = 0 Or EnterPos = 0 Or ThenPos = 0 Then Exit Sub
    pCondition = Left(pCodedText, ThenPos - 1)
    
    'Determine if Buy or Sell...
    If InStr(EnterPos + 1, pCodedText, "Buy") > 0 Then
        pBuy = 1
    Else
        pBuy = 0
    End If
    
    'Entry Price...
    EntryPos = InStr(EnterPos + 1, pCodedText, "~" & Format(cFLParen, "00"))
    If EntryPos = 0 Then
        'edAdvanced.TurnOffEditing
        'edAdvanced.ExprIsFormatted = True
        'wrkText = pCodedtext
        'edAdvanced.TextRTF = mRule.GetRTF(wrkText)
        Exit Sub
    End If
    
    'Next token is the start of Entry Price (following Function LeftParen)
    EntryPos = InStr(EntryPos + 1, pCodedText, "~")
    
    'If EntryPos is Market1 then ignore hidden parm and adjust entry
    'by skipping market and the comma following it
    If Mid(pCodedText, EntryPos + 6, 7) = "Market1" Then
        EntryPos = InStr(EntryPos + 21, pCodedText, "~")
    End If
    
    'Determine position's of comma's in Buy/Sell action...
    Parens = 1
    Commas = 0
    curPos = EntryPos + 1
    Do Until curPos > Len(pCodedText)
        curPos = InStr(curPos, pCodedText, "~")
        If curPos = 0 Then Exit Do
        Select Case Val(Mid(pCodedText, curPos + 1, 2))
            Case cFLParen: Parens = Parens + 1
            Case cFRParen
                Parens = Parens - 1
                If Parens = 0 Then
                    RParenPos = curPos
                End If
            
            Case cComma
                If Parens = 1 Then
                    Commas = Commas + 1
                    If Commas = 1 Then
                        CommaPos1 = curPos
                    Else
                        CommaPos2 = curPos
                    End If
                End If
        End Select
        curPos = curPos + 1
    Loop
    
    'Entry Price...
    pEntryPrice = Mid(pCodedText, EntryPos, CommaPos1 - EntryPos)
    
    'Order Type...
    ParmStart = InStr(CommaPos1 + 1, pCodedText, "~")
    If CommaPos2 > 0 Then
        OrderType = Mid(pCodedText, ParmStart, _
            CommaPos2 - ParmStart)
    Else
        OrderType = Mid(pCodedText, ParmStart, _
            RParenPos - ParmStart)
    End If
    OrderType = Mid(OrderType, 8, Val(Mid(OrderType, 4, 3)) - 2)
    Select Case OrderType
        Case "Market": pOrderType = C_MARKET
        Case "Stop": pOrderType = C_STOP
        Case "Limit": pOrderType = C_LIMIT
        Case "Stop with Limit": pOrderType = C_STOPLIMIT
        Case "Market on Close": pOrderType = C_MARKETCLOSE
        Case "Stop Close Only": pOrderType = C_STOPCLOSE
        Case "Limit Close Only": pOrderType = C_LIMITCLOSE
        Case "Stop with Limit Close Only": pOrderType = C_STOPLIMITCLOSE
        Case Else
            pOrderType = C_MARKET
    End Select
    pstrOrderType = OrderType
    
    'Stop/Limit price
    ParmStart = InStr(CommaPos2 + 1, pCodedText, "~")
    If CommaPos2 > 0 Then
        pStopLimitPrice = Mid(pCodedText, ParmStart, _
            RParenPos - ParmStart)
    End If
        
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cPackagedFile.AdvancedToEZ", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EZToAdvanced
'' Description: Combine all parts into Coded Text
'' Inputs:      Buy/Sell, Order Placement, Condition, Price1, Price2
'' Returns:     Coded Text
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function EZToAdvanced(ByVal bBuySell As Boolean, ByVal strOrderPlacement As String, _
    ByVal strCond As String, ByVal strPrice1 As String, ByVal strPrice2 As String) As String
On Error GoTo ErrSection:

    Dim strTemp As String

    'Condition, Then followed by <enterkey>
    strTemp = Trim(strCond) & " ~24005THEN ~80003<E> ~44005<Tab> "

    'Order action (Buy/Sell)
    If bBuySell Then
        strTemp = strTemp & "~24003BUY ~16001( "
    Else
        strTemp = strTemp & "~24004SELL ~16001( "
    End If
    
    'Entry Price plus a comma...
    'Price may be blank for existing rules (before the new tradesense merge).  In
    'these cases, default the price coded text for market orders.
    If UCase(strOrderPlacement) = "MARKET" Or UCase(strOrderPlacement) = "MARKET ON CLOSE" Then
        strTemp = strTemp & "~01013Next Bar Open ~22001, "
    Else
        strTemp = strTemp & Trim(strPrice1) & " ~22001, "
    End If
    
    'Order type
    strTemp = strTemp & "~20" & Format(Len(strOrderPlacement) + 2, "000") & _
        """" & strOrderPlacement & """"
    
    'Stop/Limit Price
    If (UCase(strOrderPlacement) = "STOP WITH LIMIT") Or (UCase(strOrderPlacement) = "STOP WITH LIMIT CLOSE ONLY") Then
        If Len(strPrice2) > 0 And strPrice2 <> "~00000" Then
            strTemp = strTemp & " ~22001, " & strPrice2
        End If
    End If
    
    'Function right paren...
    strTemp = Trim(strTemp) & " ~17001) ~80003<E> ~43005EndIf ~80003<E>"
    
    'Return reformatting rule...
    EZToAdvanced = strTemp
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPackagedFile.EZToAdvanced", eGDRaiseError_Raise, g.strAppPath
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Cleanup
'' Description: Do some cleanup
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Cleanup()
On Error GoTo ErrSection:
    
    m.vsStatusBar.Visible = False
    m.ctlStatusMsg.Visible = False
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPackagedFile.Cleanup", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize class members
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    If Len(m.ExportHeader) <> 80 Then
        Err.Raise vbObjectError + 1000, , "Program Error!!"
    Else
        With m.ExportHeader
            .Reserved1 = String(Len(.Reserved1), Chr(0))
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPackagedFile.Initialize", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Cleanup after ourselves upon termination
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
    
    Set m.StatusBar = Nothing
    Set m.vsStatusBar = Nothing
    Set m.QryDef = Nothing
    Set m.aDepLibs = Nothing
    Set m.aDepLibsDate = Nothing
    Set m.aDepLibsVers = Nothing
    Set m.astrExportFile = Nothing
    Set m.aMsg = Nothing

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ImportRuleFunctions
'' Description: Import the Rule Functions table from the import file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ImportRuleFunctions()
On Error GoTo ErrSection:
    
    Dim X As Long
    Dim rsRules As Recordset
    Dim rsChildFunctions As Recordset
    Dim rsRefs As Recordset
    Dim strOuputLine As String
    Dim lIndex As Long
    
    If m.bBuiltin Or m.bGenesisManaged Then
#If 0 Then
        lIndex = FindSection("tblFunctionRules")
        
        Set rsRules = g.dbNav.OpenRecordset("Select * from [tblRules] " & _
            " Where [LibraryID]=" & Str(m.lLibraryID) & _
            " Order by [LibraryID],[Name];", dbOpenDynaset)
        
        Set rsChildFunctions = g.dbNav.OpenRecordset("Select * from [tblFunctions] " & _
            " Order by [CodedName];", dbOpenDynaset)
            
        Set rsRefs = g.dbNav.OpenRecordset("Select * from [tblFunctionRules] " & _
            " Order by [RuleID],[FunctionIDRef];", dbOpenDynaset)
        
        'Do Until EOF(1)
        Do Until lIndex > m.astrExportFile.Size - 1
        
            'Get detail record to post...
            'Line Input #1, strOuputLine
            strOuputLine = m.astrExportFile(lIndex)
            strOuputLine = ReplaceChars(strOuputLine, "||", Chr(13) & Chr(10))
            m.b = Split(strOuputLine, vbTab)
            
            'Exit if the line retrieve is the next section header
            If Left(m.b(0), 1) = "[" Then Exit Do
            
            'Find the RuleID
            rsRules.FindFirst "[Name]='" & m.b(0) & "'"
            
            'Find the Sub FunctionID
            rsChildFunctions.FindFirst "[CodedName]='" & m.b(1) & "'"
            If rsChildFunctions.NoMatch Then
                m.aMsg.Add "Function: " & m.b(1) & " used by " & _
                    "Rule: " & UCase(m.b(0)) & " was not found."
                If Not rsRules.NoMatch Then
                    With rsRules
                        .Edit
                        !Reverify = True
                        .Update
                    End With
                End If
            Else
                'Build reference to function
                With rsRefs
                    .FindFirst "[RuleID]=" & rsRules!RuleID & _
                            " and [FunctionIDRef]=" & rsChildFunctions!FunctionID
                    If .NoMatch Then
                        .AddNew
                        !RuleID = rsRules!RuleID
                        !functionidref = rsChildFunctions!FunctionID
                        .Update
                    End If
                End With
            End If
            
            lIndex = lIndex + 1
        Loop
#End If
    End If
    
ErrExit:
    Set rsRules = Nothing
    Set rsChildFunctions = Nothing
    Set rsRefs = Nothing
    Exit Sub

ErrSection:
    RaiseError "cPackagedFile.ImportRuleFunctions", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LibraryExists
'' Description: Determines whether a given library exists in the database
'' Inputs:      Library to check
'' Returns:     True if exists in the database, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LibraryExists(strLibraryName As String) As Boolean
On Error GoTo ErrSection:
    
    Dim rs As Recordset
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblLibrarys] " & _
            "WHERE [LibraryName]='" & strLibraryName & "';", dbOpenDynaset)
    
    If rs.EOF Then
        LibraryExists = False
    Else
        LibraryExists = True
    End If

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cPackagedFile.LibraryExists", eGDRaiseError_Raise, g.strAppPath
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LibrarysExists
'' Description: Determines whether a given library exists in the database
'' Inputs:      Library to check
'' Returns:     True if exists in the database, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LibrarysExists(strLibraryName As String, strSecondLib As String) As Boolean
On Error GoTo ErrSection:
    
    Dim rs As Recordset
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblLibrarys] " & _
            "WHERE [LibraryName]='" & strLibraryName & "' OR [LibraryName]='" & _
            strSecondLib & "';", dbOpenDynaset)
    
    If rs.EOF Then
        LibrarysExists = False
    Else
        LibrarysExists = True
    End If

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cPackagedFile.LibrarysExist", eGDRaiseError_Raise, g.strAppPath
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CloseDb
'' Description: Close the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CloseDb()
On Error GoTo ErrSection:

    If Not g.dbNav Is Nothing Then
        DoEvents
        g.dbNav.Close
        Set g.dbNav = Nothing
        DoEvents
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPackagedFile.CloseDb", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadLibraryFile
'' Description: Load the library file into the big string array
'' Inputs:      Whether or not to copy/register an associated DLL
'' Returns:     True on success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LoadLibraryFile(ByVal bJustInfo As Boolean, ByVal bCanCopy As Boolean) As Boolean
On Error GoTo ErrSection:

    Dim lCRC As Long                    ' CRC from the exported file
    Dim lCRC2 As Long                   ' CRC from unzipping the string
    Dim lZipLen As Long                 ' Length of the zipped string
    Dim lBigLen As Long                 ' Length of the unzipped string
    Dim lVersion As Long                ' Version of the export file
    Dim strPath As String               ' Path of the DLL
    Dim i&
    Dim lCreatorID As Long              ' Customer ID of the library's creator
    
    Dim mb As New cMemBuffer
    Dim mbZip As New cMemBuffer
    Dim mbPassword As New cMemBuffer
    Dim mbSections() As cMemBuffer
    
    ReDim mbSections(0) As cMemBuffer
    LoadLibraryFile = False
    
    Set m.astrExportFile = New cGdArray
    m.astrExportFile.Create eGDARRAY_Strings
            
    If Right(m.strPath, 4) <> ".GLB" Then Err.Raise vbObjectError + 1000, , m.strPath & " could not be loaded"
        
    ' Get the file header information...
    If Not mb.FromFile(m.strPath) Then
        Err.Raise vbObjectError + 1000, , m.strPath & " could not be loaded"
    End If
    lVersion = mb.GetLong(0)
    lCRC = mb.GetLong
    lBigLen = mb.GetLong
    lZipLen = mb.GetLong
    
    Select Case lVersion
        Case 1
            strPath = FileBase(m.strPath) & "." & FileExt(m.strPath)
            Err.Raise vbObjectError + 1000, , strPath & "||This library file was made with System Navigator " & _
                    "and cannot be imported.  Please contact Genesis for more " & _
                    "information."
        Case Is <> kVersion
            strPath = FileBase(m.strPath) & "." & FileExt(m.strPath)
            Err.Raise vbObjectError + 1000, , strPath & "||The version of the library file is not recognized " & _
                    "by this version of the software and cannot be imported.  " & _
                    "Please contact Genesis for more information."
    End Select
    
    ' decrypt and unzip the big string
    mbZip.PutFromMemory mb.MemPtr + 16, mb.Length - 16, 0
    mbPassword.Buffer = "LibMgr2"
    gdEncrypt False, mbZip, mbPassword
    mb.Clear
    mb.Length = lBigLen + 1024 '(NOTE: must have a little extra space)
    UnzipMemory mbZip.MemPtr, mbZip.Length, mb.MemPtr, mb.Length, lCRC2
    mb.Length = lBigLen
    
    If lCRC = lCRC2 Then
        ' Break the big string into sections...
        If BreakIntoSections(mb, mbSections) > 0 Then
            lCreatorID = 0&
            If LibSection(eLibSection_CreatorID) <= UBound(mbSections) Then
                If mbSections(LibSection(eLibSection_CreatorID)).Length = 4 Then
                    lCreatorID = mbSections(LibSection(eLibSection_CreatorID)).GetLong(0)
                End If
            End If
            
            ' 1) Check the Expiration Date on the library...
            m.ExportHeader.ExpDate = 0
            If (lCreatorID = 0&) Or (lCreatorID <> g.lLCD) Then
                If LibSection(eLibSection_ExpirationDate) <= UBound(mbSections) Then
                    If mbSections(LibSection(eLibSection_ExpirationDate)).Length = 4 Then
                        m.ExportHeader.ExpDate = mbSections(LibSection(eLibSection_ExpirationDate)).GetLong(0)
                    End If
                End If
                If m.ExportHeader.ExpDate > 0 And m.ExportHeader.ExpDate < JulFromLong(g.lHonestDate) Then
                    Err.Raise vbObjectError + 1000, , "This file cannot be Imported because it has Expired"
                End If
            End If
            
            ' 2) Check the Customer ID associated with the library...
            m.strCustomerID = ""
            If (lCreatorID = 0&) Or (lCreatorID <> g.lLCD) Then
                If LibSection(eLibSection_CustomerID) <= UBound(mbSections) Then
                    m.strCustomerID = mbSections(LibSection(eLibSection_CustomerID)).Buffer
                End If
                If Len(m.strCustomerID) > 0 Then
                    If InStr("," & m.strCustomerID & ",", "," & CStr(g.lLCD) & ",") = 0 Then
                        Err.Raise vbObjectError + 1000, , "You are not authorized to Import this File"
                    End If
                End If
            End If
            
            ' 3) Handle the library tables...
            m.astrExportFile.SplitFields Replace(mbSections(LibSection(eLibSection_Library)).Buffer, vbCrLf, vbLf), vbLf
            For i = 0 To m.astrExportFile.Size - 1
                m.astrExportFile(i) = Replace(m.astrExportFile(i), "||", vbCrLf)
            Next i
        
            ' 4) Handle the associated DLL (if applicable)...
            If Not bJustInfo And LibSection(eLibSection_DLL) <= UBound(mbSections) Then
                Select Case ExtractDLL(mbSections(LibSection(eLibSection_DLL)), bCanCopy)
                    Case 0:
                        ' Error
                    Case 1:
                        Err.Raise eLibError, , m.strPath & " could not be imported " & _
                                "because it needs to replace a DLL that is already " & _
                                "loaded.  This library will automatically import the next " & _
                                "time that you start the program."
                End Select
            End If
            
            ' 5) Handle any other included files...
            If Not bJustInfo And LibSection(eLibSection_IncludeFiles) <= UBound(mbSections) Then
                ExtractIncludeFiles mbSections(LibSection(eLibSection_IncludeFiles))
            End If
        
        End If
    Else
        Err.Raise vbObjectError + 1000, , m.strPath & " failed CRC Check"
    End If
    
    LoadLibraryFile = True
    
ErrExit:
    For i = 0 To UBound(mbSections)
        Set mbSections(i) = Nothing
    Next
    Exit Function
    
ErrSection:
    RaiseError "cPackagedFile.LoadLibraryFile", eGDRaiseError_Raise, g.strAppPath
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixSystemNumbers
'' Description: Replace the bogus system numbers with the real ones
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FixSystemNumbers()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset from the database
    Dim rs2 As Recordset                ' Recordset from the database
    Dim lSystemNumber As Long           ' Real System Number for the given name
    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 0 To m.astrSystemName.Size - 1
        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblSystems] " & _
                    "WHERE [SystemName]='" & m.astrSystemName(lIndex) & "';", dbOpenDynaset)
        If Not rs.EOF Then
            lSystemNumber = rs!SystemNumber
            
            Set rs2 = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] " & _
                        "WHERE [SystemNumber]=" & Str(m.alBogusID(lIndex)) & ";", dbOpenDynaset)
            Do While Not rs2.EOF
                rs2.Edit
                rs2!SystemNumber = lSystemNumber
                rs2!CheckSum = BuildCheckSum(rs2, "tblRules")
                rs2.Update
                rs2.MoveNext
            Loop
        End If
    Next lIndex
    
ErrExit:
    Set rs = Nothing
    Set rs2 = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "cPackagedFile.FixSystemNumbers", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DatabaseVersionOK
'' Description: Make sure that the user cannot import a library that was
''              exported with a newer version of the database than what the
''              user has
'' Inputs:      None
'' Returns:     True if Database Version is OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DatabaseVersionOK() As Boolean
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset to get from the database
    Dim lIndex As Long                  ' Index of the database version in array
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblDatabase];", dbOpenSnapshot)
    If Not rs.EOF Then
        lIndex = FindSection("DatabaseVersion")
        If CLng(m.astrExportFile(lIndex)) > rs!Version Then
            DatabaseVersionOK = False
            Err.Raise vbObjectError + 1000, , "This library was made using a newer " & _
                    "version of the database.  You must upgrade the program before " & _
                    "importing this library."
        Else
            DatabaseVersionOK = True
        End If
    End If
    
ErrExit:
    Set rs = Nothing
    Exit Function

ErrSection:
    RaiseError "cPackagedFile.DatabaseVersionOK", eGDRaiseError_Raise, g.strAppPath
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BuildIncludeFiles
'' Description: Combine the include files into a Zip file
'' Inputs:      None
'' Returns:     Zip File converted into a String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BuildIncludeFiles() As cMemBuffer
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset from the database
    Dim strFile As String               ' File read into a string
    Dim lNumFiles As Long               ' Number of files
    Dim Hdr As FileHeader               ' Header for the file
    
    Dim mb As New cMemBuffer
    Dim lSize As Long
    
    ' Get all of the files that we need to include...
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFiles] " & _
                "WHERE [LibraryID]=" & Str(m.lLibraryID) & " " & _
                "ORDER BY [FileName];", dbOpenDynaset)
                
    lNumFiles = 0&
    mb.PutLong 0 '(for now)
    Do While Not rs.EOF
        strFile = AddSlash(g.strAppPath) & rs!FileName
        lSize = FileLength(strFile)
        If lSize > 0 Then
            With Hdr
                .lFlags = rs!Flags
                .dFileDate = rs!FileDate
                .lFileLen = lSize
                .lHeaderSize = 20 + Len(rs!FileName)
            End With
            mb.PutFromMemory GetAddress(Hdr), Len(Hdr)
            mb.PutStr rs!FileName
            mb.FromFile strFile, True
            lNumFiles = lNumFiles + 1
        End If
        rs.MoveNext
    Loop
    
    If lNumFiles > 0 Then
        mb.PutLong lNumFiles, 0
    End If

ErrExit:
    Set BuildIncludeFiles = mb
    Set rs = Nothing
    Exit Function
    
ErrSection:
    Set rs = Nothing
    RaiseError "cPackagedFile.BuildIncludeFiles", eGDRaiseError_Raise, g.strAppPath

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExtractIncludeFiles
'' Description: Extract the Include Files to their proper place
'' Inputs:      String of the Files
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ExtractIncludeFiles(mbFiles As cMemBuffer)
On Error GoTo ErrSection:

    Dim lNumFiles As Long               ' Number of files
    Dim lIndex As Long                  ' Index into a for loop
    Dim Hdr As FileHeader               ' Header for the file
    Dim bExtract As Boolean             ' Should we extract the file?
    Dim strFileName As String           ' Name and Path of File to Extract
    Dim lOffset As Long
    Dim mbFile As New cMemBuffer
    
    If mbFiles.Length > 0 Then
        ' Extract the number of files from the string
        lNumFiles = mbFiles.GetLong(0)
        lOffset = 4
    
        ' Walk through each file and dump it appropriately
        For lIndex = 1 To lNumFiles
            ' Extract information from the header
            If lOffset >= mbFiles.Length - Len(Hdr) Then Exit For
            CopyMemory Hdr, ByVal mbFiles.MemPtr + lOffset, Len(Hdr)
            strFileName = mbFiles.GetStr(Hdr.lHeaderSize - Len(Hdr), lOffset + Len(Hdr))
            lOffset = lOffset + Hdr.lHeaderSize
            
            ' Extract the file
            mbFile.Clear
            mbFile.PutFromMemory mbFiles.MemPtr + lOffset, Hdr.lFileLen, 0
            lOffset = lOffset + Hdr.lFileLen
            
            ' Handle the file appropriately
            If Len(strFileName) > 0 Then
                bExtract = True
                strFileName = AddSlash(g.strAppPath) & strFileName
                If FileExist(strFileName) Then
                    If Hdr.lFlags And 1 Then
                        If FileDate(strFileName) >= Hdr.dFileDate Then
                            bExtract = False
                        End If
                    End If
                End If
                
                If bExtract Then
                    If FileExist(strFileName) Then
                        ' need to overwrite even if read-only
                        KillFile strFileName, True
                    Else
                        ' make sure folder exists
                        MakeDir FilePath(strFileName), False
                    End If
                    mbFile.ToFile strFileName
                    If Hdr.lFlags And 2 Then SetAttr strFileName, GetAttr(strFileName) Or vbReadOnly
                End If
            End If
        Next lIndex
    End If

ErrExit:
    Set mbFile = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "cPackagedFile.ExtractIncludeFiles", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DLLPath
'' Description: Given the nature of the DLL, return the Path where it is
'' Inputs:      Whether it is BuiltIn, Whether it is a Register DLL
'' Returns:     Path where the DLL is located (or should be located)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DLLPath(ByVal bBuiltin As Boolean, ByVal bRegister As Boolean) As String
On Error GoTo ErrSection:

    If bBuiltin Then
        If bRegister Then
            DLLPath = AddSlash(g.strAppPath) & "..\SharedSelfReg\"
        Else
            DLLPath = AddSlash(g.strAppPath)
        End If
    Else
        DLLPath = AddSlash(g.strAppPath) & "..\LibraryDLLs\"
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPackagedFile.DLLPath", eGDRaiseError_Raise, g.strAppPath
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BreakIntoSections
'' Description: Break the big GLB string into an array of sections
'' Inputs:      GLB string, Array to store sections in
'' Returns:     Number of sections stored
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BreakIntoSections(mb As cMemBuffer, mbSections() As cMemBuffer) As Long
On Error GoTo ErrSection:

    Dim SecHeader As SectionHeader      ' Section header that tells size and type
    Dim lSections As Long               ' Number of sections read in
    Dim lOffset As Long
    Dim i&
    
    ' Intialize the array
    ReDim mbSections(0) As cMemBuffer
    Set mbSections(0) = New cMemBuffer
    
    Do While lOffset <= mb.Length - 16
        ' Get the header information for this section...
        CopyMemory SecHeader, ByVal (mb.MemPtr + lOffset), 16
        lOffset = lOffset + 16
        
        ' Resize the array if necessary...
        If SecHeader.Section > UBound(mbSections) Then
            ReDim Preserve mbSections(SecHeader.Section) As cMemBuffer
            For i = 0 To UBound(mbSections)
                If mbSections(i) Is Nothing Then
                    Set mbSections(i) = New cMemBuffer
                End If
            Next
        End If
        
        ' Copy the section string into the array...
        mbSections(SecHeader.Section).PutFromMemory mb.MemPtr + lOffset, SecHeader.lSize, 0
        lOffset = lOffset + SecHeader.lSize
        lSections = lSections + 1
    Loop
    
    ' Return the number of sections...
    BreakIntoSections = lSections

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPackagedFile.BreakIntoSections", eGDRaiseError_Raise, g.strAppPath
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExtractDLL
'' Description: Extract the DLL from the GLB file (if we can)
'' Inputs:      Sections Array
'' Returns:     0 = Failure, 1 = Set up Auto-Import, 2 = Success
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ExtractDLL(mbDLL As cMemBuffer, ByVal bCanCopy As Boolean) As Byte
On Error GoTo ErrSection:

    Dim strHeader As String             ' Header for this section
    Dim lHeaderLen As Long              ' Length of the header
    Dim lExistingCRC As Long            ' CRC of the existing DLL
    Dim astrHeader As New cGdArray      ' Split up header
    Dim strPath As String               ' Path of the existing DLL
    Dim bCopy As Boolean                ' Go ahead and copy and register DLL
    Dim bBuiltin As Boolean             ' Is this a BuiltIn library?
    Dim bRegister As Boolean            ' Is this a dll that needs registering?
    Dim dFileDate As Double             ' Date/Time of the file
    Dim mbDllFile As New cMemBuffer
    
    ExtractDLL = 0
    astrHeader.Create eGDARRAY_Strings

    If mbDLL.Length > 0 Then
        ' Get the header information...
        lHeaderLen = mbDLL.GetLong
        strHeader = mbDLL.GetStr(lHeaderLen - 4)
        astrHeader.SplitFields strHeader, ","
        
        ' Get the DLL Path and the DLL from memory...
        bBuiltin = CBool(astrHeader(DLLHdrFld(eDLLHeader_BuiltIn)))
        bRegister = CBool(astrHeader(DLLHdrFld(eDLLHeader_Register)))
        strPath = DLLPath(bBuiltin, bRegister)
        mbDllFile.PutFromMemory mbDLL.MemPtr + lHeaderLen, mbDLL.Length - lHeaderLen, 0
        dFileDate = Val(astrHeader(DLLHdrFld(eDLLHeader_FileDate)))
        
        ' TLB 10/17/2006: only allow copying newer-gdBars-version DLL's
        If dFileDate >= 38991 Then
            ' If the DLL currently exists...
            If FileExist(strPath & astrHeader(DLLHdrFld(eDLLHeader_FileName))) Then
                gdCalcFileCRC32 strPath & astrHeader(DLLHdrFld(eDLLHeader_FileName)), lExistingCRC
                
                If lExistingCRC <> CLng(astrHeader(DLLHdrFld(eDLLHeader_CRC))) Then
                    If bCanCopy Then
                        CopyAndRegisterDLL strPath & astrHeader(DLLHdrFld(eDLLHeader_FileName)), mbDllFile, bRegister, dFileDate
                        ExtractDLL = 2
                    Else
                        If Not DirExist(AddSlash(g.strAppPath) & "TempLib\") Then
                            If Not MakeDir(AddSlash(g.strAppPath) & "TempLib\") Then
                                Err.Raise vbObjectError + 1000, , "Could not create temporary directory"
                            End If
                        End If
                        mGenesis.CopyFiles m.strPath, AddSlash(g.strAppPath) & "TempLib\"
                        
                        FileFromString AddSlash(g.strAppPath) & "AutoImport.CFG", _
                            "F," & FileBase(m.strPath) & "." & FileExt(m.strPath), True, True
                        ExtractDLL = 1
                    End If
                End If
            Else
                CopyAndRegisterDLL strPath & astrHeader(DLLHdrFld(eDLLHeader_FileName)), mbDllFile, bRegister, dFileDate
                ExtractDLL = 2
            End If
        End If
    End If

ErrExit:
    Set mbDllFile = Nothing
    Set astrHeader = Nothing
    Exit Function
    
ErrSection:
    Set astrHeader = Nothing
    RaiseError "cPackagedFile.ExtractDLL", eGDRaiseError_Raise, g.strAppPath

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CopyAndRegisterDLL
'' Description: Copy and Register the given DLL
'' Inputs:      FileName and Path of DLL, DLL itself
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CopyAndRegisterDLL(ByVal strFileNameAndPath$, mbDLL As cMemBuffer, ByVal bRegister As Boolean, _
                Optional ByVal dFileDate As Double = -1#)
On Error GoTo ErrSection:

    mbDLL.ToFile strFileNameAndPath
    If bRegister Then
        If gdRegisterFiles(strFileNameAndPath, ByVal 0) = 0 Then
            Err.Raise vbObjectError + 1000, , "Could not register " & strFileNameAndPath
        End If
    End If
    
    ' If a date was passed in, set the file date and time...
    If dFileDate <> -1 Then gdSetFileDate strFileNameAndPath, dFileDate
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPackagedFile.CopyAndRegisterDLL", eGDRaiseError_Raise, g.strAppPath
    
End Sub

Private Function GetField(ByVal rs As Recordset, ByVal lIndex As Long) As String
On Error GoTo ErrSection:

    'Select Case rs.Fields.Item(lIndex).Type
    '    Case dbBoolean
    '        GetField = CStr(NullChk(rs.Fields.Item(lIndex).Value))
    '
    '    Case dbDate
    '        GetField = Str(CDate(NullChk(rs.Fields.Item(lIndex).Value)))
    '
    '        If rs.Fields.Item(lIndex).Value = 0 Then
    '            GetField = "12:00:00 AM"
    '        Else
    '            GetField = Format(rs.Fields.Item(lIndex).Value, "mm/dd/yy HH:MM AM/PM")
    '        End If
    '
    '    Case Else
    '        GetField = Str(NullChk(rs.Fields.Item(lIndex).Value))
    '
    'End Select
    
    ' DAJ 11/1/2004: The above code for handling boolean and date values does not work
    ' for libraries that were exported under European settings.  The ulitimate answer
    ' is to do the following code instead...
    GetField = Str(NullChk(rs.Fields.Item(lIndex).Value))

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPackagedFile.GetField", eGDRaiseError_Raise, g.strAppPath

End Function

Private Function GetDate(ByVal strDate As String) As Double
On Error GoTo ErrSection:

    Dim astrDate As New cGdArray
    Dim astrTime As New cGdArray
    Dim lDate As Long
    Dim lMinute As Long
    
    If strDate = "12:00:00 AM" Then
        GetDate = 0#
    ElseIf InStr(strDate, "/") <> 0 Then
        astrDate.SplitFields Parse(strDate, " ", 1), "/"
        astrTime(0) = Parse(Parse(strDate, " ", 2), ":", 1)
        astrTime(1) = Parse(Parse(strDate, " ", 2), ":", 2)
        astrTime(2) = Parse(strDate, " ", 3)
        
        If Len(Trim(astrTime(0))) = 0 Then
            astrTime(0) = "12"
            astrTime(1) = "00"
            astrTime(2) = "AM"
        End If
        
        If astrTime(0) = "12" Then astrTime(0) = "00"
        
        lDate = Val(astrDate(2)) * 10000 + Val(astrDate(0)) * 100 + Val(astrDate(1))
        If UCase(Left(astrTime(2), 1)) = "P" Then
            lMinute = (Val(astrTime(0)) + 12) * 60 + astrTime(1)
        Else
            lMinute = Val(astrTime(0)) * 60 + astrTime(1)
        End If
        
        GetDate = CDbl(JulFromLong(AddCentury(lDate))) + (CDbl(lMinute) / 1440#)
    Else
        GetDate = Val(strDate)
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPackagedFile.GetDate", eGDRaiseError_Raise, g.strAppPath
    
End Function

' TLB 5/16/2012: to try to minimize time it takes to verify after importing (esp. for Pete)
Private Function UsingNewVerifyMode() As Boolean
On Error Resume Next

#If 1 Then
    'UsingNewVerifyMode = FileExist("C:\NewImportVerify.flg") ' if true, use new method
    UsingNewVerifyMode = Not FileExist("C:\OldImportVerify.flg") ' if true, use new method
#Else
    Static iMode As Integer
    If iMode = 0 Then ' first-time check
        iMode = FileExist("C:\NewImportVerify.flg") ' if true, use new method
        If iMode >= 0 Then iMode = 1 ' otherwise use old method
    End If
    UsingNewVerifyMode = (iMode < 0)
#End If

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ImportStrategyBaskets
'' Description: Import the strategy baskets
'' Inputs:      None
'' Returns:     Strategy Basket ID Map
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ImportStrategyBaskets() As cGdTree
On Error GoTo ErrSection:
    
    Dim BasketIdMap As cGdTree          ' Strategy basket ID map
    Dim rsAllBaskets As Recordset       ' Recordset of all strategy baskets out of the database
    Dim rsBasketItems As Recordset      ' Recordset of strategy basket items in the database
    Dim astrNames As cGdArray           ' Array of strategy basket names already used
    Dim lIndex As Long                  ' Index into the big array
    Dim strOutputLine As String         ' Output line from the big array
    Dim lPos As Long                    ' Position in an array
    Dim bOK As Boolean                  ' Are we OK to continue?
    Dim strMatchName As String          ' Matching name in the recordset
    Dim X As Long                       ' Index into a for loop
    Dim lStrategyBasketID As Long       ' Strategy Basket ID
    
    lIndex = FindSection("tblStrategyBaskets")
    
    Set rsAllBaskets = g.dbNav.OpenRecordset("SELECT * FROM [tblStrategyBaskets] ORDER BY [Name];", dbOpenDynaset)
    Set BasketIdMap = New cGdTree
    Set astrNames = New cGdArray
    astrNames.Create eGDARRAY_Strings
    
    Do Until lIndex > m.astrExportFile.Size - 1
        'Get detail record to post...
        strOutputLine = m.astrExportFile(lIndex)
        strOutputLine = ReplaceChars(strOutputLine, "||", Chr(13) & Chr(10))
        m.b = Split(strOutputLine, vbTab)
        
        'Exit if the line retrieve is the next section header
        If Left(m.b(0), 1) = "[" Then Exit Do
        
        ' Build a sorted array of System astrNames for the next section...
        If Not astrNames.BinarySearch(Str(m.b(1)), lPos) Then
            astrNames.Add Str(m.b(1)), lPos
        End If
        
        With rsAllBaskets
            bOK = True
            
            'look for a matching record (always leave ourselves on the record to edit)
            strMatchName = m.b(1)
            .FindFirst "[Name]='" & strMatchName & "'"
            
            ' If there is no match, we are adding a new record
            If .NoMatch Then
                strMatchName = ""
            
            ' If found a "matching" record
            Else
                'see if in a different library
                If !LibraryID <> m.lLibraryID Then
                    'if not builtin, can't add
                    If Not m.bBuiltin And Not m.bGenesisManaged Then
                        bOK = False
                    Else
                        'else first try to rename existing system
                        .FindFirst "[Name]='_" & strMatchName & "'"
                        If .NoMatch Then
                            .FindFirst "[Name]='" & strMatchName & "'"
                            .Edit
                            !Name = "_" & strMatchName
                            !CheckSum = BuildCheckSum(rsAllBaskets, "tblStrategyBaskets")
                            .Update
                            strMatchName = ""
                        End If
                        .FindFirst "[Name]='" & strMatchName & "'"
                    End If
                End If
            End If
            
            If bOK = False Then
                InfBox "ERROR: Could not import strategy basket '" & m.b(1) & "'", "!", , "Import Error"
                m.aMsg.Add "Couldn't add strategy basket '" & UCase(m.b(1)) & "' because name exists in another library"
            Else
                If strMatchName = "" Then
                    .AddNew
                Else
                    .Edit
                End If
                
                ' Map the old strategy basket ID to the new one...
                BasketIdMap.Add Str(!StrategyBasketID), "OldID=" & Str(m.b(0))
                
                For X = 1 To .Fields.Count - 1
                    
                    'If nbr of fields in tables GREATER than the number
                    'in the text file then the text file is old.  Use the
                    'default value from the table
                    If X <= UBound(m.b) Then
                        .Fields(X).Value = MoveArrayValue(m.b(X), X, rsAllBaskets)
                    Else
                        .Fields(X).Value = MoveArrayValue("", X, rsAllBaskets)
                    End If
                    
                Next X
                
                ' Need to re-encrypt appropriate fields
                EncryptField !Password, NullChk(!Password)
                
                !LibraryID = m.lLibraryID
                !CheckSum = BuildCheckSum(rsAllBaskets, "tblStrategyBaskets")
                
                lStrategyBasketID = !StrategyBasketID
                
                .Update
                
                ' Delete any basket items ( cascade delete should delete parms as well ) for this strategy
                ' basket so that the import can recreate them...
                Set rsBasketItems = g.dbNav.OpenRecordset("SELECT * FROM [tblStrategyBasketItems] WHERE [StrategyBasketID]=" & Str(lStrategyBasketID) & ";", dbOpenDynaset)
                Do While Not rsBasketItems.EOF
                    rsBasketItems.Delete
                    rsBasketItems.MoveNext
                Loop
            End If
            
        End With
                    
        lIndex = lIndex + 1
    Loop
    
    ' Delete any existing strategy baskets in library that are NOT in the packaged
    ' library these are outdated strategy baskets and should be deleted....
    Set rsAllBaskets = g.dbNav.OpenRecordset("SELECT * FROM [tblStrategyBaskets] " & _
        "WHERE [LibraryID]=" & Str(m.lLibraryID) & " " & _
        "ORDER BY [Name];", dbOpenDynaset)
    Do Until rsAllBaskets.EOF
        If Not astrNames.BinarySearch(rsAllBaskets!Name) Then
            rsAllBaskets.Delete
        End If
        
        rsAllBaskets.MoveNext
    Loop
    
    Set ImportStrategyBaskets = BasketIdMap
    
ErrExit:
    Exit Function
    
ErrSection:
    If Err.Number = 3022 Then
        Err.Raise eDupItem, , "A strategy basket: '" & UCase(m.b(0)) & "' already exists.  In order to install this library, rename your version of the strategy basket to something else."
    Else
        RaiseError "cPackagedFile.ImportStrategyBaskets", , g.strAppPath
    End If
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ImportStrategyBasketItems
'' Description: Import the strategy basket items
'' Inputs:      Strategy Basket ID Map
'' Returns:     Strategy Basket Item ID Map
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ImportStrategyBasketItems(ByVal BasketIdMap As cGdTree) As cGdTree
On Error GoTo ErrSection:

    Dim BasketItemIdMap As cGdTree      ' Strategy basket item ID map
    Dim lIndex As Long                  ' Index into the big array
    Dim rsSystems As Recordset          ' Recordset of systems in the database
    Dim rsBaskets As Recordset          ' Recordset of strategy baskets in the database
    Dim rsBasketItems As Recordset      ' Recordset of strategy basket items in the database
    Dim strOuputLine As String          ' Output line from the big array
    Dim lSystemNumber As Long           ' System number for the name in the output line
    Dim lBasketID As Long               ' Strategy basket ID for the item
    Dim X As Long                       ' Index into a for loop
    
    lIndex = FindSection("tblStrategyBasketItems")
    
    Set rsSystems = g.dbNav.OpenRecordset("SELECT * FROM [tblSystems] ORDER BY [SystemName];", dbOpenDynaset)
    Set rsBaskets = g.dbNav.OpenRecordset("SELECT * FROM [tblStrategyBaskets] ORDER BY [Name];", dbOpenDynaset)
    Set rsBasketItems = g.dbNav.OpenRecordset("SELECT * FROM [tblStrategyBasketItems];", dbOpenDynaset)
    Set BasketItemIdMap = New cGdTree
    
    Do Until lIndex > m.astrExportFile.Size - 1
        'Get detail record to post...
        strOuputLine = m.astrExportFile(lIndex)
        strOuputLine = ReplaceChars(strOuputLine, "||", Chr(13) & Chr(10))
        m.b = Split(strOuputLine, vbTab)
        
        'Exit if the line retrieve is the next section header
        If Left(m.b(0), 1) = "[" Then Exit Do
        
        'Find the System Number
        lSystemNumber = 0
        rsSystems.FindFirst "[SystemName]='" & m.b(2) & "'"
        If rsSystems.NoMatch Then
            m.aMsg.Add "Strategy: '" & m.b(2) & "' was not found for strategy basket item '" & m.b(0) & "'"
        Else
            lSystemNumber = rsSystems!SystemNumber
        
            If BasketIdMap.Exists("OldID=" & Str(m.b(1))) = False Then
                m.aMsg.Add "Strategy basket '" & Str(m.b(1)) & "' not found for item '" & m.b(0) & "'"
            Else
                lBasketID = CLng(Val(BasketIdMap("OldID=" & Str(m.b(1)))))
                
                'Search for Rule/Update
                With rsBasketItems
                    .AddNew
                    
                    BasketItemIdMap.Add Str(!StrategyBasketItemID), "OldID=" & Str(m.b(0))
                    !StrategyBasketID = lBasketID
                    !SystemNumber = lSystemNumber
                    
                    For X = 3 To .Fields.Count - 1
                        
                        'If nbr of fields in tables GREATER than the number
                        'in the text file then the text file is old.  Use the
                        'default value from the table
                        If X <= UBound(m.b) Then
                            .Fields(X).Value = MoveArrayValue(m.b(X), X, rsBasketItems)
                        Else
                            .Fields(X).Value = MoveArrayValue("", X, rsBasketItems)
                        End If
                        
                    Next X
                    
                    !CheckSum = BuildCheckSum(rsBasketItems, "tblStrategyBasketItems")
                    .Update
                End With
            End If
        End If
        
        lIndex = lIndex + 1
    Loop
    
    Set ImportStrategyBasketItems = BasketItemIdMap
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPackagedFile.ImportStrategyBasketItems", , g.strAppPath
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ImportStrategyBasketItemParms
'' Description: Import the strategy basket item parms
'' Inputs:      Strategy Basket Item ID Map
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ImportStrategyBasketItemParms(ByVal BasketItemIdMap As cGdTree)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into the big array
    Dim rsBasketItemParms As Recordset  ' Recordset of strategy basket item parms in the database
    Dim strOuputLine As String          ' Output line from the big array
    Dim lBasketItemID As Long           ' Strategy basket item ID for the parm
    Dim X As Long                       ' Index into a for loop
    
    lIndex = FindSection("tblStrategyBasketItemParms")
    
    Set rsBasketItemParms = g.dbNav.OpenRecordset("SELECT * FROM [tblStrategyBasketItemParms];", dbOpenDynaset)
    
    Do Until lIndex > m.astrExportFile.Size - 1
        'Get detail record to post...
        strOuputLine = m.astrExportFile(lIndex)
        strOuputLine = ReplaceChars(strOuputLine, "||", Chr(13) & Chr(10))
        m.b = Split(strOuputLine, vbTab)
        
        'Exit if the line retrieve is the next section header
        If Left(m.b(0), 1) = "[" Then Exit Do
        
        If BasketItemIdMap.Exists("OldID=" & Str(m.b(0))) = False Then
            m.aMsg.Add "Strategy basket item " & Str(m.b(0)) & " not found for parm"
        Else
            lBasketItemID = CLng(Val(BasketItemIdMap("OldID=" & Str(m.b(0)))))
            
            'Search for Rule/Update
            With rsBasketItemParms
                .AddNew
                
                !StrategyBasketItemID = lBasketItemID
                
                For X = 1 To .Fields.Count - 1
                    
                    'If nbr of fields in tables GREATER than the number
                    'in the text file then the text file is old.  Use the
                    'default value from the table
                    If X <= UBound(m.b) Then
                        .Fields(X).Value = MoveArrayValue(m.b(X), X, rsBasketItemParms)
                    Else
                        .Fields(X).Value = MoveArrayValue("", X, rsBasketItemParms)
                    End If
                    
                Next X
                
                .Update
            End With
        End If
        
        lIndex = lIndex + 1
    Loop
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPackagedFile.ImportStrategyBasketItemParms", , g.strAppPath
    
End Sub
