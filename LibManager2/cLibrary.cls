VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cLibrary.cls
'' Description: Handle Library Management
''
'' Author:      Genesis Financial Data Services
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 04/03/2013   DAJ         Move Strategy Baskets into database
'' 05/01/2013   DAJ         Shadow Trading
'' 05/24/2013   DAJ         Don't allow library to be non-Restricted if Shadow
'' 05/28/2013   DAJ         Changed message on password error
'' 04/01/2014   DAJ         If user saves from the UI, clear the expiration date
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Option Compare Text

Private Type mPrivate
    lLibraryID As Long
    vLibraryName As Variant
    vLibraryDesc As Variant
    vDLLName As Variant
    vSecurityLevel As Variant
    vCannotDelete As Variant
    vPassword As Variant
    vVersion As Variant
    vLastModified As Variant
    vAuthor As Variant
    vEMail As Variant
    vPhone As Variant
    vWebSite As Variant
    vRtfFileName As Variant
    PackagedFile As cPackagedFile
    vLibraryType As Variant
    bBuiltin As Boolean
    lDLLSecurityCode As Long
    strLastExported As String
    bIsGuru As Boolean
    strOwners As String
    
    dExpDate As Double
    strCustomerID As String
    strRequiredMod As String
    bIgnore As Boolean

    vsStatusBar As VSFlexGrid
    strPath As String

    vsItems As VSFlexGrid
    bErrNbr As Byte
End Type
Private m As mPrivate

Private Const kItemTypeCatColumn = 2
Private Const kClassName = "Class:cLibrary "

Property Get ErrNbr() As Byte
    ErrNbr = m.bErrNbr
End Property
Property Get PackagedFile() As cPackagedFile
    Set PackagedFile = m.PackagedFile
End Property
Property Get LibraryID() As Variant
    LibraryID = m.lLibraryID
End Property
Property Get LibraryName() As Variant
    LibraryName = m.vLibraryName
End Property
Property Get LibraryDesc() As Variant
    LibraryDesc = m.vLibraryDesc
End Property
Property Get DLLName() As Variant
    DLLName = m.vDLLName
End Property
Property Get SecurityLevel() As Variant
    SecurityLevel = m.vSecurityLevel
End Property
Property Get CannotDelete() As Variant
    CannotDelete = m.vCannotDelete
End Property
Property Get Password() As Variant
    Password = m.vPassword
End Property
Property Get Version() As Variant
    Version = m.vVersion
End Property
Property Get LastModified() As Variant
    LastModified = m.vLastModified
End Property
Property Get Author() As Variant
    Author = m.vAuthor
End Property
Property Get EMail() As Variant
    EMail = m.vEMail
End Property
Property Get Phone() As Variant
    Phone = m.vPhone
End Property
Property Get WebSite() As Variant
    WebSite = m.vWebSite
End Property
Property Get RtfFileName() As Variant
    RtfFileName = m.vRtfFileName
End Property
Property Get LibraryType() As Variant
    LibraryType = m.vLibraryType
End Property
Property Get BuiltIn() As Boolean
    BuiltIn = m.bBuiltin
End Property
Property Get DLLSecurityCode() As Long
    DLLSecurityCode = m.lDLLSecurityCode
End Property
Property Get ExpirationDate() As Double
    ExpirationDate = m.dExpDate
End Property
Property Get CustomerID() As String
    CustomerID = m.strCustomerID
End Property
Property Get RequiredMod() As String
    RequiredMod = m.strRequiredMod
End Property
Property Get Ignore() As Boolean
    Ignore = m.bIgnore
End Property
Property Get LastExported() As String
    LastExported = m.strLastExported
End Property

Property Let Items(pData As Object)
    Set m.vsItems = pData
End Property
Property Let LibraryID(ByVal pData As Variant)
    m.lLibraryID = pData
End Property
Property Let LibraryName(pData As Variant)
    m.vLibraryName = pData
End Property
Property Let LibraryDesc(pData As Variant)
    m.vLibraryDesc = pData
End Property
Property Let DLLName(pData As Variant)
    m.vDLLName = pData
End Property
Property Let SecurityLevel(pData As Variant)
    m.vSecurityLevel = pData
End Property
Property Let CannotDelete(pData As Variant)
    m.vCannotDelete = pData
End Property
Property Let Password(pData As Variant)
    m.vPassword = pData
End Property
Property Let Version(pData As Variant)
    m.vVersion = pData
End Property
Property Let LastModified(pData As Variant)
    m.vLastModified = pData
End Property
Property Let Author(pData As Variant)
    m.vAuthor = pData
End Property
Property Let EMail(pData As Variant)
    m.vEMail = pData
End Property
Property Let Phone(pData As Variant)
    m.vPhone = pData
End Property
Property Let WebSite(pData As Variant)
    m.vWebSite = pData
End Property
Property Let RtfFileName(pData As Variant)
    m.vRtfFileName = pData
End Property
Property Let LibraryType(pData As Variant)
    m.vLibraryType = pData
End Property
Property Let BuiltIn(pData As Boolean)
    m.bBuiltin = pData
End Property
Property Let DLLSecurityCode(pData As Long)
    m.lDLLSecurityCode = pData
End Property
Property Let StatusBar(pData As Object)
    Set m.vsStatusBar = pData
    m.PackagedFile.StatusBar = pData
End Property
Property Let Path(pData As String)
    m.strPath = pData
    m.PackagedFile.Path = pData
End Property
Property Let StatusMsg(pData As Object)
    m.PackagedFile.StatusMsg = pData
End Property
Property Let ExpirationDate(ByVal pData As Double)
    m.dExpDate = pData
End Property
Property Let CustomerID(ByVal pData As String)
    m.strCustomerID = pData
End Property
Property Let RequiredMod(ByVal pData As String)
    m.strRequiredMod = FixRequiredMod(pData)
End Property
Property Let Ignore(ByVal pData As Boolean)
    m.bIgnore = pData
End Property
Property Let LastExported(ByVal strLastExported As String)
    m.strLastExported = strLastExported
End Property

Public Property Get IsGuru() As Boolean
    IsGuru = m.bIsGuru
End Property
Public Property Let IsGuru(ByVal bIsGuru As Boolean)
    m.bIsGuru = bIsGuru
End Property

Public Property Get Owners() As String
    Owners = m.strOwners
End Property
Public Property Let Owners(ByVal strOwners As String)
    m.strOwners = strOwners
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Package
'' Description: Export the library to a GLB file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Package()
On Error GoTo ErrSection:

    SaveLastExportVersion
    
    With m.PackagedFile
        .LibraryID = m.lLibraryID
        .LibraryName = m.vLibraryName
        .LibraryType = m.vLibraryType
        .BuiltIn = m.bBuiltin
        .StatusBar = m.vsStatusBar
        .Path = m.strPath
        .DLLName = m.vDLLName
        .CustomerID = m.strCustomerID
        .ExpirationDate = m.dExpDate
                
        .Build
    End With
    Set m.PackagedFile = Nothing
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cLibrary.Package", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Import
'' Description: Import a Library from a GLB file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Import()
On Error GoTo ErrSection:

    With m.PackagedFile
        .StatusBar = m.vsStatusBar
        .Path = m.strPath
        .Import
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cLibrary.Import", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Validate
'' Description: Validate the information entered about the library
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Validate()
On Error GoTo ErrSection:
    
    Dim rs As Recordset                 ' Recordset into the database
    Dim X As Integer                    ' Index into a for loop
    Dim curChar As String               ' Current character in a string
    
    'Library name must be 50 characters or less
    m.bErrNbr = 1
    If Len(m.vLibraryName) = 0 Then
        Err.Raise vbObjectError + 1000, , "Library name must be entered."
    End If
    If Len(m.vLibraryName) > 50 Then
        Err.Raise vbObjectError + 1000, , "Library name is too long (it must be 50 characters or less)."
    End If
    
    'Library name cannot contain special characters (this name is used
    'when exporting to a text file)
    For X = 1 To Len(m.vLibraryName)
        curChar = Mid(m.vLibraryName, X, 1)
        If (curChar >= "a" And curChar <= "z") Or _
           (curChar >= "0" And curChar <= "9") Or _
           curChar = " " Then
        Else
            Err.Raise vbObjectError + 1000, , "Library name must contain letters and numbers only."
        End If
    Next X

    'Version required
    m.bErrNbr = 3
    If Len(m.vVersion) = 0 Then
        Err.Raise vbObjectError + 1000, , "Version number must be entered."
    End If
    If Len(m.vVersion) > 50 Then
        Err.Raise vbObjectError + 1000, , "Version number is too long (it must be 50 characters or less)."
    End If
    
    'Library description required
    m.bErrNbr = 2
    'If Len(m.vLibraryDesc) = 0 Then
    '    Err.Raise vbObjectError + 1000, , "Library description must be entered."
    'End If
    If Len(m.vLibraryDesc) > 255 Then
        Err.Raise vbObjectError + 1000, , "Library description is too long (it must be 255 characters or less)."
    End If
    
    'Author required
    m.bErrNbr = 4
    If Len(m.vAuthor) = 0 Then
        Err.Raise vbObjectError + 1000, , "Author name must be entered."
    End If
    If Len(m.vAuthor) > 50 Then
        Err.Raise vbObjectError + 1000, , "Author name is too long (it must be 50 characters or less)."
    End If
    
    'If security level is greater than 0 then a password is required
    m.bErrNbr = 5
    If m.vSecurityLevel > 0 Then
        If IsNull(m.vPassword) Or m.vPassword = "" Then
            Err.Raise vbObjectError + 1000, , "RESTRICTED Libraries must have a password"
        End If
        If Len(m.vPassword) > 0 And (Len(m.vPassword) < 5 Or Len(m.vPassword) > 10) Then
            Err.Raise vbObjectError + 1000, , "Passwords must be 5 to 10 characters in length."
        End If
        
    Else
        'Set security level if null
        If Len(m.vSecurityLevel) = 0 Then
            m.vSecurityLevel = 0
        End If
        m.vPassword = ""
    End If
    
    'At least one item must be selected
    'If m.vsItems.Rows = 1 Then
    '    Err.Raise vbObjectError + 1000, , "At least one item must be added to your library."
    'End If
    
    'DLL name required if one or more builtin functions used.
    m.bErrNbr = 7
    If BuiltinFunctionsExist Then
        If Len(m.vDLLName) = 0 Then
            Err.Raise vbObjectError + 1000, , "A DLL name must be entered for the selected builtin functions."
        End If
        If Len(m.vDLLName) > 50 Then
            Err.Raise vbObjectError + 1000, , "DLL name is too long (it must be 50 characters or less)."
        End If
    End If
    
    'RTF file path
    m.bErrNbr = 8
    If Len(m.vRtfFileName) > 0 And Len(m.vRtfFileName) > 255 Then
        Err.Raise vbObjectError + 1000, , "File name is too long (it must be 255 characters or less)."
    End If
    
    'E-Mail name
    m.bErrNbr = 9
    If Len(m.vEMail) > 0 And Len(m.vEMail) > 255 Then
        Err.Raise vbObjectError + 1000, , "E-Mail name is too long (it must be 255 characters or less)."
    End If
    
    'Website name
    m.bErrNbr = 10
    If Len(m.vWebSite) > 0 And Len(m.vWebSite) > 255 Then
        Err.Raise vbObjectError + 1000, , "WebSite name is too long (it must be 255 characters or less)."
    End If
    
    ' Non-Zero security code with FULL Permission set on the Library
    m.bErrNbr = 11
    If m.lDLLSecurityCode > 0 And m.vSecurityLevel = 0 Then
        Err.Raise vbObjectError + 1000, , "Non-Zero Security Code requires Restricted Library Access"
    End If
    
    ' Shadow turned on, but not a restricted library...
    m.bErrNbr = 12
    If (m.bIsGuru = True) And (m.vSecurityLevel = 0) Then
        Err.Raise vbObjectError + 1000, , "A library with Shadow Trading items must be a Restricted library"
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cLibrary.Validate", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BuiltinFunctionsExist
'' Description: Determine whether or not BuiltIn functions exist in the library
'' Inputs:      None
'' Returns:     True if BuiltIn Functions exist, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BuiltinFunctionsExist() As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Integer               ' Index into a for loop
    
    With m.vsItems
        For lIndex = .FixedRows To .Rows - 1
            If .TextMatrix(lIndex, kItemTypeCatColumn) = "Compiled" Or _
               .TextMatrix(lIndex, kItemTypeCatColumn) = "Compiled Action" Or _
               .TextMatrix(lIndex, kItemTypeCatColumn) = "Internal" Then
                BuiltinFunctionsExist = True
                Exit For
            End If
        Next lIndex
    End With
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cLibrary.BuiltinFunctionsExist", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Load
'' Description: Load up the Library from the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Load()
On Error GoTo ErrSection:
    
    Dim rs As Recordset                 ' Recordset into the database
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblLibrarys] " & _
            "WHERE [LibraryID]=" & m.lLibraryID & ";", dbOpenSnapshot)
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        If rs!CheckSum = BuildCheckSum(rs, "tblLibrarys") Then
            m.lLibraryID = rs!LibraryID
            m.vLibraryName = rs!LibraryName
            m.vLibraryDesc = rs!LibraryDesc
            m.vDLLName = NullChk(rs!Path)
            m.vVersion = rs!Version
            m.vLastModified = Now
            m.vAuthor = rs!Author
            m.vRtfFileName = NullChk(rs!RtfFileName)
            m.vEMail = NullChk(rs!EMail)
            m.vPhone = NullChk(rs!Phone)
            m.vWebSite = NullChk(rs!WebSite)
            m.vSecurityLevel = NullChk(rs!SecurityLevel, 0)
            m.vCannotDelete = rs!CannotDelete
            m.vPassword = DecryptField(rs!Password)
            m.vLibraryType = NullChk(rs!LibraryType, 0)
            m.bBuiltin = rs!BuiltIn
            m.lDLLSecurityCode = NullChk(rs!DLLSecurityCode, 0)
            m.strRequiredMod = FixRequiredMod(NullChk(rs!RequiredMod))
            m.bIgnore = rs!Ignore
            m.strLastExported = NullChk(rs!LastExportedVersion, rs!Version)
            If m.strLastExported = "" Then m.strLastExported = CStr(m.vVersion)
            m.bIsGuru = rs!IsGuru
            m.strOwners = rs!Owners
        Else
            Err.Raise vbObjectError + 1000, , "Library could not be loaded because it is no longer valid"
        End If
    End If
    
ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "cLibrary.Load", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Save the Library information to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Save()
On Error GoTo ErrSection:
    
    SaveToDatabase False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cLibrary.Save", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveFromUI
'' Description: Save the Library information to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveFromUI()
On Error GoTo ErrSection:
    
    SaveToDatabase True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cLibrary.SaveFromUI", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Delete
'' Description: Delete a Libary from the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Delete()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database

    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblLibrarys] " & _
            "WHERE [LibraryID]=" & Str(m.lLibraryID) & ";", dbOpenDynaset)
    With rs
        If rs.EOF Then
            Err.Raise vbObjectError + 1000, , "Library was not found. Delete cancelled."
        Else
            rs.Delete
        End If
    End With
    
ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "cLibrary.Delete", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveToDatabase
'' Description: Save the Library information to the database
'' Inputs:      From UI?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveToDatabase(ByVal bFromUI As Boolean)
On Error GoTo ErrSection:
    
    Dim rs As Recordset                 ' Recordset into the database
    Dim rs2 As Recordset                ' Recordset into the database
    
    'Update the table first...
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblLibrarys] " & _
        "WHERE [LibraryID]=" & Str(m.lLibraryID) & ";", dbOpenDynaset)
    With rs
        If .EOF Then
            .AddNew
            Set rs2 = g.dbNav.OpenRecordset("SELECT * FROM [tblLibrarys] " & _
                    "WHERE [LibraryName]='" & Str(m.vLibraryName) & "';", dbOpenDynaset)
            If Not rs2.EOF Then
                Err.Raise vbObjectError + 1000, , "Library: " & Str(m.vLibraryName) & " already exists."
            End If
            
            'Get highest libraryID and add 1
            Set rs2 = g.dbNav.OpenRecordset("SELECT * FROM [tblLibrarys] " & _
                    " ORDER BY [LibraryID] DESC;", dbOpenSnapshot)
            If rs2.EOF Then
                m.lLibraryID = 1
            Else
                m.lLibraryID = rs2!LibraryID + 1
            End If
            !LibraryID = m.lLibraryID
        Else
            .Edit
        End If
            
        !LibraryName = m.vLibraryName
        !LibraryDesc = NullChk(m.vLibraryDesc)
        !Path = NullChk(m.vDLLName)
        !Version = m.vVersion
        !LastModified = Now
        !Author = m.vAuthor
        !SecurityLevel = NullChk(m.vSecurityLevel, 0)
        !CannotDelete = m.vCannotDelete
        EncryptField !Password, NullChk(m.vPassword)
        !RtfFileName = NullChk(m.vRtfFileName)
        !EMail = NullChk(m.vEMail)
        !Phone = NullChk(m.vPhone)
        !WebSite = NullChk(m.vWebSite)
        !LibraryType = NullChk(m.vLibraryType, 0)
        !BuiltIn = m.bBuiltin
        !DLLSecurityCode = m.lDLLSecurityCode
        !RequiredMod = FixRequiredMod(m.strRequiredMod)
        !Ignore = m.bIgnore
        !LastExportedVersion = m.strLastExported
        !IsGuru = m.bIsGuru
        !Owners = m.strOwners
        
        ' 04/01/2014 DAJ: Tim has decided that if the user had access to save the library
        ' from the UI, then clear the expiration date on the library record...
        If bFromUI Then
            !Expiration = 0
        End If
        
        !CheckSum = BuildCheckSum(rs, "tblLibrarys")
        
        .Update
    End With
    
ErrExit:
    Set rs = Nothing
    Set rs2 = Nothing
    Exit Sub
    
ErrSection:
    RaiseError "cLibrary.SaveToDatabase", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveLastExportVersion
'' Description: Save the version of the library at last export
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveLastExportVersion()
On Error GoTo ErrSection:

    Dim rs As Recordset
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblLibrarys] " & _
            "WHERE [LibraryID]=" & Str(m.lLibraryID) & ";", dbOpenDynaset)
    If Not (rs.EOF And rs.BOF) Then
        rs.Edit
        rs!LastExportedVersion = m.strLastExported
        rs!CheckSum = BuildCheckSum(rs, "tblLibrarys")
        rs.Update
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cLibrary.SaveLastExportVersion", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize members when the object is being initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.PackagedFile = New cPackagedFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cLibrary.Class_Initialize", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up when the object is being terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.PackagedFile = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cLibrary.Class_Terminate", eGDRaiseError_Raise, g.strAppPath
    
End Sub
