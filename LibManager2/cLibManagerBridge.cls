VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLibManagerBridge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cLibManagerBridge.cls
'' Description: Allow the calling application to pass in stuff to the LibMgr
''
'' Author:      Genesis Financial Data Services
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 05/01/2013   DAJ         Shadow Trading
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Public Enum eCalledFrom
    SystemNavigator = 1
    PortfolioNavigator = 2
End Enum

Private Type mPrivate
    bImportOK As Boolean
    bSaved As Boolean
    bExportOK As Boolean
End Type
Private m As mPrivate

Public Property Get ImportOK() As Boolean
    ImportOK = m.bImportOK
End Property
Public Property Get Saved() As Boolean
    Saved = m.bSaved
End Property
Public Property Get ExportOK() As Boolean
    ExportOK = m.bExportOK
End Property

Public Property Let Help(HelpObj As Object)
    Set g.Help = HelpObj
End Property
Public Property Get Help() As Object
    Set Help = g.Help
End Property

Public Property Let OwnerForm(ByVal frmOwner As Object)
    Set g.frmOwner = Nothing
    On Error Resume Next
    Set g.frmOwner = frmOwner
End Property

Public Property Get Reload() As Boolean
    Reload = g.bReload
End Property

Public Property Get ShowShadow() As Boolean
    ShowShadow = g.bShowShadow
End Property
Public Property Let ShowShadow(ByVal bShowShadow As Boolean)
    g.bShowShadow = bShowShadow
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Changed
'' Description: Return if the library has changed or not
'' Inputs:      None
'' Returns:     True if Library Changed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Get Changed() As Boolean

    Changed = g.bChanged

End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    dbNavRef
'' Description: Set the global database reference to the one passed in
'' Inputs:      Database reference
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Let dbNavRef(pData As Database)
On Error GoTo ErrSection:

    Set g.dbNav = pData
    Set g.CommonBridge = CreateObject(g.strCommonDLL & "cCommonBridge")
    g.CommonBridge.dbNavRef = pData
    g.bChanged = False
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cLibManagerBridge.dbNavRef", eGDRaiseError_Raise, g.strAppPath
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    dbNavPassword
'' Description: Pass in the password for the database
'' Inputs:      Database password
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Let dbNavPassword(ByVal strPassword As String)
    
    g.strPassword = strPassword

End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CalledFrom
'' Description: Set the calling application (SysNav or PortNav)
'' Inputs:      Calling application
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Let CalledFrom(ByVal CalledFrom As eCalledFrom)
On Error GoTo ErrSection:

    g.CalledFrom = CalledFrom
    If g.CalledFrom = SystemNavigator Then
        g.strMdbName = "Libraries.MDB"
        g.strCommonDLL = "NavCommon2."
        g.strTradeSenseOCX = "NavTradeSenseV3."
    Else
        g.strMdbName = "PortNav.MDB"
        g.strCommonDLL = "PNCommon."
        g.strTradeSenseOCX = "PNTradeSense."
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cLibManagerBridge.CalledFrom", eGDRaiseError_Raise, g.strAppPath

End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AppPath
'' Description: Pass in the application path for the calling application
'' Inputs:      Application Path
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Let AppPath(ByVal strAppPath As String)
    
    g.strAppPath = strAppPath
    g.strIniFile = AddSlash(strAppPath) & "ChartNavigator.INI"

End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ImageList
'' Description: Return the global image list
'' Inputs:      None
'' Returns:     Global image list
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Get ImageList() As ListImages
    
    Set ImageList = g.ListImages

End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ImageList
'' Description: Set the global image list
'' Inputs:      Image List
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Let ImageList(ImgList As ListImages)
    
    Set g.ListImages = ImgList

End Property

Property Let HonestDate(ByVal lHonestDate As Long)
    g.lHonestDate = lHonestDate
End Property
Property Let CustomerID(ByVal lLCD As Long)
    g.lLCD = lLCD
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NewLibrary
'' Description: Start the New Library wizard
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub NewLibrary()
On Error GoTo ErrSection:

    m.bSaved = frmLibrary.ShowMe

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cLibManagerBridge.NewLibrary", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CreateNewLibrary
'' Description: Start the New Library wizard
'' Inputs:      None
'' Returns:     Library ID if Saved, Zero otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CreateNewLibrary() As Long
On Error GoTo ErrSection:

    m.bSaved = frmLibrary.ShowMe
    If m.bSaved Then
        CreateNewLibrary = frmLibrary.LibraryID
    Else
        CreateNewLibrary = 0&
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cLibManagerBridge.CreateNewLibrary", eGDRaiseError_Raise, g.strAppPath

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EditLibrary
'' Description: Allow the user to edit a library
'' Inputs:      ID of the Library to Edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub EditLibrary(ByVal lLibraryID As Long)
On Error GoTo ErrSection:
    
    m.bSaved = frmLibrary.ShowMe(lLibraryID)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cLibManagerBridge.EditLibrary", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ViewLibrary
'' Description: Allow the user to view information about a library
'' Inputs:      ID of the Library to View
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ViewLibrary(ByVal lLibraryID As Long)
On Error GoTo ErrSection:
    
    frmLibraryViewer.ShowMe lLibraryID

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cLibManagerBridge.ViewLibrary", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowPackager
'' Description: Allow the user to Export a Library
'' Inputs:      ID of the Library to Export
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowPackager(ByVal lLibraryID As Long)
On Error GoTo ErrSection:

    m.bExportOK = frmLibraryPackager.ShowMe(lLibraryID)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cLibManagerBridge.ShowPackager", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowImporter
'' Description: Allow the user to Import a Library
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowImporter()
On Error GoTo ErrSection:
    
    m.bImportOK = frmLibraryImport.ShowMe

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cLibManagerBridge.ShowImporter", eGDRaiseError_Raise, g.strAppPath

End Sub

Public Function ExportAuto(ByVal lLibraryID As Long) As Boolean
On Error GoTo ErrSection:

    ExportAuto = frmLibraryPackager.ShowMe(lLibraryID, True)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cLibManagerBridge.ExportAuto", eGDRaiseError_Raise, g.strAppPath
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CreateLibraryFromItems
'' Description: Create a library from the name and items passed in
'' Inputs:      Name of Library, Items to Add, New Library ID
'' Returns:     True if Saved, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CreateLibraryFromItems(ByVal strName As String, ByVal strItems As String, lLibraryID As Long) As Boolean
On Error GoTo ErrSection:

    Dim strLibraryName As String        ' Name of the library to create
    Dim strAuthor As String             ' Author of the library to create
    Dim strDescription As String        ' Description of the library to create
    Dim strPassword As String           ' Password for the library to create
    Dim bFull As Boolean                ' Permissions for the library
    Dim bReturn As Boolean              ' Return value
    
    strLibraryName = strName
    bFull = True
    
    bReturn = frmLibraryInfo.ShowMe(strLibraryName, strAuthor, strDescription, bFull, strPassword)
    If bReturn = True Then
        bReturn = frmLibrary.ShowMeCreate(strLibraryName, strAuthor, strDescription, bFull, strPassword, strItems, lLibraryID)
    End If
    
    CreateLibraryFromItems = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cLibManagerBridge.CreateLibraryFromItems", eGDRaiseError_Raise, g.strAppPath
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initiliaze class members when object is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()

    g.bShowShadow = False

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Upon class termination cleanup after ourselves
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
    
    Set g.CommonBridge = Nothing
    Set g.frmOwner = Nothing
    Set g.Help = Nothing
    Set g.ListImages = Nothing
    Set g.dbNav = Nothing
    Set g.WrkJet = Nothing

End Sub
