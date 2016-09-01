VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFileItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cFileItem.cls
'' Description: Class to handle a file item
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
''
'' 02/19/2009   DAJ         Created
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    strGroupName As String              ' Name of the group for display
    strZipName As String                ' Name of the Zipped up file
    strUnzippedPath As String           ' Path for the unzipped files
    strFileMasks As String              ' File masks to be zipped up
    lDelFirst As Long                   ' Delete files before restoring?
    nOverwriteMode As eGD_OverwriteMode ' Mode for overwriting files
End Type
Private m As mPrivate

Public Property Get GroupName() As String
    GroupName = m.strGroupName
End Property
Public Property Let GroupName(ByVal strGroupName As String)
    m.strGroupName = strGroupName
End Property

Public Property Get ZipName() As String
    ZipName = m.strZipName
End Property
Public Property Let ZipName(ByVal strZipName As String)
    m.strZipName = strZipName
End Property

Public Property Get UnzippedPath() As String
    UnzippedPath = m.strUnzippedPath
End Property
Public Property Let UnzippedPath(ByVal strUnzippedPath As String)
    m.strUnzippedPath = strUnzippedPath
End Property

Public Property Get FileMasks() As String
    FileMasks = m.strFileMasks
End Property
Public Property Let FileMasks(ByVal strFileMasks As String)
    m.strFileMasks = strFileMasks
End Property

Public Property Get DelFirst() As Long
    DelFirst = m.lDelFirst
End Property
Public Property Let DelFirst(ByVal lDelFirst As Long)
    m.lDelFirst = lDelFirst
End Property

Public Property Get OverwriteMode() As eGD_OverwriteMode
    OverwriteMode = m.nOverwriteMode
End Property
Public Property Let OverwriteMode(ByVal nOverwriteMode As eGD_OverwriteMode)
    m.nOverwriteMode = nOverwriteMode
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ZipUpFiles
'' Description: Zip up the files according to the file item information
'' Inputs:      Path
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ZipUpFiles(ByVal strPath As String)
On Error GoTo ErrSection:

    Dim astrFileMasks As cGdArray       ' Array of file masks
    Dim lIndex As Long                  ' Index into a for loop
    
    Set astrFileMasks = New cGdArray
    astrFileMasks.SplitFields m.strFileMasks, ","

    For lIndex = 0 To astrFileMasks.Size - 1
        If lIndex = 0 Then
            KillFile AddSlash(strPath) & m.strZipName
        End If
        
        ZipExecute "Z", AddSlash(strPath) & m.strZipName, AddSlash(App.Path), astrFileMasks(lIndex)
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFileItems.ZipUpFiles"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExtractFiles
'' Description: Extract the files according to the file item information
'' Inputs:      Path
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ExtractFiles(ByVal strPath As String, Optional ByVal bDeleteZipFile As Boolean = True)
On Error GoTo ErrSection:

    If m.lDelFirst <> 0 Then
        DeleteFiles
    End If
    
    ZipExecute "U", AddSlash(strPath) & m.strZipName, m.strUnzippedPath
    
    If bDeleteZipFile Then
        KillFile AddSlash(strPath) & m.strZipName
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFileItem.ExtractFiles"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LogMessage
'' Description: Create a message to be sent to the log file
'' Inputs:      Backup?
'' Returns:     Log Message
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LogMessage(ByVal bBackup As Boolean) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function

    If bBackup = True Then
        strReturn = "Zipping up " & m.strFileMasks & " from " & m.strUnzippedPath & " in " & m.strZipName
    Else
        strReturn = "Unzipping files from " & m.strZipName & " into folder " & m.strUnzippedPath & " (Delete First = " & Str(m.lDelFirst) & ", Overwrite Mode = " & m.nOverwriteMode & ")"
    End If
    
    LogMessage = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cFileItem.LogMessage"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeleteFiles
'' Description: Delete files before we do a restore
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub DeleteFiles()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim astrFileMasks As New cGdArray   ' Array of file masks
    
    astrFileMasks.SplitFields m.strFileMasks, ","
    
    For lIndex = 0 To astrFileMasks.Size - 1
        KillFile AddSlash(App.Path) & astrFileMasks(lIndex), True
    Next lIndex
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFileItem.DeleteFiles"
    
End Sub
