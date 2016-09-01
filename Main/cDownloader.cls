VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDownloader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cDownloader.cls
'' Description: Class to handle an FTP file download process
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 07/28/2014   DAJ         Added Server & Local Path & Filename for CRC files
'' 08/06/2014   DAJ         Added Note property
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kDownloaderProcess = "Downloader\Downloader.exe"

Public Enum eGDDownloadStatus
    eGDDownloadStatus_Nothing = 0
    eGDDownloadStatus_Downloading = 1
    eGDDownloadStatus_Done = 2
    eGDDownloadStatus_Aborted = 3
    eGDDownloadStatus_Error = 4
End Enum

Private Type mPrivate
    strCaption As String                ' Window caption to use for the downloader process
    strNote As String                   ' Note to display in the downloader window
    
    strUserName As String               ' User name to use to login to the server
    strPassword As String               ' Password to use to login to the server
    strIP As String                     ' IP address of the server
    iPort As Integer                    ' Port to use to login to the server
    strDoneFile As String               ' Path and name of the file to look for when process is done
    
    Files As cGdTree                    ' Collection of information about files to download
    
    bDownloading As Boolean             ' Are we currently downloading the files?
End Type
Private m As mPrivate

Public Property Get Caption() As String
    Caption = m.strCaption
End Property
Public Property Let Caption(ByVal strCaption As String)
    m.strCaption = strCaption
End Property

Public Property Get Note() As String
    Note = m.strNote
End Property
Public Property Let Note(ByVal strNote As String)
    m.strNote = strNote
End Property

Public Property Get UserName() As String
    UserName = m.strUserName
End Property
Public Property Let UserName(ByVal strUserName As String)
    m.strUserName = strUserName
End Property

Public Property Get Password() As String
    Password = m.strPassword
End Property
Public Property Let Password(ByVal strPassword As String)
    m.strPassword = strPassword
End Property

Public Property Get IP() As String
    IP = m.strIP
End Property
Public Property Let IP(ByVal strIP As String)
    m.strIP = strIP
End Property

Public Property Get Port() As Integer
    Port = m.iPort
End Property
Public Property Let Port(ByVal iPort As Integer)
    m.iPort = iPort
End Property

Public Property Get NumberOfFiles() As Integer
    NumberOfFiles = m.Files.Count
End Property

Public Property Get DoneFile() As String
    DoneFile = m.strDoneFile
End Property
Public Property Let DoneFile(ByVal strDoneFile As String)
    m.strDoneFile = strDoneFile
End Property

Public Property Get Files() As cGdTree
    Set Files = m.Files
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Download
'' Description: Call the stand-alone process to download the files from the server
'' Inputs:      Name of the file of information
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Download(ByVal strFileName As String) As Boolean
On Error GoTo ErrSection:
    
    Dim bReturn As Boolean              ' Return value for the function
    Dim strExe$, strTemp$, s$
    
    bReturn = False
    If DownloaderIsRunning Then Exit Function
    
    KillFile m.strDoneFile
    ToFile strFileName
    
    ' in case .EXE was deleted by anti-virus, copy from saved file
    ' (which gives chance for user to run again after adding Genesis folder to exclusions)
    strExe = AddSlash(App.Path) & kDownloaderProcess
    strTemp = ReplaceFileExt(strExe, "BAK")
    If FileExist(strExe) Then
        ' copy to backup
        FileCopy strExe, strTemp, True
    ElseIf FileExist(strTemp) Then
        ' could have been deleted by anti-virus, so copy from backup
        s = "The Downloader.EXE may have been inadvertantly deleted by your anti-virus program!  Before continuing, please add the 'Genesis' folder to your anti-virus Exclusions list.|"
        InfBox s, "!", , "Data Install"
        FileCopy strTemp, strExe, True
    Else
        ' program does not exist!
        Exit Function
    End If
    
    ' run the downloader
    If RunProcess(strExe, Chr(34) & strFileName & Chr(34), False, vbNormalNoFocus) Then
        m.bDownloading = True
        bReturn = True
    End If

    Download = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cDownloader.Download"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Status
'' Description: Determine the status of the file download
'' Inputs:      Error ( out )
'' Returns:     Status
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Status(Optional strError As String) As eGDDownloadStatus
On Error GoTo ErrSection:

    Dim nReturn As eGDDownloadStatus    ' Return value for the function
    Dim astrDoneFile As cGdArray        ' Done file read into an array
    
    If m.bDownloading = False Then
        nReturn = eGDDownloadStatus_Nothing
    ElseIf FileExist(DoneFile) Then
        Set astrDoneFile = New cGdArray
        If astrDoneFile.FromFile(DoneFile) Then
            Select Case UCase(astrDoneFile(0))
                Case "DONE"
                    nReturn = eGDDownloadStatus_Done
                Case "ABORTED"
                    nReturn = eGDDownloadStatus_Aborted
                Case "ERROR"
                    nReturn = eGDDownloadStatus_Error
                    strError = astrDoneFile(1)
            End Select
        Else
            nReturn = eGDDownloadStatus_Error
            strError = "Done file could not be read"
        End If
        
        KillFile DoneFile
        m.bDownloading = False
    Else
        nReturn = eGDDownloadStatus_Downloading
    End If
    
    Status = nReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cDownloader.Status"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DownloaderIsRunning
'' Description: Determine if the downloader program is running
'' Inputs:      None
'' Returns:     True if running, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DownloaderIsRunning() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    bReturn = False
    If KillProcess(m.strCaption, True) > 0 Then
        bReturn = True
    End If
    
    DownloaderIsRunning = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cDownloader.DownloaderIsRunning"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToFile
'' Description: Write a file from the class information to send to the process
'' Inputs:      Filename
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ToFile(ByVal strFileName As String)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim File As cDownloaderFile         ' File information from the collection
    Dim strSection As String            ' Section to write to in the INI file

    SetIniFileProperty "Caption", Caption, "Form", strFileName
    SetIniFileProperty "Note", Note, "Form", strFileName
    
    SetIniFileProperty "UserName", UserName, "Server", strFileName
    SetIniFileProperty "Password", Password, "Server", strFileName
    SetIniFileProperty "IP", IP, "Server", strFileName
    SetIniFileProperty "Port", Port, "Server", strFileName
    
    SetIniFileProperty "NumberOfFiles", NumberOfFiles, "Files", strFileName
    SetIniFileProperty "DoneFile", DoneFile, "Files", strFileName
    
    For lIndex = 1 To m.Files.Count
        strSection = "File" & Str(lIndex)
        Set File = m.Files(lIndex)
        
        SetIniFileProperty "ServerPath", File.ServerPath, strSection, strFileName
        SetIniFileProperty "ServerFilename", File.ServerFilename, strSection, strFileName
        SetIniFileProperty "ServerPathCRC", File.ServerPathCrc, strSection, strFileName
        SetIniFileProperty "ServerFilenameCRC", File.ServerFilenameCrc, strSection, strFileName
        SetIniFileProperty "LocalPath", File.LocalPath, strSection, strFileName
        SetIniFileProperty "LocalFilename", File.LocalFilename, strSection, strFileName
        SetIniFileProperty "LocalPathCRC", File.LocalPathCrc, strSection, strFileName
        SetIniFileProperty "LocalFilenameCRC", File.LocalFilenameCrc, strSection, strFileName
        SetIniFileProperty "Size", File.Size, strSection, strFileName
        SetIniFileProperty "CRC", File.Crc, strSection, strFileName
        If File.IsZipFile Then
            SetIniFileProperty "IsZipFile", 1, strSection, strFileName
        Else
            SetIniFileProperty "IsZipFile", 0, strSection, strFileName
        End If
        SetIniFileProperty "ZipPath", File.ZipPath, strSection, strFileName
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cDownloader.ToFile"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize the class members when the class is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    m.strCaption = "Historical Data Downloader"
    m.strNote = ""
    m.strUserName = ""
    m.strPassword = ""
    m.strIP = ""
    m.iPort = 0
    m.strDoneFile = AddSlash(App.Path) & "Downloader\Finished.TXT" ' default
       
    Set m.Files = New cGdTree
    
    m.bDownloading = DownloaderIsRunning

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cDownloader.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Terminate the class members when the class is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.Files = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cDownloader.Class_Terminate"
    
End Sub
