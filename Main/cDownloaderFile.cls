VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDownloaderFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cDownloaderFile.cls
'' Description: Information about a file to download
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
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    strServerPath As String             ' Path for the file on the server
    strServerFilename As String         ' Name of the file on the server
    strServerPathCrc As String          ' Path for the CRC file on the server
    strServerFilenameCrc As String      ' Name of the CRC file on the server
    strLocalPath As String              ' Path to store the file on the local machine
    strLocalFilename As String          ' Name of the file on the local machine
    strLocalPathCrc As String           ' Path to store the CRC file on the local machine
    strLocalFilenameCrc As String       ' Name of the CRC file on the local machine
    dSize As Double                     ' Size of the file
    strCrc As String                    ' CRC of the file
    bIsZipFile As Boolean               ' Is this a file that needs to be unzipped?
    strZipPath As String                ' Path to unzip the files into on the local machine
End Type
Private m As mPrivate

Public Property Get ServerPath() As String
    ServerPath = m.strServerPath
End Property
Public Property Let ServerPath(ByVal strServerPath As String)
    m.strServerPath = strServerPath
End Property

Public Property Get ServerFilename() As String
    ServerFilename = m.strServerFilename
End Property
Public Property Let ServerFilename(ByVal strServerFilename As String)
    m.strServerFilename = strServerFilename
End Property

Public Property Get ServerPathCrc() As String
    ServerPathCrc = m.strServerPathCrc
End Property
Public Property Let ServerPathCrc(ByVal strServerPathCrc As String)
    m.strServerPathCrc = strServerPathCrc
End Property

Public Property Get ServerFilenameCrc() As String
    ServerFilenameCrc = m.strServerFilenameCrc
End Property
Public Property Let ServerFilenameCrc(ByVal strServerFilenameCrc As String)
    m.strServerFilenameCrc = strServerFilenameCrc
End Property

Public Property Get LocalPath() As String
    LocalPath = m.strLocalPath
End Property
Public Property Let LocalPath(ByVal strLocalPath As String)
    m.strLocalPath = strLocalPath
End Property

Public Property Get LocalFilename() As String
    LocalFilename = m.strLocalFilename
End Property
Public Property Let LocalFilename(ByVal strLocalFilename As String)
    m.strLocalFilename = strLocalFilename
End Property

Public Property Get LocalPathCrc() As String
    LocalPathCrc = m.strLocalPathCrc
End Property
Public Property Let LocalPathCrc(ByVal strLocalPathCrc As String)
    m.strLocalPathCrc = strLocalPathCrc
End Property

Public Property Get LocalFilenameCrc() As String
    LocalFilenameCrc = m.strLocalFilenameCrc
End Property
Public Property Let LocalFilenameCrc(ByVal strLocalFilenameCrc As String)
    m.strLocalFilenameCrc = strLocalFilenameCrc
End Property

Public Property Get Size() As Double
    Size = m.dSize
End Property
Public Property Let Size(ByVal dSize As Double)
    m.dSize = dSize
End Property

Public Property Get Crc() As String
    Crc = m.strCrc
End Property
Public Property Let Crc(ByVal strCrc As String)
    m.strCrc = strCrc
End Property

Public Property Get IsZipFile() As Boolean
    IsZipFile = m.bIsZipFile
End Property
Public Property Let IsZipFile(ByVal bIsZipFile As Boolean)
    m.bIsZipFile = bIsZipFile
End Property

Public Property Get ZipPath() As String
    ZipPath = m.strZipPath
End Property
Public Property Let ZipPath(ByVal strZipPath As String)
    m.strZipPath = strZipPath
End Property

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

    m.strServerPath = ""
    m.strServerFilename = ""
    m.strLocalPath = ""
    m.strLocalFilename = ""
    m.dSize = 0#
    m.strCrc = ""
    m.bIsZipFile = False
    m.strZipPath = ""

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cDownloaderFile.Class_Initialize"
    
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

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cDownloaderFile.Class_Terminate"
    
End Sub
