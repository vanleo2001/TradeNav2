VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLogFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Logging:
' - this method of logging is MUCH more efficient than opening a file, writing a line, then closing the file
' - since VB does not have a "flush" command, this class uses C++ calls which do support flushing
' - supports a log file with either a constant name or with the current date as part of the filename
' - the logging file will be automatically closed when the instance of this class is destroyed
'
' Benchmark comparisons (on Tim's laptop Aug 2013):
' - 2.50 ms per each open/write/close (400 logs per second) -- OLD METHOD
' - 0.02 ms per each write/flush (50,000 logs per second) -- NEW METHOD (125x more efficient)
' - 0.01 ms per each write (no flushing) -- a little faster, but not recommended
'
' To Use:
' - declare a new cLogFile in your module
' - should only need to call .OpenFile once (can use ".IsOpen" to only open it if not already open)
' - just keep using .WriteText to log (flushing is recommended)
' - if using a date in the filename, the file will automatically close and re-open when date changes
' - the file can be explicitly closed, but will also be closed whenever the cLogFile is destructed
Option Explicit

Private Type mPrivate
    hFile As Long
    strFile As String
    nCurrentDate As Long ' = 0 if not using AutoDate method
End Type
Private m As mPrivate

Private Sub Class_Terminate()

    CloseFile

End Sub

Public Sub CloseFile()
On Error Resume Next

    If m.hFile <> 0 Then
        gdFileClose m.hFile
        m.hFile = 0
    End If

End Sub

Public Property Get IsOpen() As Boolean
    IsOpen = (m.hFile <> 0)
End Property

' Can use an asterisk in the filename to indicate using the current date (YYYYMMDD) in the filename -- e.g.:
'   "C:\Genesis\Debug.LOG" --> .OpenFile("C:\Genesis\Debug.Log", False)
'   "C:\Genesis\20130721.LOG" --> .OpenFile("C:\Genesis\*.Log", True)
'   "C:\Genesis\TN20130721.LOG" --> .OpenFile("C:\Genesis\TN*.Log", True)
Public Function OpenFile(ByVal strFileName$, Optional ByVal bAppendToExistingFile As Boolean = True) As Boolean
On Error Resume Next
    
    CloseFile
    
    m.nCurrentDate = 0
    If Len(Trim(strFileName)) > 0 Then
        m.strFile = strFileName
        ' replace an asterisk in the filename with the current date (YYYYMMDD)
        If InStr(strFileName, "*") > 0 Then
            m.nCurrentDate = Date
            strFileName = Replace(strFileName, "*", Format(m.nCurrentDate, "YYYYMMDD"))
        End If
    
        ' open file
        If bAppendToExistingFile Then
            m.hFile = gdFileOpen(strFileName, "a+")
        Else
            m.hFile = gdFileOpen(strFileName, "w+")
        End If
        If m.hFile <> 0 Then
            OpenFile = True
        End If
    End If
    
End Function

Public Sub WriteText(ByVal strText$, Optional ByVal bFlush As Boolean = True)
On Error Resume Next

    ' see if we need to open a file for a new date
    If m.nCurrentDate > 0 Then
        If Date <> m.nCurrentDate Then
            OpenFile m.strFile, True
        End If
    End If

    ' write text to file
    If m.hFile <> 0 Then
        strText = Format(Now, "hh:mm:ss") & " (" & Str(gdTickCount) & ") - " & strText
        gdFileWriteLine m.hFile, strText, bFlush
    End If

End Sub

#If 0 Then
' this was just here for benchmarking purposes
Public Sub OldMethod(ByVal strText As String)
On Error Resume Next

    Dim fh As Integer

    fh = FreeFile
    Open AddSlash(App.Path) & Format(Now, "YYYYMMDD") & ".LOG" For Append As #fh
    If fh Then
        Print #fh, Format$(Now, "hh:mm:ss") & " (" & Str(gdTickCount) & ") - " & strText
        Close #fh
    End If

End Sub
#End If
