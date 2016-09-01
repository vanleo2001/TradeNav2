VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCommandLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cCommandLine
'' Description: Class to handle the command line arguments to the function
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 08/19/2009   DAJ         Created
'' 08/22/2009   DAJ         Removed Caption from the arguments
'' 08/16/2013   DAJ         Make sure default path has a slash on the end
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Enum eArchiveArgs
    eArchiveArg_Action = 0
    eArchiveArg_File
    eArchiveArg_Path
End Enum

Private Type mPrivate
    astrArgs As cGdArray                ' Array of arguments split out
End Type
Private m As mPrivate

Public Property Get Caption() As String
On Error GoTo ErrSection:

    Caption = GetIniFileProperty("Title", "Trade Navigator", "Main", g.strIniFile)
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cCommandLine.Caption.Get"
    
End Property

Public Property Get Action() As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function

    If m.astrArgs.Size > eArchiveArg_Action Then
        Select Case UCase(Trim(StripStr(m.astrArgs(eArchiveArg_Action), Chr(34))))
            Case "BACKUP", "0"
                lReturn = 0&
            Case "RESTORE", "1"
                lReturn = 1&
            Case Else
                lReturn = -1&
        End Select
    Else
        lReturn = -1&
    End If
    
    Action = lReturn
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cCommandLine.Action.Get"
    
End Property

Public Property Get ArchiveFile() As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function

    If m.astrArgs.Size > eArchiveArg_File Then
        strReturn = Trim(StripStr(m.astrArgs(eArchiveArg_File), Chr(34)))
    Else
        strReturn = ""
    End If
    
    ArchiveFile = strReturn
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cCommandLine.ArchiveFile.Get"
    
End Property

Public Property Get ArchivePath() As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function

    If m.astrArgs.Size > eArchiveArg_Path Then
        strReturn = Trim(StripStr(m.astrArgs(eArchiveArg_Path), Chr(34)))
    Else
        strReturn = GetIniFileProperty("Path", AddSlash(App.Path) & "Archive", "Main", g.strIniFile)
    End If
    
    ' 08/16/2013 DAJ: We found that if we are using the default path of C:\Genesis\Navigator Suite\Archive
    ' without the slash on the end, we are removing the Archive part of the path when we do a backup
    ' and thus storing an archive both in the archive folder and the Navigator Suite folder...
    If strReturn = AddSlash(App.Path) & "Archive" Then
        strReturn = AddSlash(strReturn)
    End If
    
    ArchivePath = strReturn
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cCommandLine.ArchivePath.Get"
    
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromCommandLine
'' Description: Split out the arguments from the command line
'' Inputs:      Command Line
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromCommandLine(ByVal strCommandLine As String)
On Error GoTo ErrSection:

    strCommandLine = Trim(strCommandLine)
    Set m.astrArgs = SplitArgs(strCommandLine)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCommandLine.FromCommandLine"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SplitArgs
'' Description: Split command line arguments
'' Inputs:      Command Line
'' Returns:     Array of Arguments
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function SplitArgs(ByVal strCommandLine As String) As cGdArray
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim astrReturn As cGdArray          ' Array to return from function
    Dim lStart As Long                  ' Starting point of the argument
    Dim bInQuotes As Boolean            ' Are we currently in the middle of quotes?
    Dim strArg As String                ' Argument
    
    Set astrReturn = New cGdArray
    astrReturn.Create eGDARRAY_Strings
    
    If Len(strCommandLine) > 0 Then
        bInQuotes = False
        lStart = 1
        strCommandLine = Trim(strCommandLine)
        
        For lIndex = 1 To Len(strCommandLine)
            If Mid(strCommandLine, lIndex, 1) = Chr(34) Then
                bInQuotes = Not bInQuotes
            ElseIf (Mid(strCommandLine, lIndex, 1) = " ") And (bInQuotes = False) Then
                strArg = Mid(strCommandLine, lStart, lIndex - lStart)
                astrReturn.Add strArg
                lStart = lIndex + 1
            End If
        Next lIndex
        
        strArg = Mid(strCommandLine, lStart)
        astrReturn.Add strArg
    End If
    
    Set SplitArgs = astrReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cCommandLine.SplitArgs"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize the class members when object is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.astrArgs = New cGdArray
    m.astrArgs.Create eGDARRAY_Strings

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCommandLine.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up the class members when object is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.astrArgs = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cCommandLine.Class_Terminate"
    
End Sub
