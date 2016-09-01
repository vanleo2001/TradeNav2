Attribute VB_Name = "mMain"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        mMain.bas
'' Description: Main module for the Trade Navigator Archive project
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 02/13/2009   DAJ         Created
'' 06/11/2009   DAJ         Added LastDirInPath function
'' 08/19/2009   DAJ         Changed command line, delete old TNA files
'' 09/02/2009   DAJ         Changed default file name to include century
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Global Const kArchiveSettings = "Provided\Archive.INI"
Global Const kRegistryFile = "Archive\RegSettings.TXT"
Global Const kBackupInfo = "Info\Backup.RTF"
Global Const kRestoreInfo = "Info\Restore.RTF"
Global Const kUndoRestoreInfo = "Info\UndoRestore.RTF"

Global Const kArchiveFile = "Archive.TNA"
Global Const kUndoRestoreFile = "BeforeRestore.TNA"

Global Const kArchiveVersion = 1&

Public Enum eGD_OverwriteMode
    eGD_OverwriteAll = 0
    eGD_OverwriteNewer
    eGD_AskForOverwrite
End Enum

'RH
Public Enum eStyleColorTypes
    eForm_Background = 0
    eFrame_Background = 1
    eFrame_Border = 2
    eButton_Background = 3
    eButton_Border = 4
    eCheck_Border = 5
    eCheck_Background = 6
    eCheck_Forecolor = 7
End Enum


Type gGlobal
    strIniFile As String                ' INI file for the application
    CommandLine As cCommandLine         ' Command line arguments
    
    'RH added styler
    Styler As New cStyler
End Type
Global g As gGlobal

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Main
'' Description: Startup routine for the application
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Main()
On Error GoTo ErrSection:

    Dim strCommandLine As String        ' Command line arguments
    Dim astrArgs As cGdArray            ' Command line arguments

    ChangePath App.Path
    g.strIniFile = AddSlash(App.Path) & "TNArchive.INI"
    
    If App.PrevInstance Then
        ActivatePrevInstance
        End
    End If
    
    Set g.CommandLine = New cCommandLine
    g.CommandLine.FromCommandLine Command$
    
    ' Delete any archive files older than 14 days...
    'KillFile AddSlash(App.Path) & "Archive\*.TNA /o=-14"
    DeleteOldFiles
    
        'RH - initialize styler
    With g.Styler
        '''.SetButtonStyleDefault iCtlBtnStyle_Flat
        'Form
        .SetColor eForm_Background, vbWhite
        
        'Frames
        .SetColor eFrame_Background, vbWhite
        .SetColor eFrame_Border, vbBlue
        
        'Buttons
        .SetColor eButton_Background, vbWhite
        .SetColor eButton_Border, vbBlue
        
        'Checkboxes (checkmarks)
        .SetColor eCheck_Border, vbBlue
        
    End With
    

    frmMain.Show

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mMain.Main", eGDRaiseError_Show
    If FormIsLoaded("frmMain") Then
        Unload frmMain
        DoEvents
    End If
    End
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LogToFile
'' Description: Dump the message to the log file
'' Inputs:      Message
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LogToFile(ByVal strMessage As String)
On Error Resume Next

    Dim fh As Integer                   ' File handle to open file with

    fh = FreeFile
    Open AddSlash(App.Path) & "Archive\" & Format(Now, "YYYYMMDD") & ".LOG" For Append Shared As #fh
    If fh Then
        Print #fh, Format$(Now, "hh:mm:ss") & " - " & strMessage
        Close #fh
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShutdownMsg
'' Description: Show the shutdown message
'' Inputs:      Action, Error Caption
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShutdownMsg(Optional ByVal strAction$ = "before continuing", Optional ByVal strErrCaption$ = "Error")
On Error GoTo ErrSection:

    Dim strText As String               ' Text for the message box

    strAction = " " & Trim(strAction)
    strText = "Error: You must shut down " & g.CommandLine.Caption & strAction
    InfBox strText, "!", , strErrCaption
    LogToFile strText
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mMain.ShutdownMsg"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LastDirInPath
'' Description: Determine the last directory in the path
'' Inputs:      Command Line
'' Returns:     Array of Arguments
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LastDirInPath(ByVal strPath As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim lPos As Long                    ' Position in the string
    
    strReturn = strPath
    If Right(strReturn, 1) = "\" Then
        strReturn = Left(strReturn, Len(strReturn) - 1)
    End If
    
    lPos = At(strReturn, "\", -1)
    If lPos > 0 Then
        strReturn = Right(strReturn, Len(strReturn) - lPos)
    End If
    
    LastDirInPath = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.LastDirInPath"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeleteOldFiles
'' Description: Delete old archive files
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DeleteOldFiles()
On Error GoTo ErrSection:

    Dim strFilePath As String           ' File path
    Dim astrFiles As cGdArray           ' Array of archive files
    Dim lIndex As Long                  ' Index into a for loop
    Dim lFileDate As Long               ' File date
    Dim lToday As Long                  ' Today's date
    Dim lFileKeptMask As Long           ' Mask of if a file was kept
    Dim strFile As String               ' Filename and path of the file to check
        
    strFilePath = AddSlash(App.Path) & "Archive\"
    
    Set astrFiles = New cGdArray
    astrFiles.GetMatchingFiles strFilePath & "*.TNA", False, False, False
    astrFiles.Sort
    
    If astrFiles.Size > 0 Then
        lToday = Date
        
        For lIndex = astrFiles.Size - 1 To 0 Step -1
            If IsDateTimeName(astrFiles(lIndex)) = True Then
                strFile = strFilePath & astrFiles(lIndex)
                lFileDate = Int(FileDate(strFile))
                Select Case lToday - lFileDate
                    Case Is < 7
                        ' Keep all of these files
                        
                    Case Is < 14
                        If GetBit(lFileKeptMask, 1) = False Then
                            SetBit lFileKeptMask, 1, True
                        Else
                            KillFile strFile
                        End If
                    
                    Case Is < 21
                        If GetBit(lFileKeptMask, 2) = False Then
                            SetBit lFileKeptMask, 2, True
                        Else
                            KillFile strFile
                        End If
                    
                    Case Is < 30
                        If GetBit(lFileKeptMask, 3) = False Then
                            SetBit lFileKeptMask, 3, True
                        Else
                            KillFile strFile
                        End If
                    
                    Case Is < 60
                        If GetBit(lFileKeptMask, 4) = False Then
                            SetBit lFileKeptMask, 4, True
                        Else
                            KillFile strFile
                        End If
                    
                    Case Is < 90
                        If GetBit(lFileKeptMask, 5) = False Then
                            SetBit lFileKeptMask, 5, True
                        Else
                            KillFile strFile
                        End If
                    
                    Case Else
                        If GetBit(lFileKeptMask, 6) = False Then
                            SetBit lFileKeptMask, 6, True
                        Else
                            KillFile strFile
                        End If
                    
                End Select
            End If
        Next lIndex
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mMain.DeleteOldFiles"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsDateTimeName
'' Description: Is the given filename one of our default date/time names?
'' Inputs:      Filename
'' Returns:     True if Date/Time name, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsDateTimeName(ByVal strFilename As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim strFileBase As String           ' File base of the given filename
    
    bReturn = True
    strFileBase = FileBase(strFilename)
    
    If Len(strFileBase) = 13 Then
        For lIndex = 1 To 13
            If lIndex = 7 Then
                If Mid(strFileBase, lIndex, 1) <> "_" Then
                    bReturn = False
                    Exit For
                End If
            Else
                If IsDigit(Mid(strFileBase, lIndex, 1)) = False Then
                    bReturn = False
                    Exit For
                End If
            End If
        Next lIndex
    ElseIf Len(strFileBase) = 15 Then
        For lIndex = 1 To 15
            If lIndex = 9 Then
                If Mid(strFileBase, lIndex, 1) <> "_" Then
                    bReturn = False
                    Exit For
                End If
            Else
                If IsDigit(Mid(strFileBase, lIndex, 1)) = False Then
                    bReturn = False
                    Exit For
                End If
            End If
        Next lIndex
    Else
        bReturn = False
    End If
    
    IsDateTimeName = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mMain.IsDateTimeName"
    
End Function
