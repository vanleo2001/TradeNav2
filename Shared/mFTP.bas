Attribute VB_Name = "mFTP"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        mFTP.BAS
'' Description: Common FTP routines to share between our products
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author  Description
'' 01/14/2002   DAJ     Created
'' 02/08/2002   DAJ     Added new Registry Info stuff
'' 02/24/2009   DAJ     Moved registry info stuff into separate module
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Public MsgForm As Form

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckAuthorization
'' Description: Check to see if the user is authorized for this request
'' Inputs:      None
'' Returns:     True if Authorized, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CheckAuthorization() As Boolean
On Error GoTo ErrSection:
    
    Dim fhInput As Integer              ' File handle for the input file
    Dim strBuffer As String             ' Buffer from the input file
    Dim strResponse As String           ' Response Code from Genesis
    Dim strMessage As String            ' Comment from Genesis
    Dim strStatus As String             ' User Status from Genesis
    Dim strUserName As String           ' User Name from Genesis
    Dim lNumDays As Long
    Dim strReturn As String
    Dim dRegisterTime As Double
    Dim strAuthBillMode As String
    
    CheckAuthorization = True
    strResponse = RI_GetIBISResponse
    strStatus = RI_GetIBISStatus
    
    Select Case UCase(strResponse)
        Case "OK"
            CheckAuthorization = True
            If UCase(strStatus) <> "TRIAL" Then KillFile AddSlash(App.Path) & "Reg.INI"
        Case "DENIED"
            CheckAuthorization = False
            Select Case UCase(strStatus)
                Case "TRIAL"
                    dRegisterTime = GetIniFileProperty("RegisterTime", -1#, "Register", AddSlash(App.Path) & "Reg.INI")
                    If dRegisterTime = -1# Then
                        RI_DownloadString "AuthBillMode", strAuthBillMode
                        If strAuthBillMode = "1" Then
                            strReturn = AskBox("h=Authorization Warning ; i=! ; b=+Register|-Exit ; Your trial data service has expired.|You will need to register to continue")
                        Else
                            DisplayMessage "Your trial data service has expired.|Please call Genesis Sales at|(800) 808-DATA to subscribe.", "Authorization Warning"
                        End If
                    ElseIf CDbl(Now) - dRegisterTime <= (15 / 1440) Then
                        DisplayMessage "Your account information has not been refreshed.  Please give us 15 minutes to refresh your account status.", "Authorization Warning"
                    Else
                        DisplayMessage "Your account information has not been refreshed.  Please call Genesis Technical Support at (719) 884-0245", "Authorization Warning"
                    End If
                    Select Case UCase(strReturn)
                        Case "R"
                            GetRegisterProgram
                            
                            If UCase(App.EXEName) = "NAVSUITE" Then
                                KillFile App.Path & "\Subscribe.DON"
                                RunProcess App.Path & "\Subscribe.EXE", "/C", True, vbNormalFocus
                                KillFile App.Path & "\Subscribe.DON"
                        
                                ' If response came back ok from the register, set the
                                ' time stamp so that we know if it has been 15 minutes
                                If UCase(RI_GetIBISResponse) = "OK" Then
                                    SetIniFileProperty "RegisterTime", Now, "Register", AddSlash(App.Path) & "Reg.INI"
                                End If
                            End If
                            Exit Function
                        Case "E"
                            Exit Function
                    End Select
                Case "TERMINATED"
                    DisplayMessage "Your data service has been terminated.||Please contact Genesis Sales at|(800) 808-DATA|to activate your account.|", "Authorization Error"
                Case "CANCELLED"
                    DisplayMessage "Your data service has been cancelled.||Please contact Genesis Sales at|(800) 808-DATA|to activate your account.|", "Authorization Error"
                Case "SUSPENDED"
                    DisplayMessage "Your data service has been suspended.||Please contact Genesis Billing at|(719) 884-0266|to activate your account.|", "Authorization Error"
                Case "ACTIVE"
                    ' Should never get here (in theory)
                    DisplayMessage "There is a problem with your|data service account.||Please call Genesis Billing at|(719) 884-0266.|", "Authorization Error"
            End Select
        Case "ERROR"
            CheckAuthorization = False
            strMessage = RI_GetMessage
            If Len(Trim(strMessage)) = 0 Then strMessage = "There was an error while trying to authorize your account.||Please call Genesis Billing at (719) 884-0266."
            DisplayMessage strMessage, "Authorization Error"
    End Select
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "mFTP.CheckAuthorization", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CheckData
'' Description: Make sure that there were no data errors that came back
'' Inputs:      None
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CheckData() As Boolean
On Error GoTo ErrSection:
    
    Dim astrFile As cGdArray            ' Array to store the file into
    Dim lIndex As Long                  ' Index into a for loop
    Dim astrError As cGdArray           ' Array of errors
    Dim bNoData As Boolean
    Dim bRealtime As Boolean
    
#If TRADENAV_EXE Then
    bRealtime = g.RealTime.Active
    ''FileCopy App.Path & "\lasterr.nod", App.Path & "\ftp\dataerr.txt"
#End If

    CheckData = True
    If FileExist(AddSlash(App.Path) & "FTP\DataErr.TXT") Then
        Set astrFile = New cGdArray
        astrFile.Create eGDARRAY_Strings
        Set astrError = New cGdArray
        astrError.Create eGDARRAY_Strings
        If astrFile.FromFile(AddSlash(App.Path) & "FTP\DataErr.TXT") Then
            For lIndex = 0 To astrFile.Size - 1
                If InStr(UCase(astrFile(lIndex)), "CODE =") > 0 Then
                    Select Case Trim(Parse(astrFile(lIndex), "=", 2))
                        Case "1"
                            astrError.Add "Cannot open specified request file"
                        Case "2"
                            astrError.Add "Badly formatted request"
                        Case "3"
                            astrError.Add "Request file is empty"
                        Case "4"
                            astrError.Add "Wildcard must be at the end of a symbol"
                        Case "5"
                            astrError.Add "Cannot open GDB"
                        Case "102", "103"
                            If Not bRealtime And Not bNoData Then
                                bNoData = True
                                astrError.Add "No data found for requested symbols" ' & UCase(Parse(astrFile(lIndex + 2), ";", 5))
                            End If
                        Case "104"
                            astrError.Add "Incorrect security type: " & UCase(Parse(astrFile(lIndex + 2), ";", 4))
                        Case "106"
                            astrError.Add "Data feed unavailable"
                        Case Else
                            'astrError.Add "Error retrieving data: " & Trim(Parse(astrFile(lIndex), "=", 2))
                    End Select
                End If
            Next lIndex
        End If
        
        If astrError.Size > 0 Then
            CheckData = False
            DisplayMessage astrError.JoinFields("|"), "Data Error"
        ElseIf Not bRealtime Then
            CheckData = False
            DisplayMessage "Unknown Error Retrieving Data", "Data Error"
        End If
        
        FileCopy AddSlash(App.Path) & "FTP\DataErr.TXT", AddSlash(App.Path) & "LastErr.TXT"
    End If

ErrExit:
    Exit Function

ErrSection:
    RaiseError "mFTP.CheckData", eGDRaiseError_Raise

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    TransferDataService
'' Description: Attempts to set up a data service transfer
'' Inputs:      None
'' Returns:     TRUE on success, FALSE on failure
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TransferDataService() As Boolean
On Error GoTo ErrSection:

    Dim strResponse As String           ' Response from the AskBox
    Dim astrFile As New cGdArray        ' File I/O array
    Dim strAction As String             ' Action from the Access.INF
    Dim strCount As String              ' Count of transfers this month
    Dim strMessage As String            ' Message form the Access.INF
    Dim strMachineID As String          ' Machine ID being transferred from
    Dim lIndex As Long                  ' Index for a for loop
    
    strResponse = AskBox("h=Transfer Data Service ; i=? ; b=+Transfer|-Cancel ; " & _
                            "This data service is only allowed to update from " & _
                            "one machine.  You are allowed to transfer this data " & _
                            "service to another machine a maximum of 4 times a month." & _
                            "||Would you like to transfer now?|")
    If UCase(strResponse) = "T" Then
        astrFile.Create eGDARRAY_Strings
        astrFile.Add "*Request DSrvTransfer"
        astrFile.Add "+ServID:" & Str(RI_GetDataServiceID)
        astrFile.Add "+PassWord:" & RI_GetUserPassword
        
        If FtpRequest(astrFile, True) = True Then
            If FileExist(AddSlash(App.Path) & "FTP\Access.INF") Then
                If astrFile.FromFile(AddSlash(App.Path) & "FTP\Access.INF") Then
                    For lIndex = 0 To astrFile.Size - 1
                        Select Case UCase(Parse(astrFile(lIndex), ":", 1))
                            Case "+ACTION"
                                strAction = StripStr(Parse(astrFile(lIndex), ":", 2), Chr(34))
                            Case "+TRANSFERCOUNT"
                                strCount = CStr(CLng(ValOfText(StripStr(Parse(astrFile(lIndex), ":", 2), Chr(34)))) + 1)
                            Case "+MESSAGE"
                                strMessage = StripStr(Parse(astrFile(lIndex), ":", 2), Chr(34))
                            Case "+MACHINEID"
                                strMessage = StripStr(Parse(astrFile(lIndex), ":", 2), Chr(34))
                        End Select
                    Next lIndex
                    
                    Select Case UCase(Trim(strAction))
                        Case "DSRVTRANSFER SUCCESS"
                            TransferDataService = True
                            If UCase(strMachineID) = RI_GetMachineID Then
                                DisplayMessage "Data Service is Ready for Transfer.||This will be transfer number|" & Trim(strCount) & " out of 4|for the month.", "Transfer Success"
                            Else
                                DisplayMessage "Data Service has been Transferred.||This is transfer number|" & Trim(strCount) & " out of 4|for the month.", "Transfer Success"
                            End If
                        Case "NO SUCH DSRVID"
                            DisplayMessage "Data Service does not exist: " & Str(RI_GetDataServiceID), "Transfer Error"
                        Case "BAD PASSWORD"
                            DisplayMessage "Password is invalid", "Transfer Error"
                        Case "COUNT REACHED"
                            DisplayMessage "You have exceeded the maximum number of transfers for the month", "Transfer Error"
                        Case "ERROR"
                            If Len(Trim(strMessage)) = 0 Then
                                DisplayMessage "An error occured during transfer", "Transfer Error"
                            Else
                                DisplayMessage strMessage, "Transfer Error"
                            End If
                        Case Else
                            DisplayMessage "An error occured during transfer", "Transfer Error"
                    End Select
                Else
                    DisplayMessage "Unable to open Response File", "Transfer Error"
                End If
                
                FileCopy AddSlash(App.Path) & "FTP\Access.INF", AddSlash(App.Path) & "FTP\Backup\Access.INF"
                KillFile AddSlash(App.Path) & "FTP\Access.INF"
            Else
                DisplayMessage "Response File does not exist", "Transfer Error"
            End If
        End If
    End If
    
ErrExit:
    Set astrFile = Nothing
    Exit Function

ErrSection:
    RaiseError "mFTP.TransferDataService", eGDRaiseError_Raise

End Function

Public Sub DisplayMessage(ByVal strMessage$, ByVal strTitle$)
On Error Resume Next

    If MsgForm Is Nothing Then
        InfBox strMessage, "!", , strTitle
    Else
        Replace strMessage, "|", " "
        MsgForm.AddDetail strMessage
    End If

End Sub
