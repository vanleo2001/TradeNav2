Attribute VB_Name = "mRegInfo"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        mRegInfo.BAS
'' Description: Routines to get information out of the registry
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author  Description
'' 02/24/2009   DAJ     Split registry functions from mFTP.bas
'' 09/16/2009   DAJ     Added ability to get/set PFG Access Key information
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Declare Sub RI_GetMachID Lib "RInfoDLL" Alias "GetMachineID" (ByVal hValue As Long)
Declare Sub RI_GetUserPass Lib "RInfoDLL" Alias "GetUserPassword" (ByVal hValue As Long)
Declare Function RI_GetDataServiceID Lib "RInfoDLL" Alias "DataServiceID" () As Long
Declare Function RI_GetLastDataServiceID Lib "RInfoDLL" Alias "LastDataServiceID" () As Long
Declare Sub RI_IBISResponse Lib "RInfoDLL" Alias "IBISResponse" (ByVal hValue As Long)
Declare Sub RI_IBISStatus Lib "RInfoDLL" Alias "IBISStatus" (ByVal hValue As Long)
Declare Function RI_GetExpirationDate Lib "RInfoDLL" Alias "ExpirationDate" () As Long
Declare Sub RI_Comment Lib "RInfoDLL" Alias "Comment" (ByVal hValue As Long)
Declare Sub RI_Message Lib "RInfoDLL" Alias "Message" (ByVal hValue As Long)

Declare Function RI_SetDataServiceID Lib "RInfoDLL" Alias "SetDataServiceID" (ByVal lServiceID As Long) As Long
Declare Function RI_SetUserPass Lib "RInfoDLL" Alias "SetUserPassword" (ByVal hValue As Long) As Long

Declare Sub RI_DLString Lib "RInfoDLL" Alias "DLString" (ByVal hKey As Long, hVal As Long)

Declare Function RI_HonestDate Lib "RInfoDLL" Alias "HonestDate" () As Long ' CCYYMMDD

Declare Sub RI_GetPAK Lib "RInfoDLL" Alias "GetPAK" (ByVal hValue As Long)
Declare Function RI_SetPAK Lib "RInfoDLL" Alias "SetPAK" (ByVal hValue As Long) As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RI_GetMachineID
'' Description: Retrieve the machine ID out of the registry and convert to VB
'' Inputs:      None
'' Returns:     Machine ID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RI_GetMachineID() As String
On Error GoTo ErrSection:

    Dim strMachID As New cGdArray
    
    ChkPath
    
    strMachID.Create eGDARRAY_gdString
    RI_GetMachID strMachID.ArrayHandle
    RI_GetMachineID = gdGetStr(strMachID.ArrayHandle)
    
ErrExit:
    Set strMachID = Nothing
    Exit Function

ErrSection:
    RaiseError "mRegInfo.RI_GetMachineID"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RI_GetUserPassword
'' Description: Retrieve the password out of the registry and convert to VB
'' Inputs:      None
'' Returns:     Password
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RI_GetUserPassword() As String
On Error GoTo ErrSection:

    Dim strUserPassword As New cGdArray
    
    ChkPath
    
    strUserPassword.Create eGDARRAY_gdString
    RI_GetUserPass strUserPassword.ArrayHandle
    RI_GetUserPassword = gdGetStr(strUserPassword.ArrayHandle)
    
ErrExit:
    Set strUserPassword = Nothing
    Exit Function

ErrSection:
    RaiseError "mRegInfo.RI_GetUserPassword"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RI_GetIBISResponse
'' Description: Retrieve the IBIS response out of the registry and convert to VB
'' Inputs:      None
'' Returns:     IBIS Response
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RI_GetIBISResponse() As String
On Error GoTo ErrSection:

    Dim strIBISResponse As New cGdArray
    
    ChkPath
    
    strIBISResponse.Create eGDARRAY_gdString
    RI_IBISResponse strIBISResponse.ArrayHandle
    RI_GetIBISResponse = gdGetStr(strIBISResponse.ArrayHandle)
    
ErrExit:
    Set strIBISResponse = Nothing
    Exit Function

ErrSection:
    RaiseError "mRegInfo.RI_GetIBISResponse"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RI_GetIBISStatus
'' Description: Retrieve the IBIS status out of the registry and convert to VB
'' Inputs:      None
'' Returns:     IBIS Status
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RI_GetIBISStatus() As String
On Error GoTo ErrSection:

    Dim strIBISStatus As New cGdArray
    
    ChkPath
    
    strIBISStatus.Create eGDARRAY_gdString
    RI_IBISStatus strIBISStatus.ArrayHandle
    RI_GetIBISStatus = gdGetStr(strIBISStatus.ArrayHandle)
    
ErrExit:
    Set strIBISStatus = Nothing
    Exit Function

ErrSection:
    RaiseError "mRegInfo.RI_GetIBISStatus"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RI_GetComment
'' Description: Retrieve the comment out of the registry and convert to VB
'' Inputs:      None
'' Returns:     Comment
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RI_GetComment() As String
On Error GoTo ErrSection:
    
    Dim strComment As New cGdArray
    
    ChkPath
    
    strComment.Create eGDARRAY_gdString
    RI_Comment strComment.ArrayHandle
    RI_GetComment = gdGetStr(strComment.ArrayHandle)
    
ErrExit:
    Set strComment = Nothing
    Exit Function

ErrSection:
    RaiseError "mRegInfo.RI_GetComment"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RI_GetMessage
'' Description: Retrieve the message out of the registry and convert to VB
'' Inputs:      None
'' Returns:     Message
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RI_GetMessage() As String
On Error GoTo ErrSection:
    
    Dim strMessage As New cGdArray
    
    ChkPath
    
    strMessage.Create eGDARRAY_gdString
    RI_Message strMessage.ArrayHandle
    RI_GetMessage = gdGetStr(strMessage.ArrayHandle)
    
ErrExit:
    Set strMessage = Nothing
    Exit Function

ErrSection:
    RaiseError "mRegInfo.RI_GetMessage"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RI_SetUserPassword
'' Description: Convert password to gdString and store it
'' Inputs:      Password
'' Returns:     Storage Success
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RI_SetUserPassword(ByVal strPassword As String) As Long
On Error GoTo ErrSection:

    Dim strUserPassword As New cGdArray
    
    ChkPath
    
    strUserPassword.Create eGDARRAY_gdString
    gdSetStr strUserPassword.ArrayHandle, 0, strPassword
    RI_SetUserPassword = RI_SetUserPass(strUserPassword.ArrayHandle)
    
ErrExit:
    Set strUserPassword = Nothing
    Exit Function

ErrSection:
    RaiseError "mRegInfo.RI_SetUserPassword"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RI_DownloadString
'' Description: Retrieve the download string out of the registry and convert to VB
'' Inputs:      Key, Value
'' Returns:     Download String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RI_DownloadString(ByVal strKey As String, strValue As String)
On Error GoTo ErrSection:

    Dim gdstrKey As New cGdArray
    Dim gdStrValue As New cGdArray
    
    ChkPath
    
    gdstrKey.Create eGDARRAY_gdString
    gdStrValue.Create eGDARRAY_gdString
    
    gdSetStr gdstrKey.ArrayHandle, 0, strKey
    RI_DLString gdstrKey.ArrayHandle, gdStrValue.ArrayHandle
    strValue = gdGetStr(gdStrValue.ArrayHandle)
    
ErrExit:
    Set gdstrKey = Nothing
    Set gdStrValue = Nothing
    Exit Function

ErrSection:
    RaiseError "mRegInfo.RI_DownloadString"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RI_GetPfgAccessKeys
'' Description: Retrieve the Access Keys out of the registry and convert to VB
'' Inputs:      None
'' Returns:     Access Keys
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RI_GetPfgAccessKeys() As String
On Error GoTo ErrSection:

    Dim gdstrAccessKeys As New cGdArray ' gdString version of String
    
    ChkPath
    
    gdstrAccessKeys.Create eGDARRAY_gdString
    RI_GetPAK gdstrAccessKeys.ArrayHandle
    RI_GetPfgAccessKeys = gdGetStr(gdstrAccessKeys.ArrayHandle)
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mRegInfo.RI_GetPfgAccessKeys"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RI_SetPfgAccessKeys
'' Description: Convert Access Keys to gdString and store it
'' Inputs:      Access Keys
'' Returns:     Storage Success
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RI_SetPfgAccessKeys(ByVal strAccessKeys As String) As Long
On Error GoTo ErrSection:

    Dim gdstrAccessKeys As New cGdArray ' gdString version of String
    
    ChkPath
    
    gdstrAccessKeys.Create eGDARRAY_gdString
    gdSetStr gdstrAccessKeys.ArrayHandle, 0, strAccessKeys
    RI_SetPfgAccessKeys = RI_SetPAK(gdstrAccessKeys.ArrayHandle)
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "mRegInfo.RI_SetPfgAccessKeys"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RI_GetPfgAccessKey
'' Description: Retrieve the Access Key out of the registry and convert to VB
'' Inputs:      User Name
'' Returns:     Access Key
''
'' Format:      UserName|AccessKey,UserName|AccessKey...
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RI_GetPfgAccessKey(ByVal strUserName As String) As String
On Error GoTo ErrSection:

    Dim astrAccessKeys As New cGdArray  ' Array of access key information
    Dim strReturn As String             ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    ChkPath
    
    strReturn = ""
    
    astrAccessKeys.SplitFields RI_GetPfgAccessKeys, ","
    For lIndex = 0 To astrAccessKeys.Size - 1
        If Parse(astrAccessKeys(lIndex), "|", 1) = strUserName Then
            strReturn = Parse(astrAccessKeys(lIndex), "|", 2)
            Exit For
        End If
    Next lIndex
    
    RI_GetPfgAccessKey = strReturn
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "mRegInfo.RI_GetPfgAccessKey"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RI_SetPfgAccessKey
'' Description: Convert Access Key to gdString and store it
'' Inputs:      User Name, Access Key
'' Returns:     Storage Success
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RI_SetPfgAccessKey(ByVal strUserName As String, ByVal strAccessKey As String) As Long
On Error GoTo ErrSection:

    Dim astrAccessKeys As New cGdArray  ' Array of access key information
    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' Was the user name found in the array?
    
    ChkPath
    
    bFound = False
    astrAccessKeys.SplitFields RI_GetPfgAccessKeys, ","
    For lIndex = 0 To astrAccessKeys.Size - 1
        If Parse(astrAccessKeys(lIndex), "|", 1) = strUserName Then
            astrAccessKeys(lIndex) = strUserName & "|" & strAccessKey
            bFound = True
            Exit For
        End If
    Next lIndex
    
    If bFound = False Then
        astrAccessKeys.Add strUserName & "|" & strAccessKey
    End If
    
    RI_SetPfgAccessKey = RI_SetPfgAccessKeys(astrAccessKeys.JoinFields(","))

ErrExit:
    Exit Function

ErrSection:
    RaiseError "mRegInfo.RI_SetPfgAccessKey"

End Function

' TLB 1/3/2012: we ran into an issue where the RINFO dll couldn't be found/loaded
' because the current path had been changed, so let's reset it on first call
Private Sub ChkPath()
On Error Resume Next

    Static bAlreadyDone As Boolean
    If Not bAlreadyDone Then
        bAlreadyDone = True
        ChangePath App.Path
    End If

End Sub

' TLB: use CAUTION -- only need this for a really unusual situation
' (e.g. if 2 machines have same MID -- e.g. from a low-level hard drive copy)
Public Sub RI_RegenerateMID()

    On Error Resume Next
    Dim s1$, s2$
    s1 = StrReverse("erawtfos") & "\" & StrReverse("tfosorcim") & "\" & StrReverse("tnerapt")
    s2 = StrReverse("datagen")
    DeleteRegistryValue rkLocalMachine, s1, s2 ' delete old one
    s2 = RI_GetMachineID ' regenerates a new one

End Sub
