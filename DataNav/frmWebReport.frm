VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmWebReport 
   AutoRedraw      =   -1  'True
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8370
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   8370
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrUnload 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6135
      Top             =   5175
   End
   Begin SHDocVwCtl.WebBrowser webBrowser 
      Height          =   3570
      Left            =   1050
      TabIndex        =   0
      Top             =   915
      Width           =   5550
      ExtentX         =   9790
      ExtentY         =   6297
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmWebReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmWebReport.frm
'' Description: Navigate to a given URL
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 08/21/2015   DAJ         Send URL's through FixURL
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Enum eHttpError
    HTTP_STATUS_BAD_REQUEST = 400
    HTTP_STATUS_DENIED = 401
    HTTP_STATUS_PAYMENT_REQ = 402
    HTTP_STATUS_FORBIDDEN = 403
    HTTP_STATUS_NOT_FOUND = 404
    HTTP_STATUS_BAD_METHOD = 405
    HTTP_STATUS_NONE_ACCEPTABLE = 406
    HTTP_STATUS_PROXY_AUTH_REQ = 407
    HTTP_STATUS_REQUEST_TIMEOUT = 408
    HTTP_STATUS_CONFLICT = 409
    HTTP_STATUS_GONE = 410
    HTTP_STATUS_LENGTH_REQUIRED = 411
    HTTP_STATUS_PRECOND_FAILED = 412
    HTTP_STATUS_REQUEST_TOO_LARGE = 413
    HTTP_STATUS_URI_TOO_LONG = 414
    HTTP_STATUS_UNSUPPORTED_MEDIA = 415
    HTTP_STATUS_RETRY_WITH = 449
    HTTP_STATUS_SERVER_ERROR = 500
    HTTP_STATUS_NOT_SUPPORTED = 501
    HTTP_STATUS_BAD_GATEWAY = 502
    HTTP_STATUS_SERVICE_UNAVAIL = 503
    HTTP_STATUS_GATEWAY_TIMEOUT = 504
    HTTP_STATUS_VERSION_NOT_SUP = 505
End Enum

Private Enum eHresult_Error
'Navigation error HRESULT Status Codes (Decimal value)
    INET_E_INVALID_URL = -2146697214
    INET_E_NO_SESSION = -2146697213
    INET_E_CANNOT_CONNECT = -2146697212
    INET_E_RESOURCE_NOT_FOUND = -2146697211
    INET_E_OBJECT_NOT_FOUND = -2146697210
    INET_E_DATA_NOT_AVAILABLE = -2146697209
    INET_E_DOWNLOAD_FAILURE = -2146697208
    INET_E_AUTHENTICATION_REQUIRED = -2146697207
    INET_E_NO_VALID_MEDIA = -2146697206
    INET_E_CONNECTION_TIMEOUT = -2146697205
    INET_E_INVALID_REQUEST = -2146697204
    INET_E_UNKNOWN_PROTOCOL = -2146697203
    INET_E_SECURITY_PROBLEM = -2146697202
    INET_E_CANNOT_LOAD_DATA = -2146697201
    INET_E_CANNOT_INSTANTIATE_OBJECT = -2146697200
    INET_E_REDIRECT_FAILED = -2146697196
    INET_E_REDIRECT_TO_DIR = -2146697195
    INET_E_CANNOT_LOCK_REQUEST = -2146697194
    INET_E_USE_EXTEND_BINDING = -2146697193
    INET_E_TERMINATED_BIND = -2146697192
    INET_E_INVALID_CERTIFICATE = -2146697191
    INET_E_CODE_DOWNLOAD_DECLINED = -2146696960
    INET_E_RESULT_DISPATCHED = -2146696704
    INET_E_CANNOT_REPLACE_SFP_FILE = -2146696448
    INET_E_CODE_INSTALL_BLOCKED_BY_HASH_POLICY = -2146695936
    INET_E_CODE_INSTALL_SUPPRESSED = -2146696192
End Enum

'URL defalut string
Private Const kDefaultURL = "http://www.TradeNavigator.com/"
Private Const kIniSection = "WebURL"

Private Type mPrivate
    strUrl As String
    bShow As Boolean
    oAlert As cAlert
End Type

Private m As mPrivate

Public Sub ShowMe(ByVal strTitle$, Optional ByVal strIcon$, Optional ByVal strUrl$, _
    Optional ByVal strPostData$ = "", Optional Alert As cAlert = Nothing)
On Error GoTo ErrSection:

    Dim strIniUrl$
    Dim strPlacement$
    Dim aPostData() As Byte
    Dim strHeaders$
    
'URL selection order:
'   1. use URL in chartnavigator.ini if it exists (intended for on-demand override for testing/debugging)
'   2. use passed in URL
'   3. use hard-coded defaults based on passed in title
'   4. use kDefaultURL constant
    
    If Len(strTitle) > 0 Then
        Me.Caption = strTitle
        strPlacement = GetIniFileProperty(strTitle, "", "Placement", g.strIniFile)
        
        'check for override in INI file
        strIniUrl = Replace(strTitle, " ", "")
        strIniUrl = GetIniFileProperty(strIniUrl, "", kIniSection, g.strIniFile)
        m.bShow = True
    Else
        ' if no title specified, then this form is not meant to be visible
        Me.Caption = ""
        strPlacement = ""
        m.bShow = False
    End If
    
    If Len(strIniUrl) > 0 Then
        m.strUrl = strIniUrl        'use URL in INI if found
    ElseIf Len(strUrl) > 0 Then
        m.strUrl = strUrl           'use passed in URL
    Else
        m.strUrl = DefaultURLByTitle(strTitle)  'use hard-coded default based on title
    End If
    If Len(m.strUrl) = 0 Then m.strUrl = kDefaultURL    'use hard-coded constant
    
    m.strUrl = FixURL(m.strUrl)
    
    If FileExist(kShowUrlFlagFile) Then StatusMsg m.strUrl
        
    Me.Visible = m.bShow
    webBrowser.Visible = m.bShow
    If m.bShow Then
        If Len(strIcon) > 0 Then
            Me.Icon = Picture16(strIcon)
        Else
            Me.Icon = Picture16("kBlank")
        End If
        'Restore/set form size & location
        If strPlacement = "" Then
            CenterTheForm Me
        Else
            SetFormPlacement Me, strPlacement
        End If
        
        Screen.MousePointer = vbHourglass
    End If
    
    tmrUnload.Enabled = False
    Set m.oAlert = Alert
    
    If IsIDE Then
        FileFromString App.Path & "\chk\Alert.txt", m.strUrl & vbCrLf & strPostData
    End If
    
    If Len(strPostData) = 0 Then
        webBrowser.Navigate2 m.strUrl
    Else
        ' VB creates a Unicode string by default so we need to
        ' convert it back to Single byte character set.
        aPostData = StrConv(strPostData, vbFromUnicode)
        strHeaders = "Content-Type: application/x-www-form-urlencoded" & vbCrLf
        webBrowser.Navigate2 m.strUrl, , , aPostData, strHeaders
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmWebReport.ShowMe"

End Sub

Private Sub Form_Resize()
On Error Resume Next

    If m.bShow Then
        webBrowser.Move 0, 0, ScaleWidth, ScaleHeight
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    tmrUnload.Enabled = False
    If m.bShow Then
        If Len(Me.Caption) > 0 Then
            'save form size & location
            SetIniFileProperty Me.Caption, GetFormPlacement(Me), "Placement", g.strIniFile
        End If
    End If
    Set m.oAlert = Nothing

End Sub

'hard-coded defaults by titles
Private Function DefaultURLByTitle(ByVal strTitle$) As String
On Error GoTo ErrSection:

    Dim strRet$, strUserID$, strPassWd$

    Select Case strTitle
        Case "Danielcode Trade Signals"
            strRet = "http://www.TradeNavigator.com/dc/matrix.aspx"
            'm.bRefresh = True
        Case "Sector Analysis"
            strUserID = EncryptToHex(RI_GetDataServiceID)
            strPassWd = EncryptToHex(RI_GetUserPassword)
            strRet = "www.TradeNavigator.com/industries/indyrank.aspx?" & Chr(38) & "U=" & strUserID & Chr(38) & "P=" & strPassWd
            Me.WindowState = vbMaximized        'JM - 11-06-2009 (temporary until Chad fixes scroll issue)
    End Select

    DefaultURLByTitle = FixURL(strRet)

ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmWebReport.DefaultURLByTitle"

End Function

Private Sub tmrUnload_Timer()

    'webBrowser.Refresh2 ' TLB 4/16/2012: looks like this is causing a 2nd web page request and is not needed
    
    tmrUnload.Enabled = False
    Unload Me
    
End Sub

Private Sub webBrowser_NavigateComplete2(ByVal pDisp As Object, Url As Variant)
On Error GoTo ErrSection:

    If Not m.oAlert Is Nothing And Not webBrowser.Document.body Is Nothing Then
        m.oAlert.WebMailStatus webBrowser.Document.body.innerText
    End If
    
    If m.bShow Then
        Screen.MousePointer = vbDefault
        ShowForm Me, eForm_Nonmodal
    Else
        tmrUnload.Enabled = True
    End If

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmWebReport.webBrowser_NavigateComplete2"
End Sub

Private Function ErrorToString(ByVal StatusCode&) As String
On Error GoTo ErrSection:

    Dim strEnum$, strEnglish$

    strEnum = Str(StatusCode)
    strEnglish = "Unknown error code"

    Select Case StatusCode
        Case HTTP_STATUS_BAD_REQUEST
            strEnum = "HTTP_STATUS_BAD_REQUEST"
            strEnglish = "The request could not be processed by the server due to invalid syntax."
        Case HTTP_STATUS_DENIED
            strEnum = "HTTP_STATUS_DENIED"
            strEnglish = "The requested resource requires user authentication."
        Case HTTP_STATUS_PAYMENT_REQ
            strEnum = "HTTP_STATUS_PAYMENT_REQ"
            strEnglish = "Not currently implemented in the HTTP protocol."
        Case HTTP_STATUS_FORBIDDEN
            strEnum = "HTTP_STATUS_FORBIDDEN"
            strEnglish = "The server understood the request, but is refusing to fulfill it."
        Case HTTP_STATUS_NOT_FOUND
            strEnum = "HTTP_STATUS_NOT_FOUND"
            strEnglish = "The server has not found anything matching the requested URI (Uniform Resource Identifier)."
        Case HTTP_STATUS_BAD_METHOD
            strEnum = "HTTP_STATUS_BAD_METHOD"
            strEnglish = "The HTTP verb used is not allowed."
        Case HTTP_STATUS_PROXY_AUTH_REQ
            strEnum = "HTTP_STATUS_PROXY_AUTH_REQ"
            strEnglish = "Proxy authentication required."
        Case HTTP_STATUS_REQUEST_TIMEOUT
            strEnum = "HTTP_STATUS_REQUEST_TIMEOUT"
            strEnglish = "The server timed out waiting for the request."
        Case HTTP_STATUS_CONFLICT
            strEnum = "HTTP_STATUS_CONFLICT"
            strEnglish = "The request could not be completed due to a conflict with the current state of the resource. The user should resubmit with more information."
        Case HTTP_STATUS_GONE
            strEnum = "HTTP_STATUS_GONE"
            strEnglish = "The requested resource is no longer available at the server, and no forwarding address is known."
        Case HTTP_STATUS_LENGTH_REQUIRED
            strEnum = "HTTP_STATUS_LENGTH_REQUIRED"
            strEnglish = "The server refuses to accept the request without a defined content length."
        Case HTTP_STATUS_PRECOND_FAILED
            strEnum = "HTTP_STATUS_PRECOND_FAILED"
            strEnglish = "The precondition given in one or more of the request header fields evaluated to false when it was tested on the server."
        Case HTTP_STATUS_REQUEST_TOO_LARGE
            strEnum = "HTTP_STATUS_REQUEST_TOO_LARGE"
            strEnglish = "The server is refusing to process a request because the request entity is larger than the server is willing or able to process."
        Case HTTP_STATUS_URI_TOO_LONG
            strEnum = "HTTP_STATUS_URI_TOO_LONG"
            strEnglish = "The server is refusing to service the request because the request URI (Uniform Resource Identifier) is longer than the server is willing to interpret."
        Case HTTP_STATUS_UNSUPPORTED_MEDIA
            strEnum = "HTTP_STATUS_UNSUPPORTED_MEDIA"
            strEnglish = "The server is refusing to service the request because the entity of the request is in a format not supported by the requested resource for the requested method."
        Case HTTP_STATUS_RETRY_WITH
            strEnum = "HTTP_STATUS_RETRY_WITH"
            strEnglish = "The request should be retried after doing the appropriate action."
        Case HTTP_STATUS_SERVER_ERROR
            strEnum = "HTTP_STATUS_SERVER_ERROR"
            strEnglish = "The server encountered an unexpected condition that prevented it from fulfilling the request."
        Case HTTP_STATUS_NOT_SUPPORTED
            strEnum = "HTTP_STATUS_NOT_SUPPORTED"
            strEnglish = "The server does not support the functionality required to fulfill the request."
        Case HTTP_STATUS_BAD_GATEWAY
            strEnum = "HTTP_STATUS_BAD_GATEWAY"
            strEnglish = "The server, while acting as a gateway or proxy, received an invalid response from the upstream server it accessed in attempting to fulfill the request."
        Case HTTP_STATUS_SERVICE_UNAVAIL
            strEnum = "HTTP_STATUS_SERVICE_UNAVAIL"
            strEnglish = "The service is temporarily overloaded."
        Case HTTP_STATUS_GATEWAY_TIMEOUT
            strEnum = "HTTP_STATUS_GATEWAY_TIMEOUT"
            strEnglish = "The request was timed out waiting for a gateway."
        Case HTTP_STATUS_VERSION_NOT_SUP
            strEnum = "HTTP_STATUS_VERSION_NOT_SUP"
            strEnglish = "The server does not support, or refuses to support, the HTTP protocol version that was used in the request message."
    
        Case INET_E_INVALID_URL
           strEnum = "INET_E_INVALID_URL"
           strEnglish = "The URL could not be parsed."
        Case INET_E_NO_SESSION
           strEnum = "INET_E_NO_SESSION"
           strEnglish = "No Internet session was established."
        Case INET_E_CANNOT_CONNECT
           strEnum = "INET_E_CANNOT_CONNECT"
           strEnglish = "The attempt to connect to the Internet has failed."
        Case INET_E_RESOURCE_NOT_FOUND
           strEnum = "INET_E_RESOURCE_NOT_FOUND"
           strEnglish = "The server or proxy was not found."
        Case INET_E_OBJECT_NOT_FOUND
           strEnum = "INET_E_OBJECT_NOT_FOUND"
           strEnglish = "The object was not found."
        Case INET_E_DATA_NOT_AVAILABLE
           strEnum = "INET_E_DATA_NOT_AVAILABLE"
           strEnglish = "An Internet connection was established, but the data cannot be retrieved."
        Case INET_E_DOWNLOAD_FAILURE
           strEnum = "INET_E_DOWNLOAD_FAILURE"
           strEnglish = "The download has failed (the connection was interrupted)."
        Case INET_E_AUTHENTICATION_REQUIRED
           strEnum = "INET_E_AUTHENTICATION_REQUIRED"
           strEnglish = "Authentication is needed to access the object."
        Case INET_E_NO_VALID_MEDIA
           strEnum = "INET_E_NO_VALID_MEDIA"
           strEnglish = "The object is not in one of the acceptable MIME types."
        Case INET_E_CONNECTION_TIMEOUT
           strEnum = "INET_E_CONNECTION_TIMEOUT"
           strEnglish = "The Internet connection has timed out."
        Case INET_E_INVALID_REQUEST
           strEnum = "INET_E_INVALID_REQUEST"
           strEnglish = "The request was invalid."
        Case INET_E_UNKNOWN_PROTOCOL
           strEnum = "INET_E_UNKNOWN_PROTOCOL"
           strEnglish = "The protocol is not known and no pluggable protocols have been entered that match."
        Case INET_E_SECURITY_PROBLEM
           strEnum = "INET_E_SECURITY_PROBLEM"
           strEnglish = "A security problem was encountered."      'note there is a win32 err code with this (check MSDN)
        Case INET_E_CANNOT_LOAD_DATA
           strEnum = "INET_E_CANNOT_LOAD_DATA"
           strEnglish = "The object could not be loaded."
        Case INET_E_CANNOT_INSTANTIATE_OBJECT
           strEnum = "INET_E_CANNOT_INSTANTIATE_OBJECT"
           strEnglish = "CoCreateInstance failed."
        Case INET_E_REDIRECT_FAILED
           strEnum = "INET_E_REDIRECT_FAILED"
           strEnglish = "Microsoft Win32 Internet (WinInet) cannot redirect. This error code might also be returned by a custom protocol handler."
        Case INET_E_REDIRECT_TO_DIR
           strEnum = "INET_E_REDIRECT_TO_DIR"
           strEnglish = "The request is being redirected to a directory."
        Case INET_E_CANNOT_LOCK_REQUEST
           strEnum = "INET_E_CANNOT_LOCK_REQUEST"
           strEnglish = "The requested resource could not be locked."
        Case INET_E_USE_EXTEND_BINDING
           strEnum = "INET_E_USE_EXTEND_BINDING"
           strEnglish = "(Microsoft internal.) Reissue request with extended binding"
        Case INET_E_TERMINATED_BIND
           strEnum = "INET_E_TERMINATED_BIND"
           strEnglish = "Binding was terminated. (See IBinding::GetBindResult.)."
        Case INET_E_INVALID_CERTIFICATE
           strEnum = "INET_E_INVALID_CERTIFICATE"
           strEnglish = "The Secure Sockets Layer (SSL) certificate is invalid. "
        Case INET_E_CODE_DOWNLOAD_DECLINED
           strEnum = "INET_E_CODE_DOWNLOAD_DECLINED"
           strEnglish = "The component download was declined by the user."
        Case INET_E_RESULT_DISPATCHED
           strEnum = "INET_E_RESULT_DISPATCHED"
           strEnglish = "The binding has already been completed and the result has been dispatched, so your abort call has been canceled."
        Case INET_E_CANNOT_REPLACE_SFP_FILE
           strEnum = "INET_E_CANNOT_REPLACE_SFP_FILE"
           strEnglish = "Cannot replace a file that is protected by System File Protection (SFP)."
        Case INET_E_CODE_INSTALL_BLOCKED_BY_HASH_POLICY
           strEnum = "INET_E_CODE_INSTALL_BLOCKED_BY_HASH_POLICY"
        Case INET_E_CODE_INSTALL_SUPPRESSED
           strEnum = "INET_E_CODE_INSTALL_SUPPRESSED"
    End Select
    
    ErrorToString = strEnum & vbCrLf & strEnglish
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "frmWebReport.ErrorToString"

End Function

Private Sub webBrowser_NavigateError(ByVal pDisp As Object, Url As Variant, frame As Variant, StatusCode As Variant, Cancel As Boolean)
On Error GoTo ErrSection:
    
    Dim strErr$
    
    strErr = ErrorToString(StatusCode)
    If m.bShow Then
        InfBox "Web navigation error " & strErr, "I", , "Web Navigation Error"
    Else
        If Not m.oAlert Is Nothing Then
            m.oAlert.WebMailStatus strErr
        End If
        tmrUnload.Enabled = True
    End If
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "frmWebReport.webBrowser_NavigateError"

End Sub

