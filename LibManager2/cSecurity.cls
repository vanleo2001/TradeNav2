VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
'=============================================================================
'SECURITY/PERMISSIONS (for Functions, Rules, Systems, and Money management).
'
'Security Table
' Level:  a-Remove|b-List|c-Preview|d-Edit|e-Save|
'   0         Y   |  Y   |   Y     |  Y   |  Y   |
'   1         P   |  Y   |   Y     |  Y   |  P   |
'   2         P   |  Y   |   N     |  P   |  P   |
'   3         N   |  N   |   N     |  N   |  N   |
'
' P - Means Password is required
'
' a - Item can be removed when "Delete" button is pressed
' b - Item can be shown in menu lists
' c - Item can be shown in preview list on menu form
' d - Item can be shown in edit form (frmRule, frmMM, etc.)
' e - Item can be Save on edit form (Save button)
'=============================================================================

Private Type mPrivate
    strPassword As String
    bCancel As Boolean
End Type
Private m As mPrivate

'Security modes
Private Const C_EDITANDVIEW = 0
Private Const C_VIEWONLY = 1
Private Const C_NOVIEW = 2
Private Const C_NOLIST = 3

'Returned from Password form
Property Let Password(pData As String)
    m.strPassword = pData
End Property
Property Let Cancel(pData As Boolean)
    m.bCancel = pData
End Property

'Authorized to remove item
' pItemDesc = "Function","Strategy","Rule","System"
Public Function CanRemove(pItemDesc As String, pSecurityLevel As Byte, _
                        pPassword As String, pCannotDelete As Boolean, _
                        Optional ByVal strItem As String = "") As Boolean
On Error GoTo ErrSection:

    Dim strPassword As String

    CanRemove = False
    
    'Item cannot be deleted
    If pCannotDelete Then
        Err.Raise vbObjectError + 1000, , "This " & pItemDesc & " cannot be removed.  " & _
                "It was marked as 'protected' by the developer."
    End If
    
    Select Case pSecurityLevel
        Case C_EDITANDVIEW: CanRemove = True
        Case C_VIEWONLY, C_NOVIEW
            CanRemove = ValidPassword(pPassword, strItem)
        
    End Select

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cSecurity.CanRemove", eGDRaiseError_Raise, g.strAppPath
    
End Function

'Authorized to Preview
Public Function CanPreview(pSecurityLevel As Byte) As Boolean
On Error GoTo ErrSection:
    
    CanPreview = False
    Select Case pSecurityLevel
        Case C_EDITANDVIEW: CanPreview = True
        Case C_VIEWONLY: CanPreview = True
        Case C_NOVIEW: CanPreview = False
    End Select

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cSecurity.CanPreview", eGDRaiseError_Raise, g.strAppPath
    
End Function

'Authorized to Edit
Public Function CanEdit(pSecurityLevel As Byte, pPassword As String, Optional ByVal strItem$ = "") As Boolean
On Error GoTo ErrSection:

    Dim strPassword As String

    CanEdit = False
    Select Case pSecurityLevel
        Case C_EDITANDVIEW: CanEdit = True
        Case C_VIEWONLY: CanEdit = True
        Case C_NOVIEW
            CanEdit = ValidPassword(pPassword, strItem)
        
    End Select

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cSecurity.CanEdit", eGDRaiseError_Raise, g.strAppPath

End Function

'Authorized to Save
Public Function CanSave(pSecurityLevel As Byte, pPassword As String, Optional ByVal strItem$ = "") As Boolean
On Error GoTo ErrSection:
    
    Dim strIniPassword As String
    Dim strPassword As String
    
    CanSave = False
    Select Case pSecurityLevel
        Case C_EDITANDVIEW: CanSave = True
        Case C_VIEWONLY, C_NOVIEW
            CanSave = ValidPassword(pPassword, strItem)
        
    End Select

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cSecurity.CanSave", eGDRaiseError_Raise, g.strAppPath

End Function

Private Function ValidPassword(ByVal strValid As String, ByVal strItem As String) As Boolean

    Static strGoodPasswords As String
    Dim strPassword As String

    strValid = Trim(UCase(strValid))

    If InStr(strGoodPasswords, vbTab & strValid & vbTab) Then
        ValidPassword = True
    Else
        strPassword = Trim(UCase(frmPassword.ShowMe(strItem)))
        If strPassword <> "" Then
            If strPassword <> strValid Then
                InfBox "Password is not correct", "!", , "Password"
            Else
                If Len(strGoodPasswords) = 0 Then strGoodPasswords = vbTab
                strGoodPasswords = strGoodPasswords & strPassword & vbTab
                ValidPassword = True
            End If
        End If
    End If

End Function

