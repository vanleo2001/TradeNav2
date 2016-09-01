Attribute VB_Name = "mRegEdit"
Option Explicit

Enum RKeyRoots
    rkClassesRoot = 1&
    rkCurrentConfig = 2&
    rkCurrentUser = 3&
    rkLocalMachine = 4&
    rkUsers = 5&
End Enum

Private Declare Function RK_Create Lib "G32_RKEY.dll" (ByVal root As RKeyRoots, ByVal gdStrKeyName&) As Long
Private Declare Function RK_Open Lib "G32_RKEY.dll" (ByVal root As RKeyRoots, ByVal gdStrKeyName&) As Long
Private Declare Function RK_OpenReadOnly Lib "G32_RKEY.dll" (ByVal root As RKeyRoots, ByVal gdStrKeyName&) As Long
Private Declare Function RK_Close Lib "G32_RKEY.dll" () As Long

Private Declare Function RK_KeyValue Lib "G32_RKEY.dll" (ByVal gdStrKeyName&, ByVal gdStrKeyValue&) As Long
Private Declare Function RK_StrValue Lib "G32_RKEY.dll" (ByVal gdStrValueName&, ByVal gdStrValue&) As Long
Private Declare Function RK_LongValue Lib "G32_RKEY.dll" (ByVal gdStrValueName&, nLongValue&) As Long
Private Declare Function RK_BinaryValue Lib "G32_RKEY.dll" (ByVal gdStrValueName&, pMemory As Any, ByVal nMaxLen&, nDataLen&) As Long

Private Declare Function RK_SetKeyValue Lib "G32_RKEY.dll" (ByVal gdStrKeyName&, ByVal gdStrKeyValue&) As Long
Private Declare Function RK_SetStrValue Lib "G32_RKEY.dll" (ByVal gdStrValueName&, ByVal gdStrValue&) As Long
Private Declare Function RK_SetLongValue Lib "G32_RKEY.dll" (ByVal gdStrValueName&, ByVal nLongValue&) As Long
Private Declare Function RK_SetBinaryValue Lib "G32_RKEY.dll" (ByVal gdStrValueName&, pMemory As Any, ByVal nDataLen&) As Long

Private Declare Function RK_Delete Lib "G32_RKEY.dll" (ByVal strValueName$) As Long
Private Declare Function RK_RecurseDelete Lib "G32_RKEY.dll" (ByVal strKeyName$) As Long

Private Declare Function RK_EnumerateKeys Lib "G32_RKEY.dll" (ByVal Index&, ByVal gdStrKeyName&) As Long

' To SET a value in the registry ...
' - rkRoot: one of enumerated registry roots
' - strKeyName: name of the key (will create it if does not exist)
' - strValueName: name of the value (if blank, will be default value for key)
' - vValue: value to store in the registry (will be stored based on data type)
' - bBinary (optional): set to true if string to be stored as binary
' - returns True if stored successfully, False if not
' EXAMPLE:  SetRegistryValue rkLocalMachine, "Software\Genesis\General", _
'                   "Number of Days", nNumDays
Public Function SetRegistryValue(ByVal rkRoot As RKeyRoots, _
        ByVal strKeyName$, ByVal strValueName$, ByVal vValue As Variant, _
        Optional ByVal bAsBinary As Boolean = False) As Boolean

    Dim gdsKeyName As New cGdArray, gdsValueName As New cGdArray, gdsValue As New cGdArray
    Dim nValue&, dValue#, strValue$, bSuccess As Boolean

    ' init gdStrings
    gdsKeyName.Create eGDARRAY_gdString
    gdsValueName.Create eGDARRAY_gdString
    gdsValue.Create eGDARRAY_gdString
    
    ' open key
    gdsKeyName(0) = FixKeyName(strKeyName)
    If RK_Open(rkRoot, gdsKeyName.ArrayHandle) <> 0 Then
        ' if not exist, then create the key
        If RK_Create(rkRoot, gdsKeyName.ArrayHandle) <> 0 Then
            ' error creating key
            Exit Function
        End If
    End If
    
    ' set value for item (based on data type)
    gdsValueName(0) = Trim(strValueName)
    Select Case VarType(vValue)
        Case vbLong, vbInteger, vbBoolean:
            If Len(gdsValueName(0)) = 0 Then
                gdsValue(0) = Trim(Str(vValue))
                If RK_SetKeyValue(gdsValueName.ArrayHandle, gdsValue.ArrayHandle) = 0 Then bSuccess = True
            Else
                nValue = vValue
                If RK_SetLongValue(gdsValueName.ArrayHandle, nValue) = 0 Then bSuccess = True
            End If
        'handle dates as doubles (NOT as strings since format can vary)
        Case vbDouble, vbSingle, vbDate:
            dValue = CDbl(vValue)
            gdsValue(0) = Trim(Str(dValue))
            If Len(gdsValueName(0)) = 0 Then
                If RK_SetKeyValue(gdsValueName.ArrayHandle, gdsValue.ArrayHandle) = 0 Then bSuccess = True
            Else
                If RK_SetStrValue(gdsValueName.ArrayHandle, gdsValue.ArrayHandle) = 0 Then bSuccess = True
            End If
        Case Else:
            If Len(gdsValueName(0)) = 0 Then
                gdsValue(0) = CStr(vValue)
                If RK_SetKeyValue(gdsValueName.ArrayHandle, gdsValue.ArrayHandle) = 0 Then bSuccess = True
            ElseIf bAsBinary Then
                strValue = CStr(vValue)
                If RK_SetBinaryValue(gdsValueName.ArrayHandle, ByVal strValue, Len(strValue)) = 0 Then bSuccess = True
            Else
                gdsValue(0) = CStr(vValue)
                If RK_SetStrValue(gdsValueName.ArrayHandle, gdsValue.ArrayHandle) = 0 Then bSuccess = True
            End If
    End Select

    RK_Close
    SetRegistryValue = bSuccess
End Function

' To GET a value in the registry ...
' - rkRoot: one of enumerated registry roots
' - strKeyName: name of the key
' - strValueName: name of the value (if blank, will get the default value for key)
' - vDefaultValue: this will be returned if key or value does not exist (will be retrieved based on this data type)
' - nMaxBytesIfBinary: pass max bytes > 0 if need to retrieve binary data (as vb string)
' - returns the value in the registry (will be same data type as vDefaultValue)
' EXAMPLE:  nNumDays = GetRegistryValue(rkLocalMachine, "Software\Genesis\General", _
'                       "Number of Days", 30)
Public Function GetRegistryValue(ByVal rkRoot As RKeyRoots, _
        ByVal strKeyName$, ByVal strValueName$, ByVal vDefaultValue As Variant, _
        Optional ByVal nMaxBytesIfBinary = 0) As Variant

    Dim gdsKeyName As New cGdArray, gdsValueName As New cGdArray, gdsValue As New cGdArray
    Dim vValue As Variant, nValue&, strValue$, rc&

    ' init gdStrings
    gdsKeyName.Create eGDARRAY_gdString
    gdsValueName.Create eGDARRAY_gdString
    gdsValue.Create eGDARRAY_gdString
    gdsValue(0) = ""
    
    ' set return value to default
    vValue = vDefaultValue
    
    ' open the key
    gdsKeyName(0) = FixKeyName(strKeyName)
    If RK_OpenReadOnly(rkRoot, gdsKeyName.ArrayHandle) = 0 Then
        gdsValueName(0) = strValueName
        ' get value of item (based on data type of default)
        Select Case VarType(vDefaultValue)
            Case vbLong, vbInteger, vbBoolean:
                If Len(gdsValueName(0)) = 0 Then
                    rc = RK_KeyValue(gdsValueName.ArrayHandle, gdsValue.ArrayHandle)
                    nValue = Val(gdsValue(0))
                Else
                    rc = RK_LongValue(gdsValueName.ArrayHandle, nValue)
                End If
                If rc = 0 Then
                    If VarType(vDefaultValue) = vbBoolean Then
                        'special handling for booleans in case registry
                        'value set by "C" code (where "true" = +1):
                        If nValue = 0 Then
                            vValue = False
                        Else
                            vValue = True
                        End If
                    Else
                        vValue = nValue
                    End If
                End If
            'dates handled as doubles (NOT as strings since format can vary)
            Case vbDouble, vbSingle, vbDate:
                If Len(gdsValueName(0)) = 0 Then
                    rc = RK_KeyValue(gdsValueName.ArrayHandle, gdsValue.ArrayHandle)
                Else
                    rc = RK_StrValue(gdsValueName.ArrayHandle, gdsValue.ArrayHandle)
                End If
                If rc = 0 Then
                    If VarType(vDefaultValue) = vbDate Then
                        vValue = CDate(Val(gdsValue(0)))
                    Else
                        vValue = Val(gdsValue(0))
                    End If
                End If
            Case Else:
                If Len(gdsValueName(0)) = 0 Then
                    If RK_KeyValue(gdsValueName.ArrayHandle, gdsValue.ArrayHandle) = 0 Then
                        vValue = gdsValue(0)
                    End If
                ElseIf nMaxBytesIfBinary > 0 Then
                    strValue = Space(nMaxBytesIfBinary + 1)
                    If RK_BinaryValue(gdsValueName.ArrayHandle, ByVal strValue, nMaxBytesIfBinary, nValue) = 0 Then
                        vValue = Left(strValue, nValue)
                    End If
                Else
                    If RK_StrValue(gdsValueName.ArrayHandle, gdsValue.ArrayHandle) = 0 Then
                        vValue = gdsValue(0)
                    End If
                End If
        End Select
        
        RK_Close
    End If

    GetRegistryValue = vValue
End Function

' fix key name (strip backslashes on ends, etc.)
Private Function FixKeyName(ByVal strKey$) As String
    strKey = Trim(strKey)
    If Left(strKey, 1) = "\" Then strKey = Mid(strKey, 2)
    If Right(strKey, 1) = "\" Then strKey = Left(strKey, Len(strKey) - 1)
    FixKeyName = strKey
End Function

' to delete this value from the registry
Public Function DeleteRegistryValue(ByVal rkRoot As RKeyRoots, _
        ByVal strKeyName$, ByVal strValueName$) As Boolean
        
    Dim gdsKeyName As New cGdArray
    Dim bSuccess As Boolean

    ' open key
    gdsKeyName.Create eGDARRAY_gdString
    gdsKeyName(0) = FixKeyName(strKeyName)
    If RK_Open(rkRoot, gdsKeyName.ArrayHandle) = 0 Then
        ' delete the specified value from the key
        If RK_Delete(strValueName) = 0 Then
            bSuccess = True
        End If
        RK_Close
    End If

    DeleteRegistryValue = bSuccess
End Function
