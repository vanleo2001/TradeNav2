Attribute VB_Name = "mDBSecurity"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        mDBSecurity.bas
'' Description: Routines for database security and encryption
''
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FieldEncryptKey
'' Description: Return the Key to encrypt/decrypt database fields
'' Inputs:      None
'' Returns:     Key to use to Encrypt/Decrypt
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FieldEncryptKey() As cMemBuffer
On Error Resume Next

    Static mbKey As New cMemBuffer
   
    If mbKey.Length = 0 Then
        mbKey.PutByte 71
        mbKey.PutByte 202
        mbKey.PutByte 123
        mbKey.PutByte 63
        mbKey.PutByte 176
        mbKey.PutByte 2
        mbKey.PutByte 70
        mbKey.PutByte 198
        mbKey.PutByte 169
        mbKey.PutByte 85
        mbKey.PutByte 10
    End If
        
    Set FieldEncryptKey = mbKey

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EncryptField
'' Description: Encrypt the given string with a special key
'' Inputs:      String to Encrypt
'' Returns:     Encrypted String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub EncryptField(fld As Field, ByVal strToEncrypt As String)
On Error GoTo ErrSection:

    Dim mb As cMemBuffer
    
    If Len(strToEncrypt) = 0 Then
        fld.Value = ""
    Else
        Set mb = New cMemBuffer
        mb.Buffer = strToEncrypt
        gdEncrypt True, mb, FieldEncryptKey
        ' All encrypted fields should now be Binary, but
        ' check for Text/Memo for backwards-compatibility
        If fld.Type = dbText Or fld.Type = dbMemo Then
            fld.Value = mb.Buffer
        Else
            fld.Value = mb.Bytes
        End If
        Set mb = Nothing
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mDBSecurity.EncryptField", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DecryptField
'' Description: Decrypt the string from the field with a special key
'' Inputs:      Field to Decrypt
'' Returns:     Decrypted String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DecryptField(fld As Field) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Encrypted string to return
    Dim mb As cMemBuffer
    
    Set mb = GetBinaryData(fld)
    gdEncrypt False, mb, FieldEncryptKey
    DecryptField = mb.Buffer

ErrExit:
    Set mb = Nothing
    Exit Function
    
ErrSection:
    RaiseError "mDBSecurity.DecryptField", eGDRaiseError_Raise, g.strAppPath
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BuildCheckSum
'' Description: Build a check sum on the current record of the recordset
'' Inputs:      Recordset, Table ID
'' Returns:     Check Sum for the current record in the recordset
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function BuildCheckSum(rsRecord As Recordset, ByVal strTableID As String, _
                    Optional ByVal lCID As Long = -1, _
                    Optional lCheckSumField As Long = -1) As Double
On Error GoTo ErrSection:

    Dim strTemp As String               ' Temporary string to work with
    Dim lIndex As Long                  ' Index into a for loop
    Dim mb As New cMemBuffer, mbFld As New cMemBuffer

'    If Not (rsRecord.BOF And rsRecord.EOF) Then
        ' Walk through the fields and append them to one giant string
        For lIndex = 0 To rsRecord.Fields.Count - 1
            With rsRecord.Fields(lIndex)
                If .SourceTable = strTableID Then
                    If .SourceField = "CheckSum" Then
                        lCheckSumField = lIndex
                    Else
                        'NOTE: we MUST use "Str" instead of "CStr" here since we must use
                        'a decimal point instead of a comma regardless of regional settings.
                        Select Case .Type
                            Case dbDate, dbBoolean
                                mb.PutStr Str(CDbl(NullChk(.Value, 0)))
                            
                            Case dbMemo, dbText
                                mb.PutStr NullChk(.Value)
                                
                            Case dbBinary, dbLongBinary
                                Set mbFld = GetBinaryData(rsRecord.Fields(lIndex))
                                mb.PutFromMemory mbFld.MemPtr, mbFld.Length
                                
                            Case Else
                                mb.PutStr Trim(Str(NullChk(.Value, 0)))
                        End Select
                    End If
                End If
            End With
        Next lIndex
            
        If lCheckSumField >= 0 Then
            ' For the Library table, we also want to include the last known Customer ID
            If strTableID = "tblLibrarys" Then
                If lCID = -1 Then
                    mb.PutStr Str(g.lLCD)
                Else
                    mb.PutStr Str(lCID)
                End If
            End If
            
            ' Encrypt the string
            gdEncrypt True, mb, FieldEncryptKey
            
            ' Calculate a checksum value on the encrypted string and return it
            BuildCheckSum = gdCalcMemCRC32(mb.MemPtr, mb.Length)
        Else
            Err.Raise vbObjectError + 1000, , "Database has become invalid"
        End If
'    End If
    
ErrExit:
    Set mb = Nothing
    Set mbFld = Nothing
    Exit Function
    
ErrSection:
    RaiseError "mDBSecurity.BuildCheckSum", eGDRaiseError_Raise, g.strAppPath
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ValidateCheckSums
'' Description: Walk through a recordset and validate that the check sum in
''              each record matches what it should be
'' Inputs:      Recordset, Table ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ValidateCheckSums(rsRecord As Recordset, ByVal strTableID As String, _
                Optional ByVal lCID As Long = -1)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim dCheckSum As Double             ' Check Sum from the recordset
    Dim lCheckSumField As Long          ' Check Sum field number in recordset
    Dim lPosition As Long               ' Current position in the recordset
    Dim strPW As String

    With rsRecord
        If Not (.BOF And .EOF) Then
            lPosition = .AbsolutePosition
            .MoveFirst
            Do While Not .EOF
If 0 And IsIDE Then
    ' for automatically changing all our passwords
    On Error Resume Next
    strPW = LCase(DecryptField(.Fields("Password")))
    On Error GoTo ErrSection:
    If strPW = "genesis1" Then
        strPW = "genesis123"
    ElseIf strPW = "larryw1" Then
        strPW = "larry123"
    Else
        strPW = ""
    End If
    If Len(strPW) > 0 Then
        .Edit
        EncryptField .Fields("Password"), strPW
        .Update
    End If
End If
                dCheckSum = BuildCheckSum(rsRecord, strTableID, lCID, lCheckSumField)
                If lCheckSumField < 0 Then
                    Err.Raise vbObjectError + 1000, , "Database has become invalid"
                ElseIf dCheckSum <> .Fields(lCheckSumField).Value Then
                    .Edit
If 0 And IsIDE Then
                    ' only used during development to reset all the checksums
                    .Fields(lCheckSumField).Value = dCheckSum
Else
                    ' set checksum to 0.5 to indicate that it's invalid
                    .Fields(lCheckSumField).Value = 0.5
End If
                    .Update
                End If
                .MoveNext
            Loop
            .AbsolutePosition = lPosition
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mDBSecurity.ValidateCheckSums", eGDRaiseError_Raise, g.strAppPath
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    IsValidCheckSum
'' Description: Determine whether this one record is valid or not
'' Inputs:      ID, ID Field Name, Table
'' Returns:     True if CheckSum's match, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsValidCheckSum(ByVal lID As Long, ByVal strIDField As String, ByVal strTable As String) As Boolean
On Error GoTo ErrSection:

    Dim rs As Recordset
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM " & strTable & _
                "WHERE " & strIDField & "=" & CStr(lID) & ";", dbOpenDynaset, dbDenyWrite)
    If Not rs.EOF Then
        If rs!CheckSum <> BuildCheckSum(rs, strTable) Then
            rs.Edit
            rs!CheckSum = 0.5
            rs.Update
        Else
            IsValidCheckSum = True
        End If
    End If
    rs.Close

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDBSecurity.IsValidCheckSum", eGDRaiseError_Raise, g.strAppPath
    
End Function

Private Function GetBinaryData(fld As Field) As cMemBuffer

    Dim i&
    Dim strReturn As String
    Dim strBytes As String
    Dim mb As New cMemBuffer
    
    If fld.Type = dbMemo Or fld.Type = dbLongBinary Then
        i = fld.FieldSize
    ElseIf Not IsNull(fld) Then
        i = fld.Size
    End If
    If i > 0 Then
        ' All encrypted fields should now be Binary, but
        ' check for Text/Memo for backwards-compatibility
        If fld.Type = dbText Or fld.Type = dbMemo Then
            mb.Buffer = fld.Value
        '    strReturn = CStr(fld.Value)
        'ElseIf IsDBCS Then
        '    strBytes = CStr(fld.Value)
        '    strReturn = Space(LenB(strBytes))
        '    For i = 1 To LenB(strBytes)
        '        Mid(strReturn, i, 1) = ChrW(AscB(MidB(strBytes, i, 1)))
        '    Next
        Else
            mb.Bytes = fld.Value
            'strReturn = mb.Buffer
        End If
    End If
    Set GetBinaryData = mb

End Function

' To fix required module (checks if was encrypted from older versions)
Public Function FixRequiredMod(ByVal strMod As String) As String
On Error GoTo ErrSection:

    Dim i&, iAsc&
    Dim mb As cMemBuffer
    
    If Len(strMod) > 0 Then
        ' check if encrypted -- e.g. if came from an older version GLB
        For i = 1 To Len(strMod)
            iAsc = Asc(Mid(strMod, i, 1))
            If iAsc < 32 Or iAsc > 126 Then
                ' decrypt it
                Set mb = New cMemBuffer
                mb.Buffer = strMod
                gdEncrypt False, mb, FieldEncryptKey
                strMod = mb.Buffer
                Set mb = Nothing
                Exit For
            End If
        Next
        FixRequiredMod = UCase(Trim(strMod))
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mDBSecurity.FixRequiredMod", eGDRaiseError_Raise, g.strAppPath
End Function
