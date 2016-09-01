Attribute VB_Name = "mCodeBaseUtil"
Option Explicit
'CodeBase NOTES (for efficiency, etc.):
' - You MUST do a "Cb4Finish" prior to ending your app.
'     ("Cb4Start" is automatically called the first time
'     you open or create a database;  it does not hurt to
'     call "Cb4Finish" if Cb4Start has never been called).
' - After opening databases, do a "Cb4Optimize True".
' - Declare a long pointer for each field, and assign them
'     once ... e.g. fptr& = d4field(tbl_ptr, "fld_name")
' - Use field functions to   ASSIGN    and     RETRIEVE
'     Char/Memo:         f4putStr(fptr, s$)   f4str(fptr)
'     Int/Long/Logical:  f4putNum(fptr, n&)   f4num(fptr)
'     Single/Double:     f4putFloat(fptr, n#) f4float(fptr)
'   (numeric field functions will automatically assign or
'   retrieve binary data if referencing a character field!)
' - Use the "TagScan" function to loop through each record.
' - Use the "VList..." functions to "marry" a database
'     with a virtual list box (Crescent tool).


' Global variables/flags (for internal use only)
Global cb4Ptr As Long   ' CodeBase pointer for application

' Custom CodeBase calls
'Declare Function tbl_create& Lib "gen_cb6.dll" (ByVal cb4Ptr&, ByVal filename$, ByVal field_str$)
'Declare Function index_create& Lib "gen_cb6.dll" (ByVal cb4Ptr&, ByVal d4Ptr&, ByVal filename$, ByVal tag_str$)

'Declare Function f4num& Lib "gen_cb6.dll" (ByVal f4ptr&)
'Declare Sub f4putNum Lib "gen_cb6.dll" (ByVal f4ptr&, ByVal num&)
'Declare Sub f4getFloat Lib "gen_cb6.dll" (ByVal f4ptr&, rtrn#)
'Declare Sub f4putFloat Lib "gen_cb6.dll" (ByVal f4ptr&, ByVal num#)

'Declare Function d4seekN% Lib "gen_cb6.dll" (ByVal d4&, ByVal seekValue$, ByVal strlen%)
'Declare Function d4seekNextN% Lib "gen_cb6.dll" (ByVal d4&, ByVal seekValue$, ByVal strlen%)


' NOTE: this routine MUST be called before ending your app.
Public Sub Cb4Finish()

    Dim rc&, wait#
    wait = 0.1

    If cb4Ptr <> 0 Then

        ' close all files and optimization down
        Cb4Optimize False

        'Sleep wait

        ' clear CodeBase pointer
        rc = code4initUndo(cb4Ptr)
        If rc Then MsgBox "CB exit error = " + Str(rc)
        If rc = r4success Then cb4Ptr = 0
        'DoEvents
    End If

End Sub

' This should be called after the databases are open
' (uses 1 meg memory, but runs much more efficient).
' Pass a zero to turn it off.
Public Sub Cb4Optimize(Optional ByVal NumMegs% = 1)

    Dim rc&
    Static bOptimizeOn As Boolean

    Cb4Start    ' if not already!

    If NumMegs <> 0 Then
        NumMegs = Abs(NumMegs)
        If (Not bOptimizeOn) And NumMegs < 2000 Then
            rc = code4memStartMax(cb4Ptr, CLng(NumMegs) * 1024 * 1024)
            If code4optStart(cb4Ptr) = r4success Then
                bOptimizeOn = True
            End If
        End If
    Else
        If bOptimizeOn Then
            rc = code4optSuspend(cb4Ptr)
            If rc = r4success Then
                bOptimizeOn = False
            Else
                If rc < 0 Then InfBox "Suspend Optimize rc: " + Str(rc)
            End If
        End If
    End If

End Sub

' NOTE: this routine will be called automatically when
' the first database is opened or created.
Public Sub Cb4Start()

    Dim rc%

    If cb4Ptr = 0 Then
        cb4Ptr = code4init()
        rc = code4singleOpen(cb4Ptr, 0)
    End If

End Sub

' Deletes a record (if exists),
' then goes back to orig record
Public Function DeleteRecord(tbl_id As Variant, ByVal rec_num&) As Boolean

    Dim cur_rec&, rc&, tbl&, bDeleted As Boolean

    On Error Resume Next
    bDeleted = False
    tbl = TblPtr(tbl_id)
    If tbl <> 0 And rec_num > 0 Then
        cur_rec = d4recNo(tbl)
        If d4go(tbl, rec_num) = r4success Then
            If d4recNo(tbl) = rec_num Then
                Call d4delete(tbl)
                If d4deleted(tbl) <> 0 Then bDeleted = True
            End If
        End If
        rc = d4go(tbl, cur_rec)
    End If

    DeleteRecord = bDeleted
End Function


Function FixStrParms$(d4Ptr&, ByVal old_parm$)

    Dim Parm$, change%, fld_id$, fld_ptr&, pos%, epos%

    'old_parm = parm
    Parm = ""

    Do While True
        change = False
        pos = InStr(UCase(old_parm), "STR(")
        If pos = 0 Then
            Parm = Parm + old_parm
            Exit Do
        End If
        If pos = 1 Then
            change = True
        ElseIf UCase(Mid(old_parm, pos - 1, 1)) <> "B" Then
            change = True
        Else
            change = False
        End If
        Parm = Parm + Left(old_parm, pos + 3)
        old_parm = Mid(old_parm, pos + 4)

        If change Then
            pos = InStr(old_parm, ",")
            epos = InStr(old_parm, ")")
            If epos < pos Or pos = 0 Then
                fld_id = Left(old_parm, epos - 1)
                fld_ptr = fldPtr(d4Ptr, fld_id)
                If fld_ptr <> 0 Then
                    Parm = Parm + fld_id + "," + Trim(Str(f4len(fld_ptr))) + "," + Trim(Str(f4decimals(fld_ptr)))
                    old_parm = Mid(old_parm, epos)
                End If
            End If
        End If
    Loop

    FixStrParms = Parm
End Function

' Go to a record (but avoids most CB errors),
' returns TRUE if the record exists
Public Function GotoRec(tbl_id As Variant, ByVal recno&) As Boolean

    Dim i%, gorec&, tbl&, bSuccess As Boolean

    tbl = TblPtr(tbl_id)
    If tbl <> 0 Then
        gorec = recno
        ' go to bottom if trying to go past
        If gorec > d4recCount(tbl) Then gorec = d4recCount(tbl)
        If gorec > 0 Then
            If d4go(tbl, gorec) = r4success Then
                ' but only return true if didn't try to go past bottom
                If gorec = recno Then bSuccess = True
            End If
        End If
    End If

    GotoRec = bSuccess
End Function

#If 0 Then
Function IndexCreate&(d4Ptr&, ByVal FileName$, tag_str$)
'            "TagName~TagExpression~TagFilter"
'e.g. tags = "SYMBOL~UPPER(SYMBOL)~.NOT.DELETED()~"

    Dim i4Ptr&, i%, fixed_tag_str$

    If d4Ptr = 0 Or cb4Ptr = 0 Then
        IndexCreate = 0
    Else
        i = d4flush(d4Ptr)
        fixed_tag_str = FixStrParms(d4Ptr, tag_str)
        i4Ptr = index_create(cb4Ptr, d4Ptr, FileName, fixed_tag_str)
        IndexCreate = i4Ptr
    End If

End Function
#End If

#If 0 Then
Function MakeCDXifGone%(ByVal db_name$, tags$)
' will recreate the index file if it is missing

    Dim tbl&, rtrn%, cdx_name$

    ' Is CDX missing?
    rtrn = True
    db_name = UCase(Trim(db_name))
    If InStr(db_name, ".") = 0 Then db_name = db_name + ".DBF"
    cdx_name = Parse(db_name, ".", 1) + ".CDX"
    If FileLength(cdx_name) = 0 Then
        KillFile cdx_name
    End If
    If Not FileExist(cdx_name) And FileExist(db_name) Then
        rtrn = False
        ' Open DBF exclusively
        TblClose db_name
        tbl = TblOpen(db_name, True, False)
        If tbl <> 0 Then
            ' Make CDX
            If IndexCreate(tbl, "", tags) = 0 Then
                'InfBox "Could not create CDX file for: " + db_name
            Else
                rtrn = True
            End If
            TblClose tbl
        End If
    End If

    MakeCDXifGone = rtrn
End Function
#End If

' Quickly calculates number of records in the current tag.
Public Function TagRecCount&(tbl_id As Variant)

    Dim rc&, SaveRec&, tbl&
    
    ' Save this record
    tbl = TblPtr(tbl_id)
    SaveRec = d4recNo(tbl)
    
    rc = d4bottom(tbl)
    TagRecCount = TagRecNo(tbl)
    
    ' Go back to original record
    rc = GotoRec(tbl, SaveRec)

End Function

' Quickly calculates how many records into the file
' the current record is relative to the current Tag.
'  (for large databases, this method is many times
'  faster than looping with a "skip 1"!)
' To get the total number of records in the tag, do
' a "d4bottom(tbl)" before calling this.
Public Function TagRecNo&(tbl_id As Variant)

    Dim CurRec&, NumRecs&, rc&, RecInc&, SaveRec&, Divisor%, tbl&

    ' Save this record
    tbl = TblPtr(tbl_id)
    SaveRec = d4recNo(tbl)
    NumRecs = d4recCount(tbl)
    If SaveRec <= 0 Or NumRecs <= 0 Or d4bof(tbl) Then
        TagRecNo = 0
        Exit Function
    End If

    ' Calculate starting increment
    Divisor = 8     ' seems to optimize around 8
    RecInc = 1
    Do While RecInc < NumRecs
        RecInc = RecInc * Divisor
    Loop

    ' Determine how many records into tag we are
    NumRecs = 0
    CurRec = SaveRec
    Do While True
        ' skip backwards "increment" number of records
        If d4skip(tbl, -(RecInc)) <> r4success Then
            ' past beginning, so decrease increment
            rc = d4go(tbl, CurRec)
            If RecInc <= 1 Then Exit Do
            RecInc = RecInc \ Divisor
        Else
            ' not past beginning, so add to sum
            CurRec = d4recNo(tbl)
            NumRecs = NumRecs + RecInc
        End If
    Loop
    NumRecs = NumRecs + 1

    ' Go back to original record
    rc = GotoRec(tbl, SaveRec)

    TagRecNo = NumRecs
End Function

#If 0 Then
' To easily set up a loop to walk through a database:
'   rc = TagScan(tbl, "TAG_NAME")   ' give tag name here
'   Do While TagScan(tbl, "")       ' blank tag name here
'       ... code to execute for record
'   Loop
Function TagScan%(tbl&, Tag$)

    Static initflag%, rtrn%

    rtrn = False
    If Len(Tag) > 0 Then
        ' initialize for scan
        initflag = -1    'invalid
        If TagSelect(tbl, Tag) <> 0 Then
            If d4top(tbl) = r4success Then
                ' valid
                initflag = 1
                rtrn = True
            End If
        End If
    ElseIf initflag = 0 Then
        ' try to go to next record
        If d4skip(tbl, 1) = r4success Then rtrn = True
    ElseIf initflag > 0 Then
        ' just initialized (don't skip yet)
        rtrn = True
        initflag = 0
    ElseIf initflag < 0 Then
        ' just had an invalid initialization
        rtrn = False
        initflag = 0
    End If

    TagScan = rtrn
End Function
#End If

' Search on table (d4Ptr) using the specified index (tag)
' for a value (seek_for).  To require exact match for
' character string, set "bExactString" to True.
Public Function TagSeek(tbl_id As Variant, tag_id As Variant, _
    SeekFor As Variant, _
    Optional ByVal bExactString As Boolean = True) As Boolean

    Dim rc&, bFound As Boolean, tbl&, strSeek$

    bFound = False
    tbl = TblPtr(tbl_id)
    
    ' Select the Tag (index)
    If TagSelect(tbl, tag_id) Then
        If VarType(SeekFor) <> vbString Then
            ' for all numbers
            rc = d4seekDouble(tbl, CDbl(SeekFor))
        ElseIf bExactString Then
            ' (d4seekN() good for binary fields)
            ' for exact chars, add spaces
            strSeek = Left(SeekFor + Space(256), 256)
            rc = d4seekN(tbl, strSeek, Len(strSeek))
        Else
            ' (d4seekN() good for binary fields)
            rc = d4seekN(tbl, SeekFor, Len(SeekFor))
        End If

        If rc = r4success Then bFound = True
    End If

    TagSeek = bFound
End Function

' Makes the specified index (tag) active for this table.
' (pass 0 to use natural record ordering)
Public Function TagSelect(tbl_id As Variant, tag_id As Variant) As Boolean

    Dim tbl&, Tag&, bSuccess As Boolean

    bSuccess = False
    tbl = TblPtr(tbl_id)
    If tbl <> 0 Then
        Tag = TagPtr(tbl, tag_id)
        Call d4tagSelect(tbl, Tag)
        If d4tagSelected(tbl) = Tag Then bSuccess = True
    End If

    TagSelect = bSuccess
End Function

' "tbl_id" can be either the filename or tbl_ptr (d4Ptr).
Public Sub TblClose(tbl_id As Variant)

    Dim tbl&, rc&, strFileName$

    Cb4Start    'if not already

    tbl = TblPtr(tbl_id)
    If tbl <> 0 Then
        ' if open, flush first then close
        rc = d4flush(tbl)
        If d4close(tbl) = r4success Then
            tbl = 0
            If VarType(tbl_id) = vbLong Then
                tbl_id = 0 ' set by reference
            End If
        End If
    End If

End Sub

#If 0 Then
Function TblCreate&(FileName$, Fields$, overwrite%)  ', WithIndexes%)
'               "FldName~Type&Length|..."
' e.g. fields = "SYMBOL~C8|NAME~C22|CONV_FACT~N2|FLAGS~C8"

    Dim d4Ptr As Long, Save%, rc%

    Cb4Start    ' if not already!

    ' save current safety setting
    'Save = code4safety(cb4Ptr, r4check)

    ' set safety
    If overwrite Then
        'rc = code4safety(cb4Ptr, False) ' safety off
        KillFile Parse(FileName, ".", 1) + ".DBF"
        KillFile Parse(FileName, ".", 1) + ".CDX"
        KillFile Parse(FileName, ".", 1) + ".FPT"
    Else
        'rc = code4safety(cb4Ptr, True)  ' safety on
    End If

    If FileExist(Parse(FileName, ".", 1) + ".DBF") Then
        ' file must be locked!
        InfBox "i=[Error] ; Could not create database:|" + FileName
        d4Ptr = 0
    Else
        ' create database
        d4Ptr = tbl_create(cb4Ptr, FileName, Fields)

        ' set optimizations
        TblOptimize d4Ptr, OPT4ALL, OPT4EXCLUSIVE
    End If

    ' restore safety setting
    'rc = code4safety(cb4Ptr, Save)

    TblCreate = d4Ptr
End Function
#End If

#If 0 Then
Function TblNextID&(d4Ptr&, fld_id As Variant)
' fld_id can be either the field name (easiest)
' or the "field ptr" (most efficient) of the ID field

    Dim fld_ptr&, rc%, id_num&

    ' Assign field ptr
    fld_ptr = fldPtr(d4Ptr, fld_id)
    If fld_ptr = 0 Then
        TblNextID = 0
        Exit Function
    End If

    ' select TAG name with same name as FLD name (if exists)
    If VarType(fld_id) = vbString Then
        rc = TagSelect(d4Ptr, CStr(fld_id))
    End If

    ' get ID of last record
    id_num = 0
    If d4bottom(d4Ptr) = r4success Then
        id_num = f4num(fld_ptr)
    End If

    ' append new record and assign incremented ID
    If d4appendBlank(d4Ptr) = r4success Then
        id_num = id_num + 1
        Call f4putNum(fld_ptr, id_num)
    End If

    TblNextID = id_num
End Function
#End If

Public Function TblOpen&(ByVal strFileName$, _
    Optional ByVal bExclusive As Boolean = False, _
    Optional ByVal bReadOnly As Boolean = False)

    Dim rc&, tbl&, strCdxFile$
    Dim savAccessMode%, savReadOnly%, savAutoOpen%

    Cb4Start    ' if not already!

    ' See if alias is recognized as already open
    tbl = code4data(cb4Ptr, FileBase(strFileName))
    If tbl = 0 Then
        If FileExist(strFileName) Then
            ' Save current settings
            savAccessMode = code4accessMode(cb4Ptr, r4check)
            savReadOnly = code4readOnly(cb4Ptr, r4check)
            savAutoOpen = code4autoOpen(cb4Ptr, r4check)
    
            ' Set access mode
            If bExclusive Then
                rc = code4accessMode(cb4Ptr, OPEN4DENY_RW)
            Else
                rc = code4accessMode(cb4Ptr, OPEN4DENY_NONE)
            End If
            If bReadOnly Then
                rc = code4readOnly(cb4Ptr, True)
            Else
                rc = code4readOnly(cb4Ptr, False)
            End If
    
            ' See if CDX file exists
            strCdxFile = Parse(strFileName, ".", 1) + ".CDX"
            If FileExist(strCdxFile) Then
                rc = code4autoOpen(cb4Ptr, True)
            Else
                rc = code4autoOpen(cb4Ptr, False)
            End If
    
            ' Open DBF file
            tbl = d4open(cb4Ptr, strFileName)
            If tbl = 0 And bExclusive Then
                ' wait a second, see if other process frees up
                Sleep 2
                tbl = d4open(cb4Ptr, strFileName)
            End If
    
            If tbl <> 0 Then
                ' open CDX if not already
                If FileExist(strCdxFile) Then
                    KillFile Parse(strCdxFile, ".", 1) + ".XXX"
                    On Error Resume Next
                    Name strCdxFile As Parse(strCdxFile, ".", 1) + ".XXX"
                    If FileExist(Parse(strCdxFile, ".", 1) + ".XXX") Then
                        ' must not be open yet
                        Name Parse(strCdxFile, ".", 1) + ".XXX" As strCdxFile
                        rc = i4open(tbl, strCdxFile)
                    End If
                End If
    
                ' set optimizations
                TblOptimize tbl, OPT4ALL, OPT4EXCLUSIVE
            End If
    
            ' restore settings
            rc = code4accessMode(cb4Ptr, savAccessMode)
            rc = code4readOnly(cb4Ptr, savReadOnly)
            rc = code4autoOpen(cb4Ptr, savAutoOpen)
        End If
    End If

    TblOpen = tbl
End Function

Public Sub TblOptimize(tbl_id As Variant, _
    Optional ByVal ReadOpt% = OPT4ALL, _
    Optional ByVal WriteOpt% = OPT4EXCLUSIVE)
' Mostly used internally (by TblOpen function).
' Options: OPT4EXCLUSIVE = -1, OPT4ALL = 1, OPT4NONE = 0

    Dim rc&, tbl&

    tbl = TblPtr(tbl_id)
    If tbl <> 0 Then
        rc = d4optimize(tbl, ReadOpt)
        rc = d4optimizeWrite(tbl, WriteOpt)
    End If

End Sub

' "tbl_id" can be either the filename or tbl_ptr (d4Ptr).
Public Function TblPtr&(tbl_id As Variant)
    
    Dim iVarType%
    iVarType = VarType(tbl_id)
    If iVarType = vbLong Then
        TblPtr = tbl_id
    ElseIf iVarType = vbString Then
        ' See if alias is recognized as already open
        TblPtr = code4data(cb4Ptr, FileBase(tbl_id))
    Else
        TblPtr = 0
    End If

End Function

' "tag_id" can be either the name or tag_ptr (t4Ptr).
Public Function TagPtr&(tbl_id As Variant, tag_id As Variant)
    
    Dim iVarType%, tbl&, Tag&
    
    iVarType = VarType(tag_id)
    If iVarType = vbLong Then
        Tag = tag_id
    ElseIf iVarType = vbString Then
        tbl = TblPtr(tbl_id)
        If tbl <> 0 Then
            Tag = d4tag(tbl, UCase(tag_id))
        End If
    Else
        Tag = 0
    End If

    TagPtr = Tag
End Function

' "fld_id" can be either the name or fld_ptr (f4Ptr).
Public Function fldPtr&(tbl_id As Variant, fld_id As Variant)
    
    Dim iVarType%, tbl&, fld&
    
    iVarType = VarType(fld_id)
    If iVarType = vbLong Then
        fld = fld_id
    ElseIf iVarType = vbString Then
        tbl = TblPtr(tbl_id)
        If tbl <> 0 Then
            fld = d4field(tbl, UCase(fld_id))
        End If
    Else
        fld = 0
    End If

    fldPtr = fld
End Function

Public Sub f4assignLogical(fldPtr&, fldVal As Boolean)
    
    If fldVal = True Then
        f4assignChar fldPtr, Asc("T")
    Else
        f4assignChar fldPtr, Asc("F")
    End If
    
End Sub

Public Function f4Logical(ByVal fldPtr&) As Boolean

    Select Case f4char(fldPtr)
        'Case Asc("Y"), Asc("y"), Asc("T"), Asc("t"), Asc("1")
        Case 89, 121, 84, 116, 49
            f4Logical = True
        Case Else
            f4Logical = False
    End Select

End Function

'To flush a table, or all tables if passed 0.
Public Sub TblFlush(Optional tbl_id As Variant = 0&)
    
    Dim tbl&
    If cb4Ptr <> 0 Then
        tbl = TblPtr(tbl_id)
        If tbl = 0 Then
            tbl = code4readLock(cb4Ptr, r4check)
            code4flush cb4Ptr
            code4unlock cb4Ptr
        Else
            d4flush tbl
            d4unlock tbl
        End If
    End If
    
End Sub
