Attribute VB_Name = "mCodeBase"
' codebase.bas  (c)Copyright Sequiter Software Inc., 1988-1999.  All rights reserved.

'Data Types Used by CodeBase

Type FIELD4INFOCB
    fName As Long        ' C string (which is different than a Basic String)
    ftype As Integer
    flength As Integer
    fdecimals As Integer
    fnulls As Integer
End Type

Type FIELD4INFO          ' Corresponding Basic structure
    fName As String
    ftype As String * 1
    flength As Integer
    fdecimals As Integer
    fnulls As Integer
End Type

Type TAG4INFOCB
    Name As Long         ' C string
    Expression As Long   ' C string
    Filter As Long       ' C string
    unique As Integer
    descending As Integer
End Type

Type TAG4INFO
    Name As String
    Expression As String
    Filter As String
    unique As Integer
    descending As Integer
End Type

'===================================================================================
'
'     CODE4 Access  function prototypes
'
'===================================================================================

Declare Function code4accessMode% Lib "c4dll.dll" (ByVal c4&, ByVal Value%)
Declare Function code4autoOpen% Lib "c4dll.dll" (ByVal c4&, ByVal Value%)
Declare Function code4calcCreate% Lib "c4dll.dll" (ByVal c4&, ByVal expr4&, ByVal fcnName$)
Declare Sub code4calcReset Lib "c4dll.dll" (ByVal c4&)
Declare Function code4codePage% Lib "c4dll.dll" (ByVal c4&, ByVal Value%)
Declare Function code4collatingSequence% Lib "c4dll.dll" (ByVal c4&, ByVal Value%)
Declare Function code4collate% Lib "c4dll.dll" (ByVal c4&, ByVal Value%)
Declare Function code4compatibility% Lib "c4dll.dll" (ByVal c4&, ByVal Value%)
Declare Function code4connect% Lib "c4dll.dll" (ByVal c4&, ByVal serverId$, ByVal processId$, ByVal userName$, ByVal password$, ByVal protocol$)
Declare Function code4close% Lib "c4dll.dll" (ByVal c4&)
Declare Function code4createTemp% Lib "c4dll.dll" (ByVal c4&, ByVal Value%)
Declare Function code4data& Lib "c4dll.dll" (ByVal c4&, ByVal AliasName$)
Declare Function code4dateFormatVB& Lib "c4dll.dll" Alias "code4dateFormat" (ByVal c4&)
Declare Function code4dateFormatSet% Lib "c4dll.dll" (ByVal c4&, ByVal fmt$)
Declare Function code4errCreate% Lib "c4dll.dll" (ByVal c4&, ByVal Value%)
Declare Function code4errDefaultUnique% Lib "c4dll.dll" (ByVal c4&, ByVal Value%)
Declare Function code4errorCode% Lib "c4dll.dll" (ByVal c4&, ByVal Value%)
Declare Function code4errExpr% Lib "c4dll.dll" (ByVal c4&, ByVal Value%)
Declare Function code4errFieldName% Lib "c4dll.dll" (ByVal c4&, ByVal Value%)
Declare Function code4errGo% Lib "c4dll.dll" (ByVal c4&, ByVal Value%)
Declare Function code4errSkip% Lib "c4dll.dll" (ByVal c4&, ByVal Value%)
Declare Function code4errTagName% Lib "c4dll.dll" (ByVal c4&, ByVal Value%)
Declare Sub code4exit Lib "c4dll.dll" (ByVal c4&)
Declare Function code4fileFlush% Lib "c4dll.dll" (ByVal c4&, ByVal Value%)
Declare Function code4flush% Lib "c4dll.dll" (ByVal c4&)
Declare Function code4hInst% Lib "c4dll.dll" (ByVal c4&, ByVal Value%)
Declare Function code4indexBlockSize% Lib "c4dll.dll" (ByVal c4&)
Declare Function code4indexBlockSizeSet% Lib "c4dll.dll" (ByVal c4&, ByVal Value%)
Declare Function code4indexExtensionVB& Lib "c4dll.dll" Alias "code4indexExtension" (ByVal c4&)
Declare Function code4hWnd& Lib "c4dll.dll" (ByVal c4&, ByVal Value&)
Declare Function code4init& Lib "c4dll.dll" Alias "code4initVB" ()
Declare Function code4initUndo% Lib "c4dll.dll" (ByVal c4&)
Declare Sub code4largeOn Lib "c4dll.dll" (ByVal c4&)
Declare Function code4lock% Lib "c4dll.dll" (ByVal c4&)
Declare Function code4lockAttempts% Lib "c4dll.dll" (ByVal c4&, ByVal Value%)
Declare Function code4lockAttemptsSingle% Lib "c4dll.dll" (ByVal c4&, ByVal Value%)
Declare Sub code4lockClear Lib "c4dll.dll" (ByVal c4&)
Declare Function code4lockDelay& Lib "c4dll.dll" (ByVal c4&, ByVal Value&)
Declare Function code4lockEnforce% Lib "c4dll.dll" (ByVal c4&, ByVal Value%)
Declare Function code4lockFileNameVB& Lib "c4dll.dll" Alias "code4lockFileName" (ByVal c4&)
Declare Function code4lockItem& Lib "c4dll.dll" (ByVal c4&)
Declare Function code4lockNetworkIdVB& Lib "c4dll.dll" Alias "code4lockNetworkId" (ByVal c4&)
Declare Function code4lockUserIdVB& Lib "c4dll.dll" Alias "code4lockUserId" (ByVal c4&)
Declare Function code4log% Lib "c4dll.dll" (ByVal c4&, ByVal Value%)
Declare Function code4logCreate% Lib "c4dll.dll" (ByVal c4&, ByVal logName$, ByVal UserID$)
Declare Function code4logFileNameVB& Lib "c4dll.dll" (ByVal c4&)
Declare Function code4logOpen% Lib "c4dll.dll" (ByVal c4&, ByVal logName$, ByVal UserID$)
Declare Sub code4logOpenOff Lib "c4dll.dll" (ByVal c4&)
Declare Function code4memExpandBlock% Lib "c4dll.dll" (ByVal c4&, ByVal Value%)
Declare Function code4memExpandData% Lib "c4dll.dll" (ByVal c4&, ByVal Value%)
Declare Function code4memExpandIndex% Lib "c4dll.dll" (ByVal c4&, ByVal Value%)
Declare Function code4memExpandLock% Lib "c4dll.dll" (ByVal c4&, ByVal Value%)
Declare Function code4memExpandTag% Lib "c4dll.dll" (ByVal c4&, ByVal Value%)
Declare Function code4memSizeBlock& Lib "c4dll.dll" (ByVal c4&, ByVal Value&)
Declare Function code4memSizeBuffer& Lib "c4dll.dll" (ByVal c4&, ByVal Value&)
Declare Function code4memSizeMemo% Lib "c4dll.dll" (ByVal c4&, ByVal Value%)
Declare Function code4memSizeMemoExpr& Lib "c4dll.dll" (ByVal c4&, ByVal Value&)
Declare Function code4memSizeSortBuffer& Lib "c4dll.dll" (ByVal c4&, ByVal Value&)
Declare Function code4memSizeSortPool& Lib "c4dll.dll" (ByVal c4&, ByVal Value&)
Declare Function code4memStartBlock% Lib "c4dll.dll" (ByVal c4&, ByVal Value%)
Declare Function code4memStartData% Lib "c4dll.dll" (ByVal c4&, ByVal Value%)
Declare Function code4memStartIndex% Lib "c4dll.dll" (ByVal c4&, ByVal Value%)
Declare Function code4memStartLock% Lib "c4dll.dll" (ByVal c4&, ByVal Value%)
Declare Function code4memStartMax& Lib "c4dll.dll" (ByVal c4&, ByVal Value&)
Declare Function code4memStartTag% Lib "c4dll.dll" (ByVal c4&, ByVal Value%)
Declare Function code4errOff% Lib "c4dll.dll" (ByVal c4&, ByVal Value%)
Declare Function code4errOpen% Lib "c4dll.dll" (ByVal c4&, ByVal Value%)
Declare Function code4optAll% Lib "c4dll.dll" (ByVal c4&)
Declare Function code4optimize% Lib "c4dll.dll" (ByVal c4&, ByVal Value%)
Declare Function code4optStart% Lib "c4dll.dll" (ByVal c4&)
Declare Function code4optSuspend% Lib "c4dll.dll" (ByVal c4&)
Declare Function code4optimizeWrite% Lib "c4dll.dll" (ByVal c4&, ByVal Value%)
Declare Function code4readLock% Lib "c4dll.dll" (ByVal c4&, ByVal Value%)
Declare Function code4readOnly% Lib "c4dll.dll" (ByVal c4&, ByVal Value%)
Declare Function code4errRelate% Lib "c4dll.dll" (ByVal c4&, ByVal Value%)
Declare Function code4safety% Lib "c4dll.dll" (ByVal c4&, ByVal Value%)
Declare Function code4singleOpen% Lib "c4dll.dll" (ByVal c4&, ByVal Value%)
Declare Function code4timeout& Lib "c4dll.dll" (ByVal c4&)
Declare Sub code4timeoutSet Lib "c4dll.dll" (ByVal c4&, ByVal Value&)
Declare Function code4tranStart% Lib "c4dll.dll" (ByVal c4&)
Declare Function code4tranStatus% Lib "c4dll.dll" Alias "code4tranStatusCB" (ByVal c4&)
Declare Function code4tranCommit% Lib "c4dll.dll" (ByVal c4&)
Declare Function code4tranRollback% Lib "c4dll.dll" (ByVal c4&)
Declare Function code4unlock% Lib "c4dll.dll" (ByVal c4&)
Declare Function code4unlockAuto% Lib "c4dll.dll" Alias "code4unlockAutoCB" (ByVal c4&)
Declare Sub code4unlockAutoSet Lib "c4dll.dll" Alias "code4unlockAutoSetCB" (ByVal c4&, ByVal Value%)
Declare Sub code4verifySet Lib "c4dll.dll" (ByVal c4&, ByVal Value$)

'===============================================================================================
'
'                                 CodeControls function prototypes
'
'-----------------------------------------------------------------------------------------------

'TLB 5/10/00: commented these out since not using them and they don't exist!
'Declare Function ctrl4init Lib "cc2.vbx" Alias "ctrl4initVB" () As Long
'Declare Function ctrl4initUndo Lib "cc2.vbx" Alias "ctrl4initUndoVB" (ByVal code As Long) As Integer

'===============================================================================================
'
'                                 Data File Functions' Prototypes
'
'-----------------------------------------------------------------------------------------------
Declare Function d4aliasCB& Lib "c4dll.dll" Alias "d4alias" (ByVal d4&)
Declare Sub d4aliasSet Lib "c4dll.dll" (ByVal d4&, ByVal AliasValue$)
Declare Function d4append% Lib "c4dll.dll" (ByVal d4&)
Declare Function d4appendBlank% Lib "c4dll.dll" (ByVal d4&)
Declare Function d4appendStart% Lib "c4dll.dll" (ByVal d4&, ByVal UseMemoEntries%)
Declare Sub d4blank Lib "c4dll.dll" (ByVal d4&)
Declare Function d4bof% Lib "c4dll.dll" (ByVal d4&)
Declare Function d4bottom% Lib "c4dll.dll" (ByVal d4&)
Declare Function d4changed% Lib "c4dll.dll" (ByVal d4&, ByVal intFlag%)
Declare Function d4check% Lib "c4dll.dll" (ByVal d4&)
Declare Function d4close% Lib "c4dll.dll" (ByVal d4&)
Declare Function d4createCB& Lib "c4dll.dll" Alias "d4create" (ByVal c4&, ByVal DbfName$, fieldinfo As Any, tagInfo As Any)
Declare Sub d4delete Lib "c4dll.dll" (ByVal d4&)
Declare Function d4deleted% Lib "c4dll.dll" (ByVal d4&)
Declare Function d4eof% Lib "c4dll.dll" (ByVal d4&)
Declare Function d4field& Lib "c4dll.dll" (ByVal d4&, ByVal FieldName$)
Declare Function d4fieldInfo& Lib "c4dll.dll" (ByVal d4&)
Declare Function d4fieldJ& Lib "c4dll.dll" (ByVal d4&, ByVal JField%)
Declare Function d4fieldNumber% Lib "c4dll.dll" (ByVal d4&, ByVal FieldName$)
Declare Function d4fileNameCB& Lib "c4dll.dll" Alias "d4fileName" (ByVal d4&)
Declare Function d4flush% Lib "c4dll.dll" (ByVal d4&)
Declare Function d4freeBlocks% Lib "c4dll.dll" (ByVal d4&)
Declare Function d4goLow% Lib "c4dll.dll" (ByVal d4&, ByVal RecNum&, ByVal goForWrite%)
Declare Function d4goBof% Lib "c4dll.dll" (ByVal d4&)
Declare Function d4goEof% Lib "c4dll.dll" (ByVal d4&)
Declare Function d4index& Lib "c4dll.dll" (ByVal d4&, ByVal IndexName$)
Declare Function d4log% Lib "c4dll.dll" Alias "d4logVB" (ByVal d4&, ByVal logging%)
Declare Function d4lock% Lib "c4dll.dll" Alias "d4lockVB" (ByVal d4&, ByVal recordNum&)
Declare Function d4lockAdd% Lib "c4dll.dll" (ByVal d4&, ByVal recordNum&)
Declare Function d4lockAddAll% Lib "c4dll.dll" (ByVal d4&)
Declare Function d4lockAddAppend% Lib "c4dll.dll" (ByVal d4&)
Declare Function d4lockAddFile% Lib "c4dll.dll" (ByVal d4&)
Declare Function d4lockAll% Lib "c4dll.dll" Alias "d4lockAllVB" (ByVal d4&)
Declare Function d4lockAppend% Lib "c4dll.dll" Alias "d4lockAppendVB" (ByVal d4&)
Declare Function d4lockFile% Lib "c4dll.dll" Alias "d4lockFileVB" (ByVal d4&)
Declare Function d4logStatus% Lib "c4dll.dll" Alias "d4logStatusCB" (ByVal d4&)
Declare Function d4memoCompress% Lib "c4dll.dll" (ByVal d4&)
Declare Function d4numFields% Lib "c4dll.dll" (ByVal d4&)
Declare Function d4open& Lib "c4dll.dll" (ByVal c4&, ByVal DbfName$)
Declare Function d4openClone& Lib "c4dll.dll" (ByVal d4&)
Declare Function d4optimize% Lib "c4dll.dll" Alias "d4optimizeVB" (ByVal d4&, ByVal OptFlag%)
Declare Function d4optimizeWrite% Lib "c4dll.dll" Alias "d4optimizeWriteVB" (ByVal d4&, ByVal OptFlag%)
Declare Function d4pack% Lib "c4dll.dll" (ByVal d4&)
Declare Function d4packData% Lib "c4dll.dll" (ByVal d4&)
Declare Function d4position# Lib "c4dll.dll" (ByVal d4&)
Declare Function d4positionSet% Lib "c4dll.dll" (ByVal d4&, ByVal Percentage#)
Declare Sub d4recall Lib "c4dll.dll" (ByVal d4&)
Declare Function d4recCount& Lib "c4dll.dll" Alias "d4recCountDo" (ByVal d4&)
Declare Function d4recNo& Lib "c4dll.dll" Alias "d4recNoLow" (ByVal d4&)
Declare Function d4record& Lib "c4dll.dll" Alias "d4recordLow" (ByVal d4&)
Declare Function d4recWidth& Lib "c4dll.dll" Alias "d4recWidth_v" (ByVal d4&)
Declare Function d4remove% Lib "c4dll.dll" (ByVal d4&)
Declare Function d4refresh% Lib "c4dll.dll" (ByVal d4&)
Declare Function d4refreshRecord% Lib "c4dll.dll" (ByVal d4&)
Declare Function d4reindex% Lib "c4dll.dll" (ByVal d4&)
Declare Function d4seek% Lib "c4dll.dll" (ByVal d4&, ByVal seekValue$)
Declare Function d4seekDouble% Lib "c4dll.dll" (ByVal d4&, ByVal Value#)
Declare Function d4seekN% Lib "c4dll.dll" (ByVal d4&, ByVal seekValue$, ByVal seekLen%)
Declare Function d4seekNext% Lib "c4dll.dll" (ByVal d4&, ByVal seekValue$)
Declare Function d4seekNextDouble% Lib "c4dll.dll" (ByVal d4&, ByVal seekValue#)
Declare Function d4seekNextN% Lib "c4dll.dll" (ByVal d4&, ByVal seekValue$, ByVal seekLen%)
Declare Function d4skip% Lib "c4dll.dll" (ByVal d4&, ByVal NumberRecords&)
Declare Function d4tag& Lib "c4dll.dll" (ByVal d4&, ByVal TagName$)
Declare Function d4tagDefault& Lib "c4dll.dll" (ByVal d4&)
Declare Function d4tagNext& Lib "c4dll.dll" (ByVal d4&, ByVal TagOn&)
Declare Function d4tagPrev& Lib "c4dll.dll" (ByVal d4&, ByVal TagOn&)
Declare Sub d4tagSelect Lib "c4dll.dll" (ByVal d4&, ByVal tPtr&)
Declare Function d4tagSelected& Lib "c4dll.dll" (ByVal d4&)
Declare Function d4tagSync% Lib "c4dll.dll" (ByVal d4&, ByVal tPtr&)
Declare Function d4top% Lib "c4dll.dll" (ByVal d4&)
Declare Function d4unlock% Lib "c4dll.dll" (ByVal d4&)
Declare Function d4unlockFiles% Lib "c4dll.dll" Alias "code4unlock" (ByVal d4&)
Declare Function d4write% Lib "c4dll.dll" Alias "d4writeVB" (ByVal d4&, ByVal RecNum&)
Declare Function d4zap% Lib "c4dll.dll" (ByVal d4&, ByVal StartRecord&, ByVal EndRecord&)

'===============================================================================================
'
'                                   Date Functions' Prototypes
'
'-----------------------------------------------------------------------------------------------
Declare Sub date4assignLow Lib "c4dll.dll" (ByVal dateForm$, ByVal julianDay&, ByVal isOle%)
Declare Function date4cdowCB& Lib "c4dll.dll" Alias "date4cdow" (ByVal dateForm$)
Declare Function date4cmonthCB& Lib "c4dll.dll" Alias "date4cmonth" (ByVal dateForm$)
Declare Function date4day% Lib "c4dll.dll" Alias "date4day_v" (ByVal dateForm$)
Declare Function date4dow% Lib "c4dll.dll" (ByVal dateForm$)
Declare Sub date4formatCB Lib "c4dll.dll" Alias "date4format" (ByVal dateForm$, ByVal Result$, ByVal pic$)
Declare Sub date4initCB Lib "c4dll.dll" Alias "date4init" (ByVal dateForm$, ByVal Value$, ByVal pic$)
Declare Function date4isLeap% Lib "c4dll.dll" (ByVal dateForm$)
Declare Function date4long& Lib "c4dll.dll" (ByVal dateForm$)
Declare Function date4month% Lib "c4dll.dll" Alias "date4month_v" (ByVal dateForm$)
Declare Sub date4timeNow Lib "c4dll.dll" (ByVal TimeForm$)
Declare Sub date4todayCB Lib "c4dll.dll" Alias "date4today" (ByVal dateForm$)
Declare Function date4year% Lib "c4dll.dll" Alias "date4year_v" (ByVal dateForm$)

'===============================================================================================
'
'                          Error  Functions' Prototypes
'
'-----------------------------------------------------------------------------------------------
Declare Function error4% Lib "c4dll.dll" Alias "error4VB" (ByVal c4&, ByVal errCode%, ByVal extraInfo&)
Declare Sub error4exitTest Lib "c4dll.dll" (ByVal c4&)
Declare Function error4describe% Lib "c4dll.dll" Alias "error4describeVB" (ByVal c4&, ByVal errCode%, ByVal extraInfo&, ByVal desc1$, ByVal desc2$, ByVal desc3$)
Declare Function error4file% Lib "c4dll.dll" (ByVal c4&, ByVal FileName$, ByVal overwrite%)
Declare Function error4set% Lib "c4dll.dll" (ByVal c4&, ByVal errCode%)
Declare Function error4textCB& Lib "c4dll.dll" Alias "error4text" (ByVal c4&, ByVal errCode&)

'===============================================================================================
'
'                          Expression Evaluation Functions' Prototypes
'
'-----------------------------------------------------------------------------------------------
Declare Function expr4data& Lib "c4dll.dll" Alias "expr4dataCB" (ByVal exprPtr&)
Declare Function expr4double# Lib "c4dll.dll" (ByVal exprPtr&)
Declare Sub expr4free Lib "c4dll.dll" Alias "expr4freeCB" (ByVal exprPtr&)
Declare Function expr4len& Lib "c4dll.dll" Alias "expr4lenCB" (ByVal exprPtr&)
Declare Function expr4nullLow% Lib "c4dll.dll" (ByVal exprPtr&, ByVal forAdd%)
Declare Function expr4parse& Lib "c4dll.dll" Alias "expr4parseCB" (ByVal d4&, ByVal Expression$)
Declare Function expr4sourceCB& Lib "c4dll.dll" Alias "expr4source" (ByVal exprPtr&)
Declare Function expr4strCB& Lib "c4dll.dll" Alias "expr4str" (ByVal exprPtr&)
Declare Function expr4true% Lib "c4dll.dll" (ByVal exprPtr&)
Declare Function expr4typeCB% Lib "c4dll.dll" (ByVal exprPtr&)

'===============================================================================================
'
'                            Field Functions' Prototypes
'
'-----------------------------------------------------------------------------------------------
Declare Sub f4assignChar Lib "c4dll.dll" Alias "f4assignCharVB" (ByVal fPtr&, ByVal Char%)
Declare Sub f4assignCurrency Lib "c4dll.dll" (ByVal fPtr&, ByVal Value$)
Declare Sub f4assignDateTime Lib "c4dll.dll" (ByVal fPtr&, ByVal Value$)
Declare Sub f4assignDouble Lib "c4dll.dll" (ByVal fPtr&, ByVal Value#)
Declare Sub f4assignField Lib "c4dll.dll" (ByVal fPtrTo&, ByVal fPtrFrom&)
Declare Sub f4assignIntVB Lib "c4dll.dll" (ByVal fPtr&, ByVal Value%)
Declare Sub f4assignLong Lib "c4dll.dll" (ByVal fPtr&, ByVal Value&)
Declare Sub f4assignN Lib "c4dll.dll" Alias "f4assignNVB" (ByVal fPtr&, ByVal Value$, ByVal Length%)
Declare Sub f4assignNull Lib "c4dll.dll" (ByVal fPtr&)
Declare Sub f4blank Lib "c4dll.dll" (ByVal fPtr&)
Declare Function f4char% Lib "c4dll.dll" (ByVal fPtr&)
Declare Function f4currencyCB& Lib "c4dll.dll" Alias "f4currency" (ByVal fPtr&, ByVal numDec%)
Declare Function f4data& Lib "c4dll.dll" (ByVal fPtr&)
Declare Function f4dateTimeCB& Lib "c4dll.dll" Alias "f4dateTime" (ByVal fPtr&)
Declare Function f4decimals% Lib "c4dll.dll" (ByVal fPtr&)
Declare Function f4double# Lib "c4dll.dll" (ByVal fPtr&)
Declare Function f4int% Lib "c4dll.dll" (ByVal fPtr&)
Declare Function f4len% Lib "c4dll.dll" Alias "f4len_v" (ByVal fPtr&)
Declare Function f4long& Lib "c4dll.dll" (ByVal fPtr&)
Declare Function f4memoAssign% Lib "c4dll.dll" (ByVal fPtr&, ByVal Value$)
Declare Function f4memoAssignN% Lib "c4dll.dll" Alias "f4memoAssignNVB" (ByVal fPtr&, ByVal Value$, ByVal Length%)
Declare Sub f4memoFree Lib "c4dll.dll" (ByVal fPtr&)
Declare Function f4memoLen& Lib "c4dll.dll" (ByVal fPtr&)
Declare Function f4memoNcpy% Lib "c4dll.dll" (ByVal fPtr&, ByVal MemPtr$, ByVal memLen%)
Declare Function f4memoPtr& Lib "c4dll.dll" (ByVal fPtr&)
Declare Function f4nameCB& Lib "c4dll.dll" Alias "f4name" (ByVal fPtr&)
Declare Function f4ncpyCB% Lib "c4dll.dll" Alias "f4ncpy" (ByVal fPtr&, ByVal MemPtr$, ByVal memLength%)
Declare Function f4number% Lib "c4dll.dll" (ByVal fPtr&)
Declare Function f4null% Lib "c4dll.dll" (ByVal fPtr&)
Declare Function f4ptr& Lib "c4dll.dll" (ByVal fPtr&)
Declare Function f4strCB& Lib "c4dll.dll" Alias "f4str" (ByVal fPtr&)
Declare Function f4true% Lib "c4dll.dll" (ByVal fPtr&)
Declare Function f4type% Lib "c4dll.dll" (ByVal fPtr&)


'===============================================================================================
'
'                             Index Functions' Prototypes
'
'-----------------------------------------------------------------------------------------------
Declare Function i4close% Lib "c4dll.dll" (ByVal i4&)
Declare Function i4createCB& Lib "c4dll.dll" Alias "i4create" (ByVal d4&, ByVal FileName As Any, tagInfo As TAG4INFOCB)
Declare Function i4fileNameCB& Lib "c4dll.dll" Alias "i4fileName" (ByVal i4&)
Declare Function i4openCB& Lib "c4dll.dll" Alias "i4open" (ByVal d4&, ByVal FileName As Any)
Declare Function i4reindex% Lib "c4dll.dll" (ByVal i4&)
Declare Function i4tag& Lib "c4dll.dll" (ByVal i4&, ByVal FileName$)
Declare Function i4tagInfo& Lib "c4dll.dll" (ByVal i4&)
Declare Function i4tagAddCB% Lib "c4dll.dll" Alias "i4tagAdd" (ByVal i4&, tagInfo As TAG4INFOCB)

'===============================================================================================
'
'                            Relate Functions' Prototypes
'
'-----------------------------------------------------------------------------------------------

Declare Function relate4bottom% Lib "c4dll.dll" (ByVal r4&)
Declare Sub relate4changed Lib "c4dll.dll" (ByVal r4&)
Declare Function relate4createSlave& Lib "c4dll.dll" (ByVal r4&, ByVal d4&, ByVal mExpr$, ByVal t4 As Any)
Declare Function relate4data& Lib "c4dll.dll" Alias "relate4dataCB" (ByVal r4&)
Declare Function relate4dataTag& Lib "c4dll.dll" Alias "relate4dataTagCB" (ByVal r4&)
Declare Function relate4doAll% Lib "c4dll.dll" (ByVal r4&)
Declare Function relate4doOne% Lib "c4dll.dll" (ByVal r4&)
Declare Function relate4eof% Lib "c4dll.dll" (ByVal r4&)
Declare Function relate4errorAction% Lib "c4dll.dll" Alias "relate4errorActionVB" (ByVal r4&, ByVal ErrAction%)
Declare Function relate4free% Lib "c4dll.dll" Alias "relate4freeVB" (ByVal r4&, ByVal CloseFlag%)
Declare Function relate4init& Lib "c4dll.dll" (ByVal d4&)
Declare Function relate4lockAdd% Lib "c4dll.dll" (ByVal r4&)
Declare Function relate4master& Lib "c4dll.dll" Alias "relate4masterCB" (ByVal r4&)
Declare Function relate4masterExprCB& Lib "c4dll.dll" (ByVal r4&)
Declare Function relate4matchLen% Lib "c4dll.dll" Alias "relate4matchLenVB" (ByVal r4&, ByVal Length%)
Declare Function relate4next% Lib "c4dll.dll" (r4&)
Declare Function relate4optimizeable% Lib "c4dll.dll" (ByVal r4&)
Declare Function relate4querySet% Lib "c4dll.dll" (ByVal r4&, ByVal Expr As String)
Declare Function relate4retrieve& Lib "c4dll.dll" (ByVal c4&, ByVal FileName$, ByVal openFiles%, ByVal dataPathName$)
Declare Function relate4save% Lib "c4dll.dll" (ByVal rel4&, ByVal FileName$, ByVal savePathNames%)
Declare Function relate4skip% Lib "c4dll.dll" (ByVal r4&, ByVal NumRecs&)
Declare Function relate4skipEnable% Lib "c4dll.dll" Alias "relate4skipEnableVB" (ByVal r4&, ByVal DoEnable%)
Declare Function relate4sortSet% Lib "c4dll.dll" (ByVal r4&, ByVal Expr As String)
Declare Function relate4top% Lib "c4dll.dll" (ByVal r4&)
Declare Function relate4topMaster& Lib "c4dll.dll" (ByVal r4&)
Declare Function relate4type% Lib "c4dll.dll" Alias "relate4typeVB" (ByVal r4&, ByVal rType%)

'===============================================================================================
'
'  Report function prototypes
'
'================================================================================================
Declare Function report4caption% Lib "c4dll.dll" (ByVal r4&, ByVal Caption$)
Declare Function report4currency% Lib "c4dll.dll" (ByVal r4&, ByVal currncy$)
Declare Function report4dateFormat% Lib "c4dll.dll" (ByVal r4&, ByVal dateFmt$)
Declare Function report4decimal% Lib "c4dll.dll" Alias "report4decimal_v" (ByVal r4&, ByVal decChar$)
Declare Function report4do% Lib "c4dll.dll" Alias "report4doCB" (ByVal r4&)
Declare Sub report4freeLow Lib "c4dll.dll" (ByVal r4&, ByVal freeRelate%, ByVal closeFiles%, ByVal doPrinterFree%)
Declare Function report4margins% Lib "c4dll.dll" (ByVal r4&, ByVal mLeft&, ByVal mRight&, ByVal mTop&, ByVal mBottom&, ByVal uType%)
Declare Function report4pageSize% Lib "c4dll.dll" (ByVal r4&, ByVal pHeight&, ByVal pWidth&, ByVal uType%)
#If Win16 Then
   Declare Function report4parent16% Lib "c4dll.dll" Alias "report4parent" (ByVal r4&, ByVal hWnd%)
#End If
#If Win32 Then
   Declare Function report4parent32% Lib "c4dll.dll" Alias "report4parent" (ByVal r4&, ByVal hWnd&)
#End If
Declare Sub report4printerSelect Lib "c4dll.dll" (ByVal r4&)
Declare Function report4querySet% Lib "c4dll.dll" (ByVal r4&, ByVal queryExpr$)
Declare Function report4relate& Lib "c4dll.dll" (ByVal r4&)
Declare Function report4retrieve& Lib "c4dll.dll" (ByVal c4&, ByVal FileName$, ByVal openFiles%, ByVal DataPath$)
Declare Function report4save% Lib "c4dll.dll" (ByVal r4&, ByVal FileName$, ByVal savePaths%)
Declare Function report4screenBreaks% Lib "c4dll.dll" (ByVal r4&, ByVal Value%)
Declare Function report4separator% Lib "c4dll.dll" Alias "report4separator_v" (ByVal r4&, ByVal separator$)
Declare Function report4sortSet% Lib "c4dll.dll" (ByVal r4&, ByVal sortExpr$)
Declare Function report4toScreen% Lib "c4dll.dll" (ByVal r4&, ByVal toScreen%)


'===============================================================================================
'
'                            Tag Functions' Prototypes
'
'-----------------------------------------------------------------------------------------------

Declare Function t4aliasCB& Lib "c4dll.dll" Alias "t4alias" (ByVal t4&)
Declare Function t4close% Lib "c4dll.dll" (ByVal t4&)
Declare Function t4descending% Lib "c4dll.dll" Alias "tfile4isDescending" (ByVal t4&)
Declare Function t4exprCB& Lib "c4dll.dll" (ByVal t4&)
Declare Function t4filterCB& Lib "c4dll.dll" (ByVal t4&)
Declare Function t4open& Lib "c4dll.dll" Alias "t4openCB" (ByVal dbPtr&, ByVal IndexName$)
Declare Function t4unique% Lib "c4dll.dll" (ByVal t4&)
Declare Function t4uniqueSet% Lib "c4dll.dll" Alias "t4uniqueSetVB" (ByVal t4&, ByVal Value%)


'=======================================================================================
'
'                Utility function prototypes
'
'-----------------------------------------------------------------------------------------------

Declare Function u4alloc& Lib "c4dll.dll" Alias "u4allocDefault" (ByVal amt&)
Declare Function u4allocFree& Lib "c4dll.dll" Alias "u4allocFreeDefault" (ByVal c4&, ByVal amt&)
Declare Sub u4free Lib "c4dll.dll" Alias "u4freeDefault" (ByVal MemPtr&)

'16-Bit versions
Declare Function u4ncpy% Lib "c4dll.dll" (ByVal MemPtr1$, ByVal memptr2&, ByVal memLength%)
Declare Function u4ncpy2% Lib "c4dll.dll" Alias "u4ncpy" (ByVal MemPtr1&, ByVal memptr2$, ByVal memLength%)

'32-Bit versions
'Declare Function u4ncpy& Lib "c4dll.dll" (ByVal MemPtr1$, ByVal memptr2&, ByVal memLength&)
'Declare Function u4ncpy2& Lib "c4dll.dll" Alias "u4ncpy" (ByVal MemPtr1&, ByVal memptr2$, ByVal memLength&)

Declare Sub u4memCpy Lib "c4dll.dll" (ByVal dest$, ByVal Source&, ByVal numCopy&)
Declare Function u4switch& Lib "c4dll.dll" ()

'=======================================================================================
'
'                Misc. function prototypes
'
'========================================================================================

Declare Function v4Cstring& Lib "c4dll.dll" (ByVal s$)
Declare Sub v4Cstringfree Lib "c4dll.dll" (ByVal s&)

'CodeBase Return Code Constants

Global Const r4success% = 0
Global Const r4same = 0
Global Const r4found% = 1
Global Const r4down = 1
Global Const r4after = 2
Global Const r4complete = 2
Global Const r4eof = 3
Global Const r4bof = 4
Global Const r4entry = 5
Global Const r4descending = 10
Global Const r4unique = 20
Global Const r4uniqueContinue = 25
Global Const r4locked = 50
Global Const r4noCreate = 60
Global Const r4noOpen = 70
Global Const r4notag = 80
Global Const r4terminate = 90
Global Const r4inactive = 110
Global Const r4active = 120
Global Const r4authorize = 140
Global Const r4connected = 150
Global Const r4logOpen = 170
Global Const r4logOff = 180
Global Const r4null = 190

Global Const relate4filterRecord = 101
Global Const relate4doRemove = 102
Global Const relate4skipped = 104
Global Const relate4blank = 105
Global Const relate4skipRec = 106
Global Const relate4terminate = 107
Global Const relate4exact = 108
Global Const relate4scan = 109
Global Const relate4approx = 110
Global Const relate4sortSkip = 120
Global Const relate4sortDone = 121

'CodeBasic Field Definition Constants
Global Const r4logLen = 1
Global Const r4dateLen = 8
Global Const r4memoLen = 10
Global Const r4bin = "B"        ' Binary
Global Const r4str$ = "C"       ' Character
Global Const r4charBin$ = "Z"   ' Character (binary)
Global Const r4currency$ = "Y"  ' Currency
Global Const r4date$ = "D"      ' Date
Global Const r4dateTime$ = "T"  ' DateTime
Global Const r4double$ = "B"    ' Double
Global Const r4float$ = "F"     ' Float
Global Const r4gen$ = "G"       ' General
Global Const r4int$ = "I"       ' Integer
Global Const r4log$ = "L"       ' Logical
Global Const r4memo$ = "M"      ' Memo
Global Const r4memoBin$ = "X"   ' Memo (binary)
Global Const r4num$ = "N"       ' Numeric
Global Const r4dateDoub$ = "d"  ' Date as Double
Global Const r4numDoub$ = "n"   ' Numeric as Double
Global Const r4unicode$ = "n"   ' Unicode character

Global Const r4system$ = "0"    ' used by FoxPro for null field value field
Global Const r5wstrLen$ = "O"
Global Const r5ui4$ = "P"
Global Const r5i2$ = "Q"
Global Const r5ui2$ = "R"
Global Const r5guid$ = "V"
Global Const r5wstr$ = "W"
Global Const r5i8$ = "1"        ' 8-byte long signed value (LONGLONG)
Global Const r5dbDate$ = "2"    ' struct DBDATE (6 bytes)
Global Const r5dbTime$ = "3"    ' struct DBTIME (6 bytes)
Global Const r5dbTimeStamp$ = "4" ' struct DBTIMESTAMP (16 bytes)
Global Const r5date$ = "5"

'Other CodeBase Constants
Global Const cp0 = 0          'code4.codePage constant
Global Const cp437 = 1
Global Const cp1252 = 3
Global Const LOCK4OFF = 0
Global Const LOCK4ALL = 1
Global Const LOCK4DATA = 2
Global Const LOG4TRANS = 0
Global Const LOG4ON = 1
Global Const LOG4ALWAYS = 2
Global Const OPEN4DENY_NONE = 0
Global Const OPEN4DENY_RW = 1
Global Const OPEN4DENY_WRITE = 2
Global Const OPT4EXCLUSIVE = -1
Global Const OPT4OFF = 0
Global Const OPT4ALL = 1
Global Const r4check = -5
Global Const r4maxVBStringLen = 65500
Global Const r4maxVBStrFunction = 32767
Global Const collate4machine = 1
Global Const collate4general = 1001
Global Const collate4special = 1002
Global Const sort4machine = 0 'code4.collatingSequence constant
Global Const sort4general = 1
Global Const WAIT4EVER = -1


'CodeBasic Error Code Constants
Global Const e4close = -10
Global Const e4create = -20
Global Const e4len = -30
Global Const e4lenSet = -40
Global Const e4lock = -50
Global Const e4open = -60
Global Const e4permiss = -61
Global Const e4access = -62
Global Const e4numFiles = -63
Global Const e4fileFind = -64
Global Const e4instance = -69
Global Const e4read = -70
Global Const e4remove = -80
Global Const e4rename = -90
Global Const e4seek = -250
Global Const e4unlock = -110
Global Const e4write = -120
Global Const e4data = -200
Global Const e4fieldName = -210
Global Const e4fieldType = -220
Global Const e4recordLen = -230
Global Const e4append = -240
Global Const e4entry = -300
Global Const e4index = -310
Global Const e4tagName = -330
Global Const e4unique = -340
Global Const e4tagInfo = -350
Global Const e4commaExpected = -400
Global Const e4complete = -410
Global Const e4dataName = -420
Global Const e4lengthErr = -422
Global Const e4notConstant = -425
Global Const e4numParms = -430
Global Const e4overflow = -440
Global Const e4rightMissing = -450
Global Const e4typeSub = -460
Global Const e4unrecFunction = -470
Global Const e4unrecOperator = -480
Global Const e4unrecValue = -490
Global Const e4undetermined = -500
Global Const e4tagExpr = -510
Global Const e4opt = -610
Global Const e4optSuspend = -620
Global Const e4optFlush = -630
Global Const e4relate = -710
Global Const e4lookupErr = -720
Global Const e4relateRefer = -730
Global Const e4info = -910
Global Const e4memory = -920
Global Const e4parm = -930
Global Const e4parmNull = -935
Global Const e4demo = -940
Global Const e4result = -950
Global Const e4verify = -960
Global Const e4struct = -970
Global Const e4notSupported = -1090
Global Const e4version = -1095
Global Const e4memoCorrupt = -1110
Global Const e4memoCreate = -1120
Global Const e4transViolation = -1200
Global Const e4trans = -1210
Global Const e4rollback = -1220
Global Const e4commit = -1230
Global Const e4transAppend = -1240
Global Const e4corrupt = -1300
Global Const e4connection = -1310
Global Const e4socket = -1320
Global Const e4net = -1330
Global Const e4loadlib = -1340
Global Const e4timeOut = -1350
Global Const e4message = -1360
Global Const e4packetLen = -1370
Global Const e4packet = -1380
Global Const e4max = -1400
Global Const e4codeBase = -1410
Global Const e4name = -1420
Global Const e4authorize = -1430
Global Const e4server = -2100
Global Const e4config = -2110
Global Const e4cat = -2120

'ADO Constants
Global Const e5badBinding = 200
Global Const e5conversion = 210
Global Const e5delete = 220

'CodeControls Constants
Global Const CB_TOP = 1
Global Const CB_PREV = 2
Global Const CB_SEARCH = 3
Global Const CB_NEXT = 4
Global Const CB_BOTTOM = 5
Global Const CB_APPEND = 6
Global Const CB_DEL = 7
Global Const CB_UNDO = 8
Global Const CB_FLUSH = 9
Global Const CB_GO = 10

'=======================================================================================
'
'                CodeControls function prototypes
'
'========================================================================================


'CodeControls Constants

Global Const MASTER4NODATA% = 1
Global Const MASTER4NOTAG% = 2
Global Const MASTER4BADEXPR% = 3
Global Const CTRL4BADFIELD% = 4
Global Const CTRL4NOTAG% = 5
Global Const CTRL4BADEXPR% = 6

Function b4String$(p&)
    'This is a utility function for copying a 'C' string to a VB string.

    Dim s As String * 256
    Dim rc As Integer

    s$ = ""

    If p& <> 0 Then
       rc = u4ncpy(s, p, 256)
    End If

    b4String$ = Left$(s, rc)
End Function

Function code4dateFormat$(c4Ptr&)
    'This function returns the CODE4.dateFormat member
    code4dateFormat = b4String(code4dateFormatVB(c4Ptr&))
End Function

Function code4indexExtension$(c4Ptr&)
    'This function returns the CodeBase DLL index format
    code4indexExtension = b4String(code4indexExtensionVB(c4Ptr&))
End Function

Function code4lockFileName$(c4Ptr&)
'This function returns the locked file name
    code4lockFileName = b4String(code4lockFileNameVB(c4Ptr&))
End Function

Function code4lockNetworkId$(c4Ptr&)
    'This function returns the user's network id
    'who has locked the current file
    code4lockNetworkId = b4String(code4lockNetworkIdVB(c4Ptr&))
End Function

Function code4lockUserId$(c4Ptr&)
    'This function returns the user's name
    'who has locked the current file
    code4lockUserId = b4String(code4lockUserIdVB(c4Ptr&))
End Function

Function code4logFileName$(c4Ptr&)
    'This function returns the locked file name
    code4logFileName = b4String(code4lockFileNameVB(c4Ptr&))
End Function

Function d4alias$(dbPtr&)
    'This function returns the data file alias
   d4alias = b4String(d4aliasCB(dbPtr))
End Function

Function d4create&(ByVal cb&, dbname$, d() As FIELD4INFO, n() As TAG4INFO)

    ' d4create calls d4createCB() to create a new database.
    ' This function is the same as d4createData() except that
    ' it requires an additional parameter which it uses to
    ' create tag information for a database.
    '
    ' Variable n is an array of type TAG4INFO which corresponds
    ' to TAG4INFOCB, a structure that can be used by d4create.
    ' The difference once again is merely the difference in the
    ' representation of strings between C and Basic.
    ' d4create takes the contents from the TAG4INFO structure
    ' and builds a TAG4INFOCB structure which it passes to d4createCB().
    ' Note: the TAG4INFOCB array is one size larger than the TAG4INFO
    ' array.  The extra empty (zero filled) array element is the
    ' way that d4createCB() detects the end of the array.


    Dim i%

    Dim flb%
    Dim fub%
    Dim fs%

    Dim tlb%
    Dim tub%
    Dim ts%


    flb = LBound(d)
    fub = UBound(d)
    fs = fub - flb + 1


    ReDim F(1 To (fs + 1)) As FIELD4INFOCB
    For i = 1 To fs
        F(i).fName = v4Cstring(d((flb - 1) + i).fName)    ' note: this function allocates memory
        F(i).ftype = Asc(d((flb - 1) + i).ftype)
        F(i).flength = d((flb - 1) + i).flength
        F(i).fdecimals = d((flb - 1) + i).fdecimals
        F(i).fnulls = d((flb - 1) + i).fnulls
    Next i

    tlb = LBound(n)
    tub = UBound(n)
    ts = tub - tlb + 1
    ReDim t(1 To (ts + 1)) As TAG4INFOCB
    For i = 1 To ts
        t(i).Name = v4Cstring(n((tlb - 1) + i).Name)
        t(i).Expression = v4Cstring(n((tlb - 1) + i).Expression)
        t(i).Filter = v4Cstring(n((tlb - 1) + i).Filter)
        t(i).unique = n((tlb - 1) + i).unique
        t(i).descending = n((tlb - 1) + i).descending
    Next i
    d4create = d4createCB(cb&, ByVal (dbname$), F(1), t(1))

    ' Since v4Cstring allocates memory for the storage of
    ' C strings, we must free the memory after it has been
    ' used.
    For i = 1 To fs
         Call v4Cstringfree(F(i).fName)
    Next i

    For i = 1 To ts
        Call v4Cstringfree(t(i).Name)
        Call v4Cstringfree(t(i).Expression)
        Call v4Cstringfree(t(i).Filter)
    Next i

End Function

Function d4createData&(ByVal cb&, dbname$, d() As FIELD4INFO)

    ' d4createData() calls d4createCB() to create a new database.
    ' d4create() builds the FIELD4INFOCB array which is
    ' the one recognized by d4create (note that the only difference
    ' is that the fname field is a string in type FIELD4INFO
    ' and type long in FIELD4INFOCB which is how strings are represented
    ' in C).  Furthermore, the size of f (our FIELD4INFOCB array) is one
    ' larger than the size s of FIELD4INFO d.  This is because
    ' d4create doesn't know the size of the array f and therefore it stops
    ' when it reaches an array element that is filled with zeros which
    ' the extra (s+1)'th element of f provides.

    Dim i%
    Dim lb%
    Dim ub%
    Dim s%

    lb = LBound(d)
    ub = UBound(d)
    s = ub - lb + 1


    ReDim F(1 To (s + 1)) As FIELD4INFOCB
    For i = 1 To s
        F(i).fName = v4Cstring(d((lb - 1) + i).fName) ' note: this function allocates memory
        F(i).ftype = Asc(d((lb - 1) + i).ftype)
        F(i).flength = d((lb - 1) + i).flength
        F(i).fdecimals = d((lb - 1) + i).fdecimals
        F(i).fnulls = d((lb - 1) + i).fnulls
    Next i

    d4createData = d4createCB(cb&, ByVal (dbname$), F(1), ByVal (0&))

    ' Since v4Cstring allocates memory for the storage of
    ' C strings, we must free the memory after it has been
    ' used.
    For i = 1 To s
      Call v4Cstringfree(F(i).fName)
    Next i

End Function

Function d4encodeHandle(temp As Long)
    Dim EncodedString As String
    EncodedString = "#" + Str$(temp)
    d4encodeHandle = EncodedString
End Function

Function d4fileName$(dbfPtr&)
   d4fileName$ = b4String(d4fileNameCB(dbfPtr))
End Function

Function d4go%(DATA4&, recordNumber&)
   d4go = d4goLow(DATA4, recordNumber, 1)
End Function

Sub date4assign(dateString$, julianDay&)
'This functions converts the julian day into standard format
'and puts the result in dateString

   'Size dateString$
   dateString$ = Space$(8 + 1)

   Call date4assignLow(dateString, julianDay, 0)
   dateString$ = Left$(dateString$, 8)
End Sub

Function date4cdow$(dateString$)

'This function returns the day of the week in a character
'string based on the value in 'DateString'

    'Validate "dateString"
   If dateString = "" Or Len(dateString) < 8 Then Exit Function

   Dim datePtr&
   datePtr& = date4cdowCB(dateString)    'Get pointer to day
   If datePtr = 0 Then Exit Function     'Illegal date

   date4cdow = b4String(datePtr)

End Function

Function date4cmonth$(dateString$)

'This function returns the month in 'DateString' as a
'character string

   'Validate "DateString"
   If dateString = "" Or Len(dateString) < 8 Then Exit Function

   Dim datePtr&
   datePtr& = date4cmonthCB(dateString)    'Get pointer to month
   If datePtr = 0 Then Exit Function       'Illegal date

   date4cmonth = b4String(datePtr)                 'Return month name

End Function

Sub date4format(dateString$, Result$, pic$)

'This functions formats Result$ using the date value
' in 'dateString$' and the format info. in 'Pic$'

   'Size Result$
   Result$ = Space$(Len(pic$) + 1)

   Call date4formatCB(dateString$, Result$, pic$)
   Result$ = Left$(Result$, Len(pic$))

End Sub

Sub date4init(Result$, dateString$, pic$)

'This functions copies the date, specified by dateString,
'and formatted according to pic, into Result. The date copied
'will be in standard dBASE format,

   'Size Result$
   Result$ = Space$(9)

   Call date4initCB(Result$, dateString$, pic$)
   Result$ = Left$(Result$, 8)

End Sub

Sub date4today(dateS As String)
   If Len(dateS) < 8 Then dateS = Space$(8)
   Call date4todayCB(dateS)
End Sub

Function error4text$(c4&, errCode&)
    'This function returns the error message string
    error4text = b4String(error4textCB(c4, errCode))
End Function

Function expr4null%(exPtr&)
    expr4null = expr4nullLow(exPtr, 1)
End Function

Function expr4source$(exPtr&)
    'This function returns a copy of the original
    'dBASE expression string
    expr4source = b4String(expr4sourceCB(exPtr))
End Function

Function expr4str$(exPtr&)
'This function returns the parsed string

   Dim exprPtr&
   Dim buf As String

   'Get pointer to alias string
   exprPtr& = expr4strCB(exPtr)

   If exprPtr& = 0 Then Exit Function

   expr4str = Left$(b4String(exprPtr), expr4len(exPtr))
End Function

Function expr4type$(exPtr&)
'This function returns the type of the parsed string

   Dim exprType%

   'Get ASCII value of type
   exprType = expr4typeCB(exPtr)

   If exprType = 0 Then Exit Function

   expr4type = Chr$(exprType)

End Function

Sub f4assign(fPtr As Long, fStr As String)
   Call f4assignN(fPtr, fStr, Len(fStr))
End Sub

Sub f4assignInt(fldPtr&, fldVal%)
   Call f4assignIntVB(fldPtr, fldVal)
End Sub

Function f4currency$(field&, numDec%)
    'This function returns the contents of a field
    f4currency = b4String(f4currencyCB(field, numDec))
End Function

Function f4dateTime$(field&)
    'This function returns the contents of a field
    f4dateTime = b4String(f4dateTimeCB(field))
End Function

Function f4memoStr$(fPtr&)
'This function returns a string corresponding to the memo
'field pointer argument.

    Dim r4line$
    r4line = Chr$(10) + Chr$(13)

    Dim MemoLen&, MemoPtr&

    MemoLen& = f4memoLen(fPtr)          'Get memo length

    If MemoLen > &H7FFFFFFF Then
        MsgBox "Error #: -910" + r4line + "Unexpected Information" + r4line + "Memo entry too long to return in a Visual Basic string." + r4line + "Field Name:" + r4line + f4name(fPtr), 16, "CodeBase Error"
        Exit Function
    End If

    MemoPtr& = f4memoPtr(fPtr)
    If MemoPtr& = 0 Then Exit Function

    Dim MemoString$
    MemoString = String$(MemoLen&, " ")

    'Copy 'MemoPtr' to VB string 'MemoString'
    u4memCpy MemoString, MemoPtr, MemoLen

    f4memoStr = MemoString

End Function

Sub f4memoStr64(fPtr As Long, src As String)

    'This function copies a large memo entry (32K-64K)
    'into a user supplied string

    Dim r4line$
    r4line = Chr$(10) + Chr$(13)

    Dim MemoLen&, MemoPtr&

    MemoLen& = f4memoLen(fPtr)          'Get memo length

    ' 'r4maxVBStringLen' defined in 'Constants' section of this file
    If MemoLen > r4maxVBStringLen Then
        MsgBox "Error #: -910" + r4line + "Unexpected Information" + r4line + "Memo entry too long to retrieve into VB string." + r4line + "Field Name:" + r4line + f4name(fPtr), 16, "CodeBasic Error"
        Exit Sub
    End If

    MemoPtr& = f4memoPtr(fPtr)
    If MemoPtr& = 0 Then Exit Sub

    src = String$(MemoLen&, " ")

    'Copy 'MemoPtr' to VB string 'src'
    u4memCpy src, MemoPtr, MemoLen

End Sub

Function f4name$(fPtr&)

'This function returns the name of a field

    Dim FldNamePtr&                     'Pointer to field name
    Dim FldName As String * 11          'String to hold info
    FldNamePtr& = f4nameCB(fPtr)        'Get pointer

         f4name = b4String(FldNamePtr)

End Function

Function f4nCpy(field&, s$, slen%)
    'This function copies the fields contents into a string
   Dim fPtr&

    s = Space$(slen)                    'Make s$ one byte longer for null character that u4ncpy adds
   fPtr& = f4ptr(field)

   If fPtr& = 0 Then Exit Function

   u4memCpy s, fPtr, slen

   f4nCpy = Len(s)

End Function

Function f4str$(field&)
    'This function returns the contents of a field
    Dim s$, fPtr&, flen%

    fPtr& = f4ptr(field)
    If fPtr& = 0 Then Exit Function

    flen = f4len(field)                 'Get field length
    s = Space$(flen)                    'Make s$ one byte longer for null character that u4ncpy adds

    u4memCpy s, fPtr, flen

    f4str = s
End Function

Function i4create&(ByVal dbPtr&, IndexName$, n() As TAG4INFO)

    ' i4create() calls i4createCB() to create a new
    ' index file. Variable n is an array of type TAG4INFO
    ' which corresponds to TAG4INFOCB, a structure that
    ' can be used by i4createCB(). The difference once
    ' again is merely the difference in the representation
    ' of strings between C and Basic.
    '
    ' i4create() takes the contents from the TAG4INFO
    ' structure and builds a TAG4INFOCB structure which
    ' it passes to i4createCB(). Note: the TAG4INFOCB
    ' arrary is one size larger than the TAG4INFO array.
    ' The extra empty (zero filled) array element is the
    ' way that i4create detects the end of the array.
    '
    ' Note also, that if 'IndexName' is an empty string,
    ' the index file that is created will become a
    ' "production" index file. i.e. it will be opened every
    ' time the corresponding data file is opened.

    Dim i%

    Dim tlb%
    Dim tub%
    Dim ts%


    tlb = LBound(n)
    tub = UBound(n)
    ts = tub - tlb + 1
    ReDim t(1 To (ts + 1)) As TAG4INFOCB
    For i = 1 To ts
        t(i).Name = v4Cstring(n((tlb - 1) + i).Name)
        t(i).Expression = v4Cstring(n((tlb - 1) + i).Expression)
        t(i).Filter = v4Cstring(n((tlb - 1) + i).Filter)
        t(i).unique = n((tlb - 1) + i).unique
        t(i).descending = n((tlb - 1) + i).descending
    Next i

    If IndexName$ = "" Then    'User wants production index file
        i4create = i4createCB(dbPtr&, ByVal 0&, t(1))
    Else
        i4create = i4createCB(dbPtr&, IndexName$, t(1))
    End If

    ' Since v4Cstring allocates memory for the storage of
    ' C strings, we must free the memory after it has been
    ' used.

    For i = 1 To ts
         Call v4Cstringfree(t(i).Name)
         Call v4Cstringfree(t(i).Expression)
         Call v4Cstringfree(t(i).Filter)
    Next i


End Function

Function i4fileName$(iPtr&)
    'This function returns the file name of an index tag
    i4fileName = b4String(i4fileNameCB(iPtr))
End Function

Function i4open&(d4&, fName$)
   If fName = "" Then
      i4open = i4openCB(d4&, ByVal 0&)   'Use data file name
   Else
      i4open = i4openCB(d4&, fName$)     'Use supplied name
   End If
End Function

Function i4tagAdd%(ByVal i4Ptr&, n() As TAG4INFO)
    ' i4tagAdd adds additional tags to an existing
    ' index.

    ' i4tagAdd takes the contents from the TAG4INFO
    ' structure and builds a TAG4INFOCB structure which
    ' is passed to i4tagAddCB.

    Dim i%

    Dim tlb%
    Dim tub%
    Dim ts%


    tlb = LBound(n)
    tub = UBound(n)
    ts = tub - tlb + 1
    ReDim t(1 To (ts + 1)) As TAG4INFOCB
    For i = 1 To ts
        t(i).Name = v4Cstring(n((tlb - 1) + i).Name)
        t(i).Expression = v4Cstring(n((tlb - 1) + i).Expression)
        t(i).Filter = v4Cstring(n((tlb - 1) + i).Filter)
        t(i).unique = n((tlb - 1) + i).unique
        t(i).descending = n((tlb - 1) + i).descending
    Next i

    i4tagAdd = i4tagAddCB(i4Ptr&, t(1))

End Function

Function relate4masterExpr$(r4Ptr&)
    'This function returns the Relations expression string
    relate4masterExpr = b4String(relate4masterExprCB(r4Ptr&))
End Function

Function report4parent%(ByVal r4&, ByVal hWnd&)
   #If Win16 Then
      report4parent = report4parent16(r4, hWnd)
   #End If
   #If Win32 Then
      report4parent = report4parent32(r4, hWnd)
   #End If
End Function

Sub report4free(pReport&, freeRelate%, closeFiles%)
    Call report4freeLow(pReport, freeRelate, closeFiles, 1)
End Sub

Function t4Alias$(tPtr&)
    t4Alias = b4String(t4aliasCB(tPtr))
End Function

Function t4expr$(tPtr&)
    'This function returns the original tag expression
    t4expr = b4String(t4exprCB(tPtr))
End Function

Function t4filter$(tPtr&)

'This function returns the tag filter expression

    Dim FilterPtr&
    Dim Filter As String * 255

    'Get pointer to parsed filter expression
    FilterPtr& = t4filterCB(tPtr&)

    If FilterPtr& = 0 Then
        t4filter = ""
        Exit Function         'No filter
    End If

   t4filter = b4String(FilterPtr)

End Function

Function u4descend$(charString$)
   Dim Result$, i%

   For i = 1 To Len(charString)
      Result = Result + Chr$(128 And Asc(Mid$(charString, i, 1)))
   Next

   u4descend = Result
End Function

