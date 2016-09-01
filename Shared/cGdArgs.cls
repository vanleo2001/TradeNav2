VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGdArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const GET_NEXT_ARG = 0

' Return Code definitions
Const RC_OK = &H0               'normal termination
Const RC_PARM_COUNT = &H100     'Parameter count mismatch
Const RC_PARM_MATCH = &H200     'Parameter type mismatch
Const RC_PARM_INVALID = &H300   'Not a valid Parameter
Const RC_FUNCT_NONE = &H400     'Function not found
Const RC_PARM_OOB = &H500       'Function Parameter out of Bounds
Const RC_USER_ERROR = &HFF00    'Function specific error

Enum eGdArg_Type
    eGDARG_Empty = 0
    eGDARG_Double = 1
    eGDARG_PtrString = 2
    eGDARG_Boolean = 3
    eGDARG_PtrDblArray = 4
    eGDARG_PtrBars = 5
    eGDARG_Long = 6
    eGDARG_PtrLong = 7
    eGDARG_PtrDouble = 8
    eGDARG_PtrFunct = 9
    eGDARG_PtrLongArray = 10
    eGDARG_PtrStringArray = 11
    eGDARG_PtrAction = 12
    eGDARG_PtrVoid = 13
    eGDARG_PtrEntry = 14
End Enum

Private Type mPrivate
    hArgs As Long
    nCount As Long
    nError As Long
    strErrMsg As String
    nPrevArgNum As Long
End Type
Private m As mPrivate

Private Sub Class_Initialize()
    ArgsHandle = 0
End Sub

Public Property Get ArgsHandle() As Long
    ArgsHandle = m.hArgs
End Property

Public Property Let ArgsHandle(ByVal hArgs As Long)
    m.nError = 0
    m.strErrMsg = ""
    m.hArgs = hArgs
    If hArgs = 0 Then
        m.nCount = 0
    Else
        m.nCount = gdGetArgCount(m.hArgs)
    End If
    m.nPrevArgNum = 0
End Property

Public Property Get Count() As Long
    Count = m.nCount
End Property

Public Property Get FromBar() As Long
    FromBar = gdGetArgFromBar(m.hArgs)
End Property

Public Property Get ArgType(ByVal nArgNum&) As eGdArg_Type
    ArgType = gdGetArgType(m.hArgs, nArgNum)
End Property

Public Property Get Error() As Long
    Error = m.nError
End Property

Public Property Get ErrorMessage() As String
    ErrorMessage = m.strErrMsg
End Property

Private Sub SetError(ByVal nError As Long, ByVal strErrMsg$)
    'only store first error
    If m.nError = 0 Then
        m.nError = nError
        m.strErrMsg = strErrMsg
    End If
End Sub

Private Sub CheckArgNum(nArgNum As Long)
    'fix arg num
    If nArgNum = GET_NEXT_ARG Then
        nArgNum = m.nPrevArgNum + 1
    End If
    m.nPrevArgNum = nArgNum
    
    'don't bother checking if already an error
    If m.nError = 0 Then
        If nArgNum < 1 Or nArgNum > m.nCount Then
            SetError RC_PARM_COUNT Or (m.nCount And &HFF), _
                "Tried to process arg " _
                & Trim(Str(nArgNum)) & " when function is defined to have " _
                & Trim(Str(m.nCount)) & " args."
        End If
    End If
End Sub

Public Function GetArg(vArg As Variant, _
        Optional ByVal nArgNum As Long = GET_NEXT_ARG) As Boolean

    Dim bSuccess As Boolean, strType$
    Dim hHandle&, dNumber#
    
    On Error GoTo ArgExit
    
    CheckArgNum nArgNum
    If m.nError <> 0 Then GoTo ArgExit
    
    Select Case VarType(vArg)
        Case vbString:
            ' get gdString arg
            If gdGetArgAsHandle(m.hArgs, nArgNum, hHandle) Then
                strType = UCase(Chr(gdGetType(hHandle)))
                If strType = "$" Then
                    vArg = gdGetStr(hHandle, 0)
                    bSuccess = True
                End If
            End If
        
        Case vbObject:
            strType = Trim(UCase(TypeName(vArg)))
            If vArg Is Nothing Then
                ' object not initialized (forgot to Dim with "New"?)
                SetError RC_PARM_MATCH Or nArgNum, _
                    "Error in arg " & Trim(Str(nArgNum)) _
                    & ":  Object not initialized (forgot to Dim with 'New'?)."
            ElseIf strType = "CGDBARS" Then
                ' get gdBars arg
                If gdGetArgAsHandle(m.hArgs, nArgNum, hHandle) Then
                    strType = UCase(Chr(gdGetType(hHandle)))
                    If strType = "B" Then
                        vArg.SetBarsHandle hHandle, False
                        bSuccess = True
                    End If
                End If
            ElseIf strType = "CGDARRAY" Then
                ' get gdArray arg
                If gdGetArgAsHandle(m.hArgs, nArgNum, hHandle) Then
                    strType = UCase(Chr(gdGetType(hHandle)))
                    If InStr("DFLT$", strType) > 0 Then
                        vArg.SetArrayHandle hHandle, False
                        bSuccess = True
                    End If
                End If
            End If
        
        Case vbVariant, vbSingle, vbLong, vbInteger, _
         vbDouble, vbDate, vbCurrency, vbByte, vbBoolean:
            ' get numeric arg
            If gdGetArgAsNumber(m.hArgs, nArgNum, dNumber) Then
                If VarType(vArg) <> vbBoolean Then
                    vArg = dNumber
                ElseIf dNumber <> 0 And dNumber <> kNullData Then
                    vArg = True
                Else
                    vArg = False
                End If
                bSuccess = True
            End If
    End Select
    
ArgExit:
    On Error Resume Next
    If Not bSuccess Then
        SetError RC_PARM_MATCH Or nArgNum, _
            "Data type mismatch for arg " & Trim(Str(nArgNum))
        GetArg = False
    Else
        GetArg = True
    End If
End Function

#If 0 Then
Public Function GetArrayArg(Optional ByVal nArgNum As Long = GET_NEXT_ARG) As cGdArray

    Dim bSuccess As Boolean, strType$, hHandle&
    Dim gda As New cGdArray
    
    On Error GoTo ArgExit
    
    CheckArgNum nArgNum
    If m.nError Then GoTo ArgExit
    
    ' get gdBars arg
    If gdGetArgAsHandle(m.hArgs, nArgNum, hHandle) Then
        strType = UCase(Chr(gdGetType(hHandle)))
        If InStr("DFLT", strType) > 0 Then
            gda.SetArrayHandle hHandle, False
            bSuccess = True
        End If
    End If
    
ArgExit:
    On Error Resume Next
    If Not bSuccess Then ArgMismatch nArgNum
    Set GetArrayArg = gda
End Function

Private Sub ArgMismatch(ByVal nArgNum&)
    SetError RC_PARM_MATCH Or nArgNum, _
        "Data type mismatch for arg " & Trim(Str(nArgNum))
End Sub


Public Function ArgTest(ByVal hResults&) As cGdArray

    Dim bSuccess As Boolean, strType$, hHandle&
    Dim gda As New cGdArray
    
    On Error GoTo ArgExit
    hHandle = hResults
    
    gda.SetArrayHandle hHandle, False
    bSuccess = True
    
ArgExit:
    On Error Resume Next
    'If Not bSuccess Then ArgMismatch nArgNum
    Set ArgTest = gda
End Function
#End If


Public Property Get InstanceMemPtr() As Long
    InstanceMemPtr = gdGetArgInstanceMemPtr(m.hArgs)
End Property

Public Property Let InstanceMemPtr(ByVal hMemPtr As Long)
    gdSetArgInstanceMemPtr m.hArgs, hMemPtr
End Property

Public Function AddDrawingCommand(ByVal strDrawingCommand$) As Boolean

    Dim hArray&
    hArray = gdGetArgDrawingCommands(m.hArgs)
    If hArray <> 0 Then
        If gdInsertStr(hArray, strDrawingCommand, gdGetSize(hArray)) <> 0 Then
            AddDrawingCommand = True
        End If
    End If

End Function
