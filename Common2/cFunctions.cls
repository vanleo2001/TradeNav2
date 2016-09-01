VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Option Compare Text

Private Type mPrivate
    Functions As Collection
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Load
'' Description: Load functions from tblFunctions into cFunctions collection
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Load()
On Error GoTo ErrSection:
    
    Dim rs As Recordset                 ' Recordset into the database
    Dim rsParms As Recordset            ' Recordset into the database
    Dim Func As cFunction               ' Temporary Function object
    Dim nBarParms As Byte               ' Number of Bar Parameters
    
    Set rs = g.dbNav.OpenRecordset("SELECT tblFunctions.*, tblLibrarys.*, tblFunctionCategories.FunctionCategory " & _
                "FROM tblFunctionCategories INNER JOIN (tblLibrarys INNER JOIN tblFunctions ON tblLibrarys.LibraryID = tblFunctions.LibraryID) ON tblFunctionCategories.FunctionCategoryID = tblFunctions.FunctionCategoryID " & _
                "WHERE ((tblLibrarys.Ignore) = 0);", dbOpenDynaset)

    ValidateCheckSums rs, "tblFunctions"
    ValidateCheckSums rs, "tblLibrarys"
    If Not (rs.BOF And rs.EOF) Then rs.MoveFirst
                
    Do Until rs.EOF
        If rs![tblFunctions.CheckSum] <> 0.5 And rs![tblLibrarys.CheckSum] <> 0.5 Then
            Set Func = New cFunction
            With Func
                .FunctionID = rs!FunctionID
                .FunctionName = rs!FunctionName
                .Description = rs!Description
                .FunctionCategoryID = rs!FunctionCategoryID
                .ImplementationTypeID = rs!ImplementationTypeID
                .LateCalculating = rs!LateCalculating
                .LastModified = rs![tblFunctions.LastModified]
                .CodedName = rs!CodedName
                .Formatted = DecryptField(rs!Formatted)
                .FormattedWithFillWords = DecryptField(rs!FormattedWithFillWords)
                .CodedText = DecryptField(rs!CodedText)
                .Usage = rs!Usage
                .DataTypeID = rs!ReturnTypeID
                .LibraryID = rs![tblFunctions.LibraryID]
                .SecurityLevel = NullChk(rs![tblFunctions.SecurityLevel], 0)
                .Password = DecryptField(rs![tblFunctions.Password])
                .CannotDelete = rs![tblFunctions.CannotDelete]
                .TradeSenseUsage = NullChk(rs!TradeSenseUsage)
                .Reverify = rs!Reverify
                .Abbrev = NullChk(rs!Abbrev)
                .UsesNextBarHLC = rs!UsesHLCNextBar
                
                .RequiredMod = NullChk(rs![tblFunctions.RequiredMod])
                
                'MT 6/2001: TradeSense compatibility =======================
                .ReturnTypeID = rs!ReturnTypeID
                .ListID = rs!ListID
                .DLLName = ""
                .ClassName = rs!CodedName
                .Preview = DecryptField(rs!FormattedWithFillWords)
                'MT 6/2001: TradeSense compatibility =======================
                            
                ' Load parameters into the collection of inputs for this function
                nBarParms = 0
                Set rsParms = g.dbNav.OpenRecordset("SELECT tblFunctionParms.*, tblParmTypes.ParmUsage " & _
                            "FROM tblParmTypes INNER JOIN tblFunctionParms ON tblParmTypes.ParmTypeID = tblFunctionParms.ParmTypeID " & _
                            "WHERE (((tblFunctionParms.FunctionID)=" & Str(.FunctionID) & "))" & _
                            "ORDER BY [ParmNbr];", dbOpenSnapshot)
    
                Do Until rsParms.EOF
                    .LoadInput rsParms!ParmID, rsParms!ParmNbr, rsParms!ParmText, rsParms!ParmDesc, rsParms!ParmTypeID, _
                        rsParms!ListID, rsParms!Required, rsParms!ParmUsage, rsParms!ToValue, rsParms!DefaultValue, _
                        rsParms!FillPre, rsParms!FillPost
                    'Internal Data structure (3)
                    If rsParms!ParmUsage = 3 Then nBarParms = nBarParms + 1
                    rsParms.MoveNext
                Loop
                .BarParms = nBarParms
                
                'MT 6/2001: TradeSense compatibility =======================
                .HiddenInputs = nBarParms
                'MT 6/2001: TradeSense compatibility =======================
                            
                m.Functions.Add Func, Str(.FunctionID)
            End With
        End If
        
        rs.MoveNext
    Loop
    rs.Close
    
ErrExit:
    Set rs = Nothing
    Set rsParms = Nothing
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    Set rsParms = Nothing
    RaiseError "cFunctions.Load", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Add
'' Description: Add a function to the collection
'' Inputs:      Function to add, ID for the Function
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Add(pFunction As cFunction, ByVal lFunctionID As Long)
On Error GoTo ErrSection:

    m.Functions.Add pFunction, Str(lFunctionID)
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cFunctions.Add", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    NameFound
'' Description: Search for function name
'' Inputs:      Name to find, ID of the found item
'' Returns:     True if Found, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NameFound(strName As String, lID As Long) As Boolean
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    ' Search by function name
    NameFound = False
    For lIndex = 1 To m.Functions.Count
        With Item(lIndex)
            If .FunctionName = strName Then
                lID = .FunctionID
                NameFound = True
                Exit Function
            End If
        End With
    Next lIndex
    
    ' Search by abbreviation
    For lIndex = 1 To m.Functions.Count
        With Item(lIndex)
            If .Abbrev = strName Then
                lID = .FunctionID
                NameFound = True
                Exit Function
            End If
        End With
    Next lIndex
    
ErrExit:
    Exit Function

ErrSection:
    RaiseError "cFunctions.NameFound", eGDRaiseError_Raise, g.strAppPath

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Found
'' Description: Find an item in the collection for a given key
'' Inputs:      Key to find
'' Returns:     True if Found, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Get Found(ByVal vKey As Variant) As Boolean
On Error GoTo ErrSection:

    Dim TmpFunction As cFunction        ' Temporary function object
    
    Found = True
    Set TmpFunction = m.Functions(vKey)

ErrExit:
    Exit Property
    
ErrSection:
    If Err.Number = 91 Or Err.Number = 5 Then
        Found = False
        Resume Next
    Else
        RaiseError "cFunctions.Found", eGDRaiseError_Raise, g.strAppPath
    End If

End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UpdateTradeSenseUsage
'' Description: Method to update the TradeSenseUsage field for all functions
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub UpdateTradeSenseUsage()
On Error GoTo ErrSection:

    Dim strTradeSenseText As String     ' Acutal TradeSense text
    Dim lLocOfLeftParen As Integer      ' Location of the Left Paren in the string
    Dim lIndex As Integer               ' Index into a for loop
    Dim bInputs As Byte                 ' Number of inputs for the function
    Dim lIndex2 As Integer              ' Index into a for loop
        
    For lIndex = 1 To m.Functions.Count
        With m.Functions.Item(lIndex)
            'Default the function name to beginning of TradeSense text
            strTradeSenseText = .FunctionName & " ("
            lLocOfLeftParen = 0
            bInputs = 0
    
            If Not .Inputs Is Nothing Then
                For lIndex2 = 1 To .Inputs.Count
                    If .Inputs.Item(lIndex2).ParmTypeID <> C_RetTrades And _
                       .Inputs.Item(lIndex2).ParmTypeID <> C_RetBars Then
                        strTradeSenseText = strTradeSenseText & .Inputs.Item(lIndex2).ParmName & ", "
                        bInputs = bInputs + 1
                    End If
                Next lIndex2
            End If
        
            'Add parm right paren to end of string
            If bInputs > 0 Then
                strTradeSenseText = Left(strTradeSenseText, Len(strTradeSenseText) - 2) + ")"
            Else
                strTradeSenseText = Left(strTradeSenseText, Len(strTradeSenseText) - 2)
            End If
            
            .TradeSenseUsage = strTradeSenseText
            .Save
        End With
    Next lIndex
    
ErrExit:
    Exit Sub

ErrSection:
    RaiseError "cFunctions.UpdateTradeSenseUsage", eGDRaiseError_Raise, g.strAppPath

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Item
'' Description: Get an item from the collection
'' Inputs:      Key of the item to retrieve
'' Returns:     Item in the collection
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Get Item(ByVal vKey As Variant) As cFunction
    Set Item = m.Functions(vKey)
End Property

Property Let Item(ByVal vKey As Variant, F As cFunction)
    If Found(vKey) Then m.Functions.Remove vKey
    m.Functions.Add F, F.FunctionID
End Property

Public Sub ReloadFunction(ByVal vKey As Variant)
    If Found(vKey) Then m.Functions.Item(vKey).Load
End Sub

Public Sub Remove(ByVal vKey As Variant)
    If Found(vKey) Then m.Functions.Remove vKey
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Count
'' Description: Get a count of the items in the collection
'' Inputs:      None
'' Returns:     Number of items in the collection
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Get Count() As Long
    Count = m.Functions.Count
End Property

Private Sub Class_Initialize()
    Set m.Functions = New Collection
End Sub

Private Sub Class_Terminate()
    Set m.Functions = Nothing
End Sub
