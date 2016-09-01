VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLibrarys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cLibrarys.cls
'' Description: Handle a collection of cLibrary objects
''
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Compare Text
Option Explicit

Private Type mPrivate
    Librarys As Collection
End Type
Private m As mPrivate

Private Const kClassName = "Class: cLibrarys "

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Found
'' Description: Determine whether a given ID is in the collection
'' Inputs:      ID to find
'' Returns:     True if Found, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Property Get Found(pID As Long) As Boolean
On Error GoTo ErrSection:

    Dim Library As cLibrary
    
    Found = True
    Set Library = Item(Str(pID))

ErrExit:
    Exit Property

ErrSection:
    If Err.Number = 91 Or Err.Number = 5 Then
        Found = False
        Resume Next
    Else
        RaiseError "cLibrarys.Found", eGDRaiseError_Raise, g.strAppPath
        Resume ErrExit
    End If

End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Load
'' Description: Load up the collection with Librarys from the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Load()
On Error GoTo ErrSection:
    
    Dim rs As Recordset                 ' Recordset into the database
    Dim Library As cLibrary            ' Temporary library object
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblLibrarys] " & "ORDER BY [LibraryName];", dbOpenSnapshot)
    If Not (rs.BOF And rs.EOF) Then rs.MoveFirst
    Do Until rs.EOF
        If rs!CheckSum = BuildCheckSum(rs, "tblLibrarys") Then
            Set Library = New cLibrary
            With Library
                .LibraryID = rs!LibraryID
                .LibraryName = rs!LibraryName
                .LibraryDesc = rs!LibraryDesc
                .SecurityLevel = NullChk(rs!SecurityLevel, 0)
                .CannotDelete = rs!CannotDelete
                .Password = DecryptField(rs!Password)
                .Version = rs!Version
                .LastModified = rs!LastModified
                .Author = rs!Author
                .RtfFileName = rs!RtfFileName
                .LibraryType = NullChk(rs!LibraryType, 0)
                .BuiltIn = rs!BuiltIn
                .RequiredMod = FixRequiredMod(NullChk(rs!RequiredMod))
                .Ignore = rs!Ignore
            End With
            m.Librarys.Add Library, Str(rs!LibraryID)
        End If
        
        rs.MoveNext
    Loop
    
ErrExit:
    Set rs = Nothing
    Set Library = Nothing
    Exit Sub

ErrSection:
    RaiseError "cLibrarys.Load", eGDRaiseError_Raise, g.strAppPath
    Resume ErrExit

End Sub

Property Get Item(ByVal pKey As Variant) As cLibrary
    Set Item = m.Librarys(pKey)
End Property

Property Get Count() As Long
    Count = m.Librarys.Count
End Property

Public Sub Remove(ByVal pLibraryID As Variant)
    m.Librarys.Remove pLibraryID
End Sub

Private Sub Class_Initialize()
    Set m.Librarys = New Collection
End Sub

Private Sub Class_Terminate()
    Set m.Librarys = Nothing
End Sub
