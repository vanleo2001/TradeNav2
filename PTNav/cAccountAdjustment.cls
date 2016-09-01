VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAccountAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cAccountAdjustment.cls
'' Description: Wrapper class for account adjustments
'' Author:      Genesis Financial Data Services
''              425 Wind Chime Pl
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    lAdjumtmentID As Long               ' Unique identifier for the adjustment
    lAccountID As Long                  ' Account ID for the adjustment
    strDescription As String            ' Description
    dTime As Double                     ' Date and time the adjustment was made
    dAmount As Double                   ' Amount of the adjustment
    strNotes As String                  ' Notes for the adjustment
End Type
Private m As mPrivate

Public Property Get AdjustmentID() As Long
    AdjustmentID = m.lAdjumtmentID
End Property

Public Property Get AccountID() As Long
    AccountID = m.lAccountID
End Property
Public Property Let AccountID(ByVal lAccountID As Long)
    m.lAccountID = lAccountID
End Property

Public Property Get Description() As String
    Description = m.strDescription
End Property
Public Property Let Description(ByVal strDescription As String)
    m.strDescription = strDescription
End Property

Public Property Get AdjustmentTime() As Double
    AdjustmentTime = m.dTime
End Property
Public Property Let AdjustmentTime(ByVal dTime As Double)
    m.dTime = dTime
End Property

Public Property Get Amount() As Double
    Amount = m.dAmount
End Property
Public Property Let Amount(ByVal dAmount As Double)
    m.dAmount = dAmount
End Property

Public Property Get Notes() As String
    Notes = m.strNotes
End Property
Public Property Let Notes(ByVal strNotes As String)
    m.strNotes = strNotes
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Load
'' Description: Load the adjustment from the database
'' Inputs:      Adjustment ID, Recordset
'' Returns:     True if Loaded, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Load(ByVal lAdjustmentID As Long, Optional ByVal rs As Recordset = Nothing) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    bReturn = False
    If rs Is Nothing Then
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccountAdjustments] " & _
                    "WHERE [AdjustmentID]=" & Str(lAdjustmentID) & ";", dbOpenDynaset)
    End If
    If Not (rs.BOF And rs.EOF) Then
        m.lAdjumtmentID = rs!AdjustmentID
        m.lAccountID = rs!AccountID
        m.strDescription = rs!Description
        m.dTime = rs!AdjustmentTime
        m.dAmount = rs!AdjustmentAmount
        m.strNotes = rs!Notes
        
        bReturn = True
    End If
    
    Load = bReturn

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cAccountAdjustment.Load"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Save the adjustment to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Save()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccountAdjustments] " & _
                "WHERE [AdjustmentID]=" & Str(m.lAdjumtmentID) & ";", dbOpenDynaset)
    If rs.BOF And rs.EOF Then
        rs.AddNew
    Else
        rs.Edit
    End If
    
    rs!AccountID = m.lAccountID
    rs!Description = m.strDescription
    rs!AdjustmentTime = m.dTime
    rs!AdjustmentAmount = m.dAmount
    rs!Notes = m.strNotes
    m.lAdjumtmentID = rs!AdjustmentID
    
    rs.Update
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccountAdjustment.Save"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Delete
'' Description: Delete the adjustment from the database
'' Inputs:      None
'' Returns:     True if deleted, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Delete() As Boolean
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblAccountAdjustments] " & _
                "WHERE [AdjustmentID]=" & Str(m.lAdjumtmentID) & ";", dbOpenDynaset)
    If Not (rs.BOF And rs.EOF) Then
        rs.Delete
        bReturn = True
    End If
    
    Delete = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAccountAdjustment.Delete"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToString
'' Description: Create an english text string based on the adjustment info
'' Inputs:      None
'' Returns:     Text String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToString() As String
On Error GoTo ErrSection:

    ToString = Str(m.dAmount) & " at " & DateFormat(m.dTime, MM_DD_YYYY, HH_MM_SS, AMPM_UPPER) & " (" & m.strDescription & ")"

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cAccountAdjustment.ToString"
    
End Function
