VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cInfoFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cInfoFile.cls
'' Description: Information about the backup stored in a file
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author  Description
'' 02/24/2009   DAJ     Created
'' 09/09/2009   DAJ     Different routines for customer number/data service ID
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    lArchiveVersion As Long             ' Structural version of the archive
    strCustomerID As String             ' Customer ID
    strMachineID As String              ' Machine ID
    dDateTime As Double                 ' Date/Time of the archive
    lTradeNavBuild As Long              ' Trade Navigator build number
End Type
Private m As mPrivate

Public Property Get ArchiveVersion() As Long
    ArchiveVersion = m.lArchiveVersion
End Property
Public Property Let ArchiveVersion(ByVal lArchiveVersion As Long)
    m.lArchiveVersion = lArchiveVersion
End Property

Public Property Get CustomerID() As String
    CustomerID = m.strCustomerID
End Property
Public Property Let CustomerID(ByVal strCustomerID As String)
    m.strCustomerID = strCustomerID
End Property

Public Property Get MachineID() As String
    MachineID = m.strMachineID
End Property
Public Property Let MachineID(ByVal strMachineID As String)
    m.strMachineID = strMachineID
End Property

Public Property Get DateTime() As Double
    DateTime = m.dDateTime
End Property
Public Property Let DateTime(ByVal dDateTime As Double)
    m.dDateTime = dDateTime
End Property

Public Property Get TradeNavBuild() As Long
    TradeNavBuild = m.lTradeNavBuild
End Property
Public Property Let TradeNavBuild(ByVal lTradeNavBuild As Long)
    m.lTradeNavBuild = lTradeNavBuild
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetInformation
'' Description: Fill the class members with the appropriate information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetInformation()
On Error GoTo ErrSection:

    Dim strTnVersion As String          ' Version of Trade Navigator

    m.lArchiveVersion = kArchiveVersion
    m.strCustomerID = RI_GetDataServiceID
    m.strMachineID = RI_GetMachineID
    m.dDateTime = Now
    
    strTnVersion = FileVersion(AddSlash(App.Path) & "NavSuite.EXE")
    m.lTradeNavBuild = CLng(Val(Parse(strTnVersion, ".", 4)))

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cInfoFile.GetInformation"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToFile
'' Description: Dump the information to a file
'' Inputs:      Filename
'' Returns:     True on Success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToFile(ByVal strFilename As String) As Boolean
On Error GoTo ErrSection:

    Dim mb As New cMemBuffer            ' Buffer to hold info to send to file
    
    mb.PutLong m.lArchiveVersion
    mb.PutLong Len(m.strCustomerID)
    mb.PutStr m.strCustomerID
    mb.PutStr m.strMachineID
    mb.PutDouble m.dDateTime
    mb.PutLong m.lTradeNavBuild
    
    ToFile = mb.ToFile(strFilename)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cInfoFile.ToFile"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromFile
'' Description: Get the information from a file
'' Inputs:      Filename
'' Returns:     True on Success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FromFile(ByVal strFilename As String) As Boolean
On Error GoTo ErrSection:

    Dim mb As New cMemBuffer            ' Buffer to hold info retreived from a file
    Dim bReturn As Boolean              ' Return value for the function
    Dim lStrLength As Long              ' Length of customer ID string
    
    bReturn = mb.FromFile(strFilename)
    If bReturn = True Then
        m.lArchiveVersion = mb.GetLong
        lStrLength = mb.GetLong
        m.strCustomerID = mb.GetStr(lStrLength)
        m.strMachineID = mb.GetStr(11)
        m.dDateTime = mb.GetDouble
        m.lTradeNavBuild = mb.GetLong
    End If
    
    FromFile = bReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cInfoFile.FromFile"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CustomerNumber
'' Description: Get the customer number from the customer ID
'' Inputs:      None
'' Returns:     Customer Number
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CustomerNumber() As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    If Len(m.strCustomerID) > 3 Then
        strReturn = Left(m.strCustomerID, Len(m.strCustomerID) - 3)
    End If
    
    CustomerNumber = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cInfoFile.CustomerNumber"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DataService
'' Description: Get the data service from the customer ID
'' Inputs:      None
'' Returns:     Data Service
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DataService() As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    If Len(m.strCustomerID) > 3 Then
        strReturn = Right(m.strCustomerID, 3)
    End If
    
    DataService = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cInfoFile.DataService"
    
End Function

