VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOrderLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cOrderLink.cls
'' Description: Class to handle an order link
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 08/02/2011   DAJ         Added order objects for each order
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    lOrderLinkID As Long                ' Order Link ID from the database
    
    lDatabaseId1 As Long                ' Database ID for Order 1
    strBrokerId1 As String              ' Broker ID for Order 1
    Order1 As cPtOrder                  ' Order 1
    
    lDatabaseId2 As Long                ' Database ID for Order 2
    strBrokerId2 As String              ' Broker ID for Order 2
    Order2 As cPtOrder                  ' Order 2
    
    nStatus As eGDOrderLinkStatus       ' Status of the link
End Type
Private m As mPrivate

Public Property Get OrderLinkID() As Long
    OrderLinkID = m.lOrderLinkID
End Property

Public Property Get DatabaseId1() As Long
    DatabaseId1 = m.lDatabaseId1
End Property
Public Property Let DatabaseId1(ByVal lDatabaseId1 As Long)
    m.lDatabaseId1 = lDatabaseId1
End Property

Public Property Get BrokerId1() As String
    BrokerId1 = m.strBrokerId1
End Property
Public Property Let BrokerId1(ByVal strBrokerId1 As String)
    m.strBrokerId1 = strBrokerId1
End Property

Public Property Get Order1() As cPtOrder
    Set Order1 = m.Order1
End Property
Public Property Let Order1(ByVal Order As cPtOrder)
    Set m.Order1 = Order
End Property

Public Property Get DatabaseId2() As Long
    DatabaseId2 = m.lDatabaseId2
End Property
Public Property Let DatabaseId2(ByVal lDatabaseId2 As Long)
    m.lDatabaseId2 = lDatabaseId2
End Property

Public Property Get BrokerId2() As String
    BrokerId2 = m.strBrokerId2
End Property
Public Property Let BrokerId2(ByVal strBrokerId2 As String)
    m.strBrokerId2 = strBrokerId2
End Property

Public Property Get Order2() As cPtOrder
    Set Order2 = m.Order2
End Property
Public Property Let Order2(ByVal Order As cPtOrder)
    Set m.Order2 = Order
End Property

Public Property Get Status() As eGDOrderLinkStatus
    Status = m.nStatus
End Property
Public Property Let Status(ByVal nStatus As eGDOrderLinkStatus)
    m.nStatus = nStatus
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Load
'' Description: Load up an order link from the database
'' Inputs:      ID, Recordset
'' Returns:     True on success, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Load(ByVal lOrderLinkID As Long, Optional ByVal rs As Recordset = Nothing) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim Order As cPtOrder               ' Order object
    
    If rs Is Nothing Then
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrderLinks] " & _
                "WHERE [OrderLinkID]=" & Str(lOrderLinkID) & ";", dbOpenDynaset)
    End If
    
    bReturn = False
    If Not (rs.BOF And rs.EOF) Then
        m.lOrderLinkID = lOrderLinkID
        m.lDatabaseId1 = rs!OrderID1
        m.strBrokerId1 = rs!BrokerId1
        m.lDatabaseId2 = rs!OrderID2
        m.strBrokerId2 = rs!BrokerId2
        m.nStatus = rs!Status
        
        If m.lDatabaseId1 = 0 Then
            Set m.Order1 = Nothing
        Else
            Set Order = New cPtOrder
            If Order.Load(m.lDatabaseId1) Then
                Set m.Order1 = Order.MakeCopy
            Else
                Set m.Order1 = Nothing
            End If
        End If
        
        If m.lDatabaseId2 = 0 Then
            Set m.Order2 = Nothing
        Else
            Set Order = New cPtOrder
            If Order.Load(m.lDatabaseId2) Then
                Set m.Order2 = Order.MakeCopy
            Else
                Set m.Order2 = Nothing
            End If
        End If
        
        bReturn = True
    End If
    
    Load = bReturn

ErrExit:
    Exit Function

ErrSection:
    RaiseError "cOrderLink.Load"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Save an order link to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Save()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrderLinks] " & _
            "WHERE [OrderLinkID]=" & Str(m.lOrderLinkID) & ";", dbOpenDynaset)
    
    If (rs.BOF And rs.EOF) Then
        rs.AddNew
        m.lOrderLinkID = rs!OrderLinkID
    Else
        rs.Edit
    End If
    
    rs!OrderID1 = m.lDatabaseId1
    rs!BrokerId1 = m.strBrokerId1
    rs!OrderID2 = m.lDatabaseId2
    rs!BrokerId2 = m.strBrokerId2
    rs!Status = m.nStatus
    
    rs.Update
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderLink.Save"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Delete
'' Description: Delete an order link from the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Delete()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrderLinks] " & _
            "WHERE [OrderLinkID]=" & Str(m.lOrderLinkID) & ";", dbOpenDynaset)
            
    If Not (rs.BOF And rs.EOF) Then
        rs.Delete
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderLink.Delete"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OtherDatabaseID
'' Description: Determine the other database ID given the one passed in
'' Inputs:      Database ID
'' Returns:     Other DatabaseID or 0 if passed in one doesn't exist
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OtherDatabaseID(ByVal lDatabaseID As Long) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    
    lReturn = 0&
    If m.lDatabaseId1 = lDatabaseID Then
        lReturn = m.lDatabaseId2
    ElseIf m.lDatabaseId2 = lDatabaseID Then
        lReturn = m.lDatabaseId1
    End If
    
    OtherDatabaseID = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOrderLink.OtherDatabaseID"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OtherBrokerID
'' Description: Determine the other broker ID given the one passed in
'' Inputs:      Broker ID
'' Returns:     Other Broker ID or "" if passed in one doesn't exist
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OtherBrokerID(ByVal strBrokerID As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    If m.strBrokerId1 = strBrokerID Then
        strReturn = m.strBrokerId2
    ElseIf m.strBrokerId2 = strBrokerID Then
        strReturn = m.strBrokerId1
    End If
    
    OtherBrokerID = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOrderLink.OtherBrokerID"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToString
'' Description: Convert the information in the link to a string
'' Inputs:      None
'' Returns:     String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ToString() As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    
    strReturn = Str(m.lDatabaseId1) & " (" & m.strBrokerId1 & ");"
    strReturn = strReturn & Str(m.lDatabaseId2) & " (" & m.strBrokerId2 & ");"
    strReturn = strReturn & OrderLinkStatusString(m.nStatus)
    
    ToString = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOrderLink.ToString"
    
End Function
