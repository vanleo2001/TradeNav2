VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOrderLegs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cOrderLegs.cls
'' Description: Wrapper class for a collection of order legs
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 04/03/2012   DAJ         Clear the collection before loading
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    OrderLegs As cGdTree                ' Collection of order legs
End Type
Private m As mPrivate

Public Property Get Count() As Long
    Count = m.OrderLegs.Count
End Property

Public Property Get Item(ByVal lLegNumber As Long) As cOrderLeg
Attribute Item.VB_UserMemId = 0
    
    Dim OrderLeg As cOrderLeg           ' Order leg object in case we need to create one

    If m.OrderLegs.Exists(lLegNumber) = False Then
        Set OrderLeg = New cOrderLeg
        OrderLeg.LegNumber = lLegNumber
        Add OrderLeg
    End If
        
    Set Item = m.OrderLegs(lLegNumber)

End Property
Public Property Let Item(ByVal lLegNumber As Long, OrderLeg As cOrderLeg)
    Set m.OrderLegs(lLegNumber) = OrderLeg
End Property

Public Function Exists(ByVal lLegNumber As Long) As Boolean
    Exists = m.OrderLegs.Exists(lLegNumber)
End Function

Public Function Add(OrderLeg As cOrderLeg) As Long
    Add = m.OrderLegs.Add(OrderLeg, , OrderLeg.LegNumber, eTREE_Myself)
End Function

Public Function Remove(ByVal lLegNumber As Long) As Boolean
    Remove = m.OrderLegs.Remove(lLegNumber)
End Function

Public Sub Clear()
    m.OrderLegs.Clear
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Load
'' Description: Load the order legs from the database for the given order
'' Inputs:      Order ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Load(ByVal lOrderID As Long, Optional ByVal rs As Recordset = Nothing)
On Error GoTo ErrSection:

    Dim OrderLeg As cOrderLeg           ' Order leg object for adding to collection

    If rs Is Nothing Then
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblOrderLegs] " & _
                    "WHERE [OrderID]=" & Str(lOrderID) & " " & _
                    "ORDER BY [LegNumber];", dbOpenDynaset)
    End If
    
    m.OrderLegs.Clear
    Do While Not rs.EOF
        Set OrderLeg = New cOrderLeg
        If OrderLeg.Load(lOrderID, rs!LegNumber, rs) Then
            Add OrderLeg
        End If
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderLegs.Load"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Save the order legs to the database
'' Inputs:      Order ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Save(ByVal lOrderID As Long)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To m.OrderLegs.Count
        Item(lIndex).OrderID = lOrderID
        Item(lIndex).Save
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderLegs.Save"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MakeCopy
'' Description: Make a copy of the order
'' Inputs:      None
'' Returns:     Copy of the order
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MakeCopy() As cOrderLegs
On Error GoTo ErrSection:

    Dim aCopy As New cOrderLegs         ' Copy to return
    
    aCopy.SetPrivateData m
    Set MakeCopy = aCopy

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cOrderLegs.MakeCopy"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetPrivateData
'' Description: Set the private data for copying purposes
'' Inputs:      Private members to copy from
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub SetPrivateData(mCopyFrom As mPrivate)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    m = mCopyFrom
    
    Set m.OrderLegs = New cGdTree
    For lIndex = 1 To mCopyFrom.OrderLegs.Count
        m.OrderLegs.Add mCopyFrom.OrderLegs(lIndex).MakeCopy
    Next lIndex
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderLegs.SetPrivateData"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize class members when the object is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.OrderLegs = New cGdTree

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderLegs.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Terminate class members when the object is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.OrderLegs = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cOrderLegs.Class_Terminate"
    
End Sub
