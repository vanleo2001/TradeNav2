VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPtActivity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cPtActivity.cls
'' Description: Properties and Methods for a Trade Tracker Activity
''
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    lActivityID As Long
    lAccountID As Long
    strName As String
    ActivityType As eTT_ActivityType
    StartDate As Date
    EndDate As Date
    strJournal As String
    dProfitTarget As Double
    dStopLoss As Double
    DropDeadDate As Date
    dMarginDeposit As Double
    dAdjustAmount As Double
    dClosedProfit As Double
    dClosedProfitDB As Double
    bIsOpen As Boolean
    
    Positions As cGdTree
End Type
Private m As mPrivate

Public Property Get ActivityID() As Long
    ActivityID = m.lActivityID
End Property
Public Property Get AccountID() As Long
    AccountID = m.lAccountID
End Property
Public Property Get Name() As String
    Name = m.strName
End Property
Public Property Get ActivityType() As eTT_ActivityType
    ActivityType = m.ActivityType
End Property
Public Property Get StartDate() As Date
    StartDate = m.StartDate
End Property
Public Property Get EndDate() As Date
    EndDate = m.EndDate
End Property
Public Property Get Journal() As String
    Journal = m.strJournal
End Property
Public Property Get ProfitTarget() As Double
    ProfitTarget = m.dProfitTarget
End Property
Public Property Get StopLoss() As Double
    StopLoss = m.dStopLoss
End Property
Public Property Get DropDeadDate() As Date
    DropDeadDate = m.DropDeadDate
End Property
Public Property Get MarginDeposit() As Double
    MarginDeposit = m.dMarginDeposit
End Property
Public Property Get AdjustAmount() As Double
    AdjustAmount = m.dAdjustAmount
End Property
Public Property Get ClosedProfit() As Double
    ClosedProfit = m.dClosedProfit
End Property
Public Property Get ClosedProfitDB() As Double
    ClosedProfitDB = m.dClosedProfitDB
End Property
Public Property Get IsOpen() As Boolean
    IsOpen = m.bIsOpen
End Property
Public Property Get Positions() As cGdTree
    Set Positions = m.Positions
End Property

Public Property Let AccountID(ByVal pData As Long)
    m.lAccountID = pData
End Property
Public Property Let Name(ByVal pData As String)
    m.strName = pData
End Property
Public Property Let ActivityType(ByVal pData As eTT_ActivityType)
    m.ActivityType = pData
End Property
Public Property Let StartDate(ByVal pData As Date)
    m.StartDate = pData
End Property
Public Property Let EndDate(ByVal pData As Date)
    m.EndDate = pData
End Property
Public Property Let Journal(ByVal pData As String)
    m.strJournal = pData
End Property
Public Property Let ProfitTarget(ByVal pData As Double)
    m.dProfitTarget = pData
End Property
Public Property Let StopLoss(ByVal pData As Double)
    m.dStopLoss = pData
End Property
Public Property Let DropDeadDate(ByVal pData As Date)
    m.DropDeadDate = pData
End Property
Public Property Let MarginDeposit(ByVal pData As Double)
    m.dMarginDeposit = pData
End Property
Public Property Let AdjustAmount(ByVal pData As Double)
    m.dAdjustAmount = pData
End Property
Public Property Let ClosedProfit(ByVal dClosedProfit As Double)
    m.dClosedProfit = dClosedProfit
End Property
Public Property Let IsOpen(ByVal bIsOpen As Boolean)
    m.bIsOpen = bIsOpen
End Property
Public Property Let Positions(ByVal pData As cGdTree)
    Set m.Positions = pData
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Load
'' Description: Load the activity
'' Inputs:      Activity ID, Recordset
'' Returns:     True if loaded, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Load(ByVal lActivityID As Long, Optional ByVal rs As Recordset = Nothing) As Boolean
On Error GoTo ErrSection

    Dim Position As New cPtPosition     ' Position object
    Dim dClosedProfit As Double         ' Closed Profit
    
    ' Load the Activity information...
    If rs Is Nothing Then
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblActivities] " & _
                    "WHERE [ActivityID]=" & Str(lActivityID) & ";", dbOpenDynaset)
    End If
    
    If Not rs.EOF Then
        m.lActivityID = rs!ActivityID
        m.lAccountID = rs!AccountID
        m.strName = NullChk(rs!ActivityName)
        m.ActivityType = rs!ActivityType
        m.StartDate = NullChk(rs!StartingDate, 0)
        m.EndDate = NullChk(rs!EndingDate, 0)
        m.strJournal = NullChk(rs!Journal)
        m.dProfitTarget = rs!ProfitTarget
        m.dStopLoss = rs!StopLoss
        m.DropDeadDate = NullChk(rs!GetOutDate, 0)
        m.dMarginDeposit = rs!MarginDeposit
        m.dAdjustAmount = NullChk(rs!AdjustAmount, 0)
        m.dClosedProfitDB = NullChk(rs!ClosedProfit, 0#)
        m.bIsOpen = NullChk(rs!IsOpen, False)
    End If
    
    dClosedProfit = 0#
    
    If m.ActivityType <> eTT_ActivityType_Adjustment Then
        ' Load Positions for this Activity...
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblPositions] " & _
                    "WHERE [ActivityID]=" & Str(lActivityID) & ";", dbOpenDynaset)
        Do While Not rs.EOF
            Set Position = New cPtPosition
            If Position.Load(rs!PositionID, rs) Then
                m.Positions.Add Position, Str(Position.PositionID)
                dClosedProfit = dClosedProfit + Position.ClosedProfit
            End If
            
            rs.MoveNext
        Loop
        
        m.dClosedProfit = dClosedProfit
    End If
    
    Load = True
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtActivity.Load"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Save
'' Description: Save the activity
'' Inputs:      Save Positions also?
'' Returns:     True if saved, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Save(Optional ByVal bSavePositions As Boolean = True) As Boolean
On Error GoTo ErrSection

    Dim rs As Recordset                 ' Recordset into the database
    Dim lIndex As Long                  ' Index into a for loop
    
    ' Determine if the activity is still open...
    m.bIsOpen = False
    For lIndex = 1 To m.Positions.Count
        If m.Positions(lIndex).IsOpen = True Then
            m.bIsOpen = True
            Exit For
        End If
    Next lIndex
    
    ' Save the Activity information...
    Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblActivities] " & _
                "WHERE [ActivityID]=" & Str(m.lActivityID) & ";", dbOpenDynaset)
    If Not rs.EOF Then
        rs.Edit
    Else
        rs.AddNew
        rs!AccountID = m.lAccountID
    End If
    rs!ActivityName = m.strName
    rs!ActivityType = m.ActivityType
    rs!StartingDate = m.StartDate
    rs!EndingDate = m.EndDate
    rs!Journal = m.strJournal
    rs!ProfitTarget = m.dProfitTarget
    rs!StopLoss = m.dStopLoss
    rs!GetOutDate = m.DropDeadDate
    rs!MarginDeposit = m.dMarginDeposit
    rs!AdjustAmount = m.dAdjustAmount
    rs!ClosedProfit = m.dClosedProfit
    rs!IsOpen = m.bIsOpen
    m.lActivityID = rs!ActivityID
    rs.Update
    
    m.dClosedProfitDB = m.dClosedProfit
    
    If bSavePositions = True Then
        ' Save the Positions for this Activity...
        For lIndex = 1 To m.Positions.Count
            m.Positions(lIndex).ActivityID = m.lActivityID
            m.Positions(lIndex).Save
            m.Positions.Key(lIndex) = m.Positions(lIndex).PositionID
        Next lIndex
        
        ' Delete Positions that no longer exist in this Activity...
        Set rs = g.dbPaper.OpenRecordset("SELECT * FROM [tblPositions] " & _
                    "WHERE [ActivityID]=" & Str(m.lActivityID) & ";", dbOpenDynaset)
        Do While Not rs.EOF
            If Not m.Positions.Exists(Str(rs!PositionID)) Then rs.Delete
            rs.MoveNext
        Loop
    End If
    
    Save = True
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cPtActivity.Save"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize members when the class is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.Positions = New cGdTree
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtActivity.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Terminate members when the class is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.Positions = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cPtActivity.Class_Terminate"
    
End Sub
