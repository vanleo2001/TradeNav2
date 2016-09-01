VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cAccount.cls
'' Description: Class that holds information about an account
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    strID As String                     ' Account ID from the database on the server
    strAccountNumber As String          ' Account number
    strName As String                   ' Display name for the account
    strFeedyardID As String             ' ID for the feedyard that the account is associated with
    nBroker As eTT_AccountType          ' Broker type for this account
    strFcmNumber As String              ' FCM account number
    bHasFills As Boolean                ' Do we have fills for this account?
    bDeleted As Boolean                 ' Is the account marked as deleted?
    bInRefresh As Boolean               ' Did this account come in a refresh?
End Type
Private m As mPrivate

Public Property Get ID() As String
    ID = m.strID
End Property
Public Property Let ID(ByVal strID As String)
    m.strID = strID
End Property

Public Property Get AccountNumber() As String
    AccountNumber = m.strAccountNumber
End Property
Public Property Let AccountNumber(ByVal strAccountNumber As String)
    m.strAccountNumber = strAccountNumber
End Property

Public Property Get Name() As String
    Name = m.strName
End Property
Public Property Let Name(ByVal strName As String)
    m.strName = strName
End Property

Public Property Get FeedYardID() As String
    FeedYardID = m.strFeedyardID
End Property
Public Property Let FeedYardID(ByVal strFeedyardID As String)
    m.strFeedyardID = strFeedyardID
End Property

Public Property Get Broker() As eTT_AccountType
    Broker = m.nBroker
End Property
Public Property Let Broker(ByVal nBroker As eTT_AccountType)
    m.nBroker = nBroker
End Property

Public Property Get FcmNumber() As String
    FcmNumber = m.strFcmNumber
End Property
Public Property Let FcmNumber(ByVal strFcmNumber As String)
    m.strFcmNumber = strFcmNumber
End Property

Public Property Get HasFills() As Boolean
    HasFills = m.bHasFills
End Property
Public Property Let HasFills(ByVal bHasFills As Boolean)
    m.bHasFills = bHasFills
End Property

Public Property Get Deleted() As Boolean
    Deleted = m.bDeleted
End Property
Public Property Let Deleted(ByVal bDeleted As Boolean)
    m.bDeleted = bDeleted
End Property

Public Property Get InRefresh() As Boolean
    InRefresh = m.bInRefresh
End Property
Public Property Let InRefresh(ByVal bInRefresh As Boolean)
    m.bInRefresh = bInRefresh
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize the class members when the class is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    m.strID = ""
    m.strAccountNumber = ""
    m.strName = ""
    m.strFeedyardID = ""
    m.nBroker = -1&
    m.strFcmNumber = ""
    m.bHasFills = False
    m.bDeleted = False
    m.bInRefresh = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccount.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up the class members when the class is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cAccount.Class_Terminate"
    
End Sub
