VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLotColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cLotColumn.cls
'' Description: Information about a lot column
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 11/15/2013   DAJ         Added Category ID, Subcategory ID, IsAggregate
'' 11/18/2013   DAJ         Added Ordinal
'' 11/26/2013   DAJ         Changed IsAggregate from Boolean to Long; Added DisplayFormat
'' 12/05/2013   DAJ         Aggregate Column Modes
'' 01/23/2014   DAJ         Multiple owners per lot
'' 03/07/2014   DAJ         Moved into NavCattle.DLL
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    lID As Long                         ' Primary Key
    strColumnHeader As String           ' Column header
    strTooltipText As String            ' Tooltip text
    strFormat As String                 ' Format
    strDisplayFormat As String          ' Display format
    strKeyValueField As String          ' Name of the key-value key
    bTotal As Boolean                   ' Can we total this column?
    lCategoryID As Long                 ' Category ID
    lSubCategoryID As Long              ' Sub Category ID
    lIsAggregate As Long                ' Can this column be aggregated
    lOrdinal As Long                    ' Ordinal value
    
    bAlwaysHidden As Boolean            ' Column is always hidden?
    bFeedyardHidden As Boolean          ' Column is hidden for feedyard customer?
    bUserHidden As Boolean              ' Column is hidden by user?
    bAllowUserMove As Boolean           ' Allow the user to move the column?
    
    lPosition As Long                   ' Position in the grid
    lWidth As Long                      ' Column width
End Type
Private m As mPrivate

Public Property Get ID() As Long
    ID = m.lID
End Property
Public Property Let ID(ByVal lID As Long)
    m.lID = lID
End Property

Public Property Get ColumnHeader() As String
    ColumnHeader = m.strColumnHeader
End Property
Public Property Let ColumnHeader(ByVal strColumnHeader As String)
    m.strColumnHeader = strColumnHeader
End Property

Public Property Get TooltipText() As String
    TooltipText = m.strTooltipText
End Property
Public Property Let TooltipText(ByVal strTooltipText As String)
    m.strTooltipText = strTooltipText
End Property

Public Property Get Format() As String
    Format = m.strFormat
End Property
Public Property Let Format(ByVal strFormat As String)
    m.strFormat = strFormat
End Property

Public Property Get DisplayFormat() As String
    DisplayFormat = m.strDisplayFormat
End Property
Public Property Let DisplayFormat(ByVal strDisplayFormat As String)
    m.strDisplayFormat = strDisplayFormat
End Property

Public Property Get KeyValueField() As String
    KeyValueField = m.strKeyValueField
End Property
Public Property Let KeyValueField(ByVal strKeyValueField As String)
    m.strKeyValueField = strKeyValueField
End Property

Public Property Get Total() As Boolean
    Total = m.bTotal
End Property
Public Property Let Total(ByVal bTotal As Boolean)
    m.bTotal = bTotal
End Property

Public Property Get CategoryID() As Long
    CategoryID = m.lCategoryID
End Property
Public Property Let CategoryID(ByVal lCategoryID As Long)
    m.lCategoryID = lCategoryID
End Property

Public Property Get SubCategoryID() As Long
    SubCategoryID = m.lSubCategoryID
End Property
Public Property Let SubCategoryID(ByVal lSubCategoryID As Long)
    m.lSubCategoryID = lSubCategoryID
End Property

Public Property Get IsAggregate() As Long
    IsAggregate = m.lIsAggregate
End Property
Public Property Let IsAggregate(ByVal lIsAggregate As Long)
    m.lIsAggregate = lIsAggregate
End Property

Public Property Get AggregateIsSingle() As Boolean
    AggregateIsSingle = GetBit(m.lIsAggregate, 1)
End Property

Public Property Get AggregateIsGroup() As Boolean
    AggregateIsGroup = GetBit(m.lIsAggregate, 2)
End Property

Public Property Get AggregateIsOwner() As Boolean
    AggregateIsOwner = GetBit(m.lIsAggregate, 5)
End Property

Public Property Get AggregateMode() As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the property
    
    lReturn = 0&
    SetBit lReturn, 1, GetBit(m.lIsAggregate, 3)
    SetBit lReturn, 2, GetBit(m.lIsAggregate, 4)
    
    AggregateMode = lReturn

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cLotColumn.AggregateMode.Get"
    
End Property

Public Property Get Ordinal() As Long
    Ordinal = m.lOrdinal
End Property
Public Property Let Ordinal(ByVal lOrdinal As Long)
    m.lOrdinal = lOrdinal
End Property

Public Property Get AlwaysHidden() As Boolean
    AlwaysHidden = m.bAlwaysHidden
End Property
Public Property Let AlwaysHidden(ByVal bAlwaysHidden As Boolean)
    m.bAlwaysHidden = bAlwaysHidden
End Property

Public Property Get FeedyardHidden() As Boolean
    FeedyardHidden = m.bFeedyardHidden
End Property
Public Property Let FeedyardHidden(ByVal bFeedyardHidden As Boolean)
    m.bFeedyardHidden = bFeedyardHidden
End Property

Public Property Get UserHidden() As Boolean
    UserHidden = m.bUserHidden
End Property
Public Property Let UserHidden(ByVal bUserHidden As Boolean)
    m.bUserHidden = bUserHidden
End Property

Public Property Get AllowUserMove() As Boolean
    AllowUserMove = m.bAllowUserMove
End Property
Public Property Let AllowUserMove(ByVal bAllowUserMove As Boolean)
    m.bAllowUserMove = bAllowUserMove
End Property

Public Property Get Position() As Long
    Position = m.lPosition
End Property
Public Property Let Position(ByVal lPosition As Long)
    m.lPosition = lPosition
End Property

Public Property Get Width() As Long
    Width = m.lWidth
End Property
Public Property Let Width(ByVal lWidth As Long)
    m.lWidth = lWidth
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromString
'' Description: Extract the values from the given string
'' Inputs:      Lot Column Info
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromString(ByVal strLotColumnInfo As String)
On Error GoTo ErrSection:

    Dim turnkeyMessage As cBrokerMessage ' Turnkey message
    
    Set turnkeyMessage = New cBrokerMessage
    turnkeyMessage.FromString strLotColumnInfo
    
    m.lID = CLng(Val(turnkeyMessage("ID")))
    m.strColumnHeader = turnkeyMessage("ColumnHeader")
    m.strTooltipText = turnkeyMessage("TooltipText")
    m.strFormat = turnkeyMessage("Format")
    m.strKeyValueField = turnkeyMessage("KeyValueField")
    m.bTotal = (turnkeyMessage("Totalable") = "1")
    m.lCategoryID = CLng(Val(turnkeyMessage("CategoryID")))
    m.lSubCategoryID = CLng(Val(turnkeyMessage("SubCategoryID")))
    m.lIsAggregate = CLng(Val(turnkeyMessage("IsAggregate")))
    m.lOrdinal = CLng(Val(turnkeyMessage("Ordinal")))
    
    m.bAlwaysHidden = (UCase(m.strFormat) = "HIDDEN")
    m.bFeedyardHidden = False
    m.bUserHidden = False
    m.bAllowUserMove = Not m.bAlwaysHidden

    Select Case UCase(m.strFormat)
        Case "NUMBER"
            m.strDisplayFormat = turnkeyMessage("DisplayFormat")
            If Len(m.strDisplayFormat) = 0 Then
                m.strDisplayFormat = "#,##0"
            End If
        
        Case "DATE"
            m.strDisplayFormat = DateFormat("Format", MM_DD_YYYY)
        
        Case "CURRENCY"
            m.strDisplayFormat = turnkeyMessage("DisplayFormat")
            If Len(m.strDisplayFormat) = 0 Then
                m.strDisplayFormat = "$#,##0.00"
            End If
    
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cLotColumn.FromString"
    
End Sub

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

    m.lID = 0
    m.strColumnHeader = ""
    m.strTooltipText = ""
    m.strFormat = ""
    m.strDisplayFormat = ""
    m.strKeyValueField = ""
    m.bTotal = False
    m.lCategoryID = 0&
    m.lSubCategoryID = 0&
    m.lIsAggregate = 0&
    m.lOrdinal = 0&
    
    m.bAlwaysHidden = False
    m.bFeedyardHidden = False
    m.bUserHidden = False
    m.bAllowUserMove = True
    
    m.lPosition = 0&
    m.lWidth = 0&

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cLotColumn.Class_Initialize"
    
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
    RaiseError "cLotColumn.Class_Terminate"
    
End Sub
