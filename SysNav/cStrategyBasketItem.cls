VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStrategyBasketItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cStrategyBasketItem.cls
'' Description: Wrapper class for a strategy basket item
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 02/21/2012   DAJ         Added contract multiplier
'' 04/03/2013   DAJ         Move Strategy Baskets into the database
'' 08/09/2013   DAJ         Build sort key when components change, not each time
'' 10/04/2013   DAJ         Optionally validate checksums on strategy basket items
'' 08/19/2014   DAJ         Expose Strategy Basket Item Inputs
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    lID As Long                         ' Strategy basket item ID
    lStrategyBasketID As Long           ' Strategy basket ID
    lStrategyID As Long                 ' Unique ID for the strategy
    strStrategyName As String           ' Name of the strategy to run
    strSymbolGroupID As String          ' Symbol group to run the strategy on
    strSymbolGroupName As String        ' Symbol group to run the strategy on
    lSymbolID As Long                   ' Symbol ID to run the strategy on
    strSymbol As String                 ' Symbol to run the strategy on
    strPeriod As String                 ' Bar Period for the data
    lFromDate As Long                   ' Date to run the strategy from
    lToDate As Long                     ' Date to run the strategy to
    bToEndOfData As Boolean             ' Run the strategy through the end of the data?
    strToDateDisplay As String          ' Display string for the "To Date"
    bSplitAdjust As Boolean             ' Split adjust the data?
    dContractMultiplier As Double       ' Contract multiplier
    
    strKey As String                    ' Key for the strategy basket item
    
    Parms As cGdTree                    ' Parm overrides
End Type
Private m As mPrivate

Public Property Get ID() As Long
    ID = m.lID
End Property
Public Property Let ID(ByVal lID As Long)
    m.lID = lID
End Property

Public Property Get StrategyBasketID() As Long
    StrategyBasketID = m.lStrategyBasketID
End Property
Public Property Let StrategyBasketID(ByVal lStrategyBasketID As Long)
    m.lStrategyBasketID = lStrategyBasketID
End Property

Public Property Get StrategyID() As Long
    StrategyID = m.lStrategyID
End Property
Public Property Let StrategyID(ByVal lStrategyID As Long)
    m.lStrategyID = lStrategyID
    BuildKey
End Property

Public Property Get StrategyName() As String
    StrategyName = m.strStrategyName
End Property
Public Property Let StrategyName(ByVal strStrategyName As String)
    m.strStrategyName = strStrategyName
End Property

Public Property Get SymbolGroupID() As String
    SymbolGroupID = m.strSymbolGroupID
End Property
Public Property Let SymbolGroupID(ByVal strSymbolGroupID As String)
    m.strSymbolGroupID = strSymbolGroupID
    BuildKey
End Property

Public Property Get SymbolGroupName() As String
    SymbolGroupName = m.strSymbolGroupName
End Property
Public Property Let SymbolGroupName(ByVal strSymbolGroupName As String)
    m.strSymbolGroupName = strSymbolGroupName
End Property

Public Property Get SymbolID() As Long
    SymbolID = m.lSymbolID
End Property
Public Property Let SymbolID(ByVal lSymbolID As Long)
    m.lSymbolID = lSymbolID
    BuildKey
End Property

Public Property Get Symbol() As String
    Symbol = m.strSymbol
End Property
Public Property Let Symbol(ByVal strSymbol As String)
    m.strSymbol = strSymbol
    BuildKey
End Property

Public Property Get SymbolOrSymbolID() As Variant
    If m.lSymbolID = 0& Then
        SymbolOrSymbolID = m.strSymbol
    Else
        SymbolOrSymbolID = m.lSymbolID
    End If
End Property
Public Property Let SymbolOrSymbolID(ByVal vSymbolOrSymbolID As Variant)
    m.lSymbolID = GetSymbolID(vSymbolOrSymbolID)
    m.strSymbol = GetSymbol(vSymbolOrSymbolID)
    BuildKey
End Property

Public Property Get Period() As String
    Period = m.strPeriod
End Property
Public Property Let Period(ByVal strPeriod As String)
    m.strPeriod = strPeriod
    BuildKey
End Property

Public Property Get FromDate() As Long
    FromDate = m.lFromDate
End Property
Public Property Let FromDate(ByVal lFromDate As Long)
    m.lFromDate = lFromDate
End Property

Public Property Get ToDate() As Long
    ToDate = m.lToDate
End Property
Public Property Let ToDate(ByVal lToDate As Long)
    m.lToDate = lToDate
End Property

Public Property Get ToEndOfData() As Boolean
    ToEndOfData = m.bToEndOfData
End Property
Public Property Let ToEndOfData(ByVal bToEndOfData As Boolean)
    m.bToEndOfData = bToEndOfData
End Property

Public Property Get ToDateDisplay() As String
    If m.bToEndOfData Then
        ToDateDisplay = "End of Data"
    Else
        ToDateDisplay = DateFormat(m.lToDate, MM_DD_YYYY)
    End If
End Property
Public Property Let ToDateDisplay(ByVal strToDateDisplay As String)
    m.strToDateDisplay = strToDateDisplay
End Property

Public Property Get Split() As Boolean
    Split = m.bSplitAdjust
End Property
Public Property Let Split(ByVal bSplit As Boolean)
    m.bSplitAdjust = bSplit
End Property

Public Property Get SplitDisplay() As String
    If m.bSplitAdjust Then
        SplitDisplay = "Split"
    Else
        SplitDisplay = "Unsplit"
    End If
End Property
Private Property Let SplitDisplay(ByVal strSplitDisplay As String)
    m.bSplitAdjust = (UCase(strSplitDisplay) = "SPLIT")
End Property

Public Property Get ContractMultiplier() As Double
    ContractMultiplier = m.dContractMultiplier
End Property
Public Property Let ContractMultiplier(ByVal dContractMultiplier As Double)
    m.dContractMultiplier = dContractMultiplier
End Property

Public Property Get Parms() As cGdTree
    Set Parms = m.Parms
End Property

Public Property Get Key() As String
    Key = m.strKey
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromString
'' Description: Split out the object properties from the given string
'' Inputs:      Strategy Basket Item information
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromString(ByVal strStrategyBasketItem As String)
On Error GoTo ErrSection:

    Dim astrFields As cGdArray          ' Array of fields about the strategy basket
    Dim astrOverrides As cGdArray       ' Overrides broken out into an array
    Dim lIndex As Long                  ' Index into a for loop
    Dim Parm As cStrategyBasketItemParm ' Strategy basket item paramater object
    
    Set astrFields = New cGdArray
    astrFields.SplitFields strStrategyBasketItem, vbTab
    
    m.strStrategyName = astrFields(0)
    m.lStrategyID = CLng(Val(astrFields(1)))
    m.strSymbolGroupID = astrFields(3)
    If Len(m.strSymbolGroupID) = 0 Then
        m.strSymbol = astrFields(2)
    Else
        m.strSymbolGroupName = astrFields(2)
    End If
    m.strPeriod = astrFields(4)
    m.lFromDate = CLng(DateOf(astrFields(5)))
    m.lToDate = CLng(DateOf(astrFields(6)))
    m.bToEndOfData = CBool(Val(astrFields(7)))
    m.strToDateDisplay = astrFields(8)
    SplitDisplay = astrFields(9)
    m.lSymbolID = CLng(Val(astrFields(11)))
    
    If astrFields.Size > 12 Then
        m.dContractMultiplier = Val(astrFields(12))
    Else
        m.dContractMultiplier = 1#
    End If

    If Len(astrFields(10)) > 0 Then
        Set astrOverrides = New cGdArray
        astrOverrides.SplitFields astrFields(10), ","
        
        For lIndex = 0 To astrOverrides.Size - 1
            Set Parm = New cStrategyBasketItemParm
            Parm.FromString astrOverrides(lIndex)
            
            If m.Parms.Exists(Parm.Key) = False Then
                m.Parms.Add Parm, Parm.Key
            End If
        Next lIndex
    End If
    
    BuildKey

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStrategyBasketItem.FromString"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadDb
'' Description: Load the strategy basket item from the database
'' Inputs:      Strategy Basket Item ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LoadDb(ByVal lStrategyBasketItemID As Long, Optional ByVal bValidateCheckSums As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database

    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblStrategyBasketItems] WHERE [StrategyBasketItemID]=" & Str(lStrategyBasketItemID) & ";", dbOpenDynaset)
    LoadDb = LoadDbFromRs(rs, bValidateCheckSums)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cStrategyBasketItem.LoadDb"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadDbFromRs
'' Description: Load the strategy basket item from the database
'' Inputs:      Strategy Basket Item ID, Recordset
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LoadDbFromRs(ByVal rs As Recordset, Optional ByVal bValidateCheckSums As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function

    bReturn = False
    
    If Not (rs.EOF And rs.BOF) Then
        If bValidateCheckSums = True Then
            ValidateCheckSums rs, "tblStrategyBasketItems"
        End If
        
        If rs!CheckSum <> 0.5 Then
            m.lID = rs!StrategyBasketItemID
            m.lStrategyBasketID = rs!StrategyBasketID
            m.lStrategyID = rs!SystemNumber
            m.strStrategyName = SystemNameForID(m.lStrategyID)
            m.strSymbolGroupID = rs!SymbolGroupID
            m.strSymbolGroupName = rs!SymbolGroupName
            m.lSymbolID = rs!SymbolID
            m.strSymbol = rs!Symbol
            m.strPeriod = rs!BarPeriod
            m.lFromDate = rs!FromDate
            m.lToDate = rs!ToDate
            m.bToEndOfData = rs!ToEndOfData
            m.bSplitAdjust = rs!SplitAdjust
            m.dContractMultiplier = rs!ContractMultiplier
            BuildKey
            
            LoadParmsDb
        
            bReturn = True
        End If
    End If
    
    LoadDbFromRs = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cStrategyBasketItem.LoadDbFromRs"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveDb
'' Description: Save the strategy basket item to the database
'' Inputs:      Database
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveDb(Optional DB As Database = Nothing)
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database

    If DB Is Nothing Then
        Set DB = g.dbNav
    End If

    Set rs = DB.OpenRecordset("SELECT * FROM [tblStrategyBasketItems] WHERE [StrategyBasketItemID]=" & Str(m.lID) & ";", dbOpenDynaset)
    If (rs.BOF And rs.EOF) Then
        rs.AddNew
    Else
        rs.Edit
    End If
    
    m.lID = rs!StrategyBasketItemID
    rs!StrategyBasketID = m.lStrategyBasketID
    rs!SystemNumber = m.lStrategyID
    rs!SymbolGroupID = m.strSymbolGroupID
    rs!SymbolGroupName = m.strSymbolGroupName
    rs!SymbolID = m.lSymbolID
    rs!Symbol = m.strSymbol
    rs!BarPeriod = m.strPeriod
    rs!FromDate = m.lFromDate
    rs!ToDate = m.lToDate
    rs!ToEndOfData = m.bToEndOfData
    rs!SplitAdjust = m.bSplitAdjust
    rs!ContractMultiplier = m.dContractMultiplier
    
    rs!CheckSum = BuildCheckSum(rs, "tblStrategyBasketItems")
    rs.Update
    
    SaveParmsDb DB

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStrategyBasketItem.SaveDb"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Validate
'' Description: Make sure the strategy and symbol group still exist for this item
'' Inputs:      None
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Validate(Optional bStrategyNameOrIdChanged As Boolean) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lStrategyID As Long             ' Strategy ID
    Dim strStrategyName As String       ' Strategy name

    bStrategyNameOrIdChanged = False
    
    bReturn = True
    If Len(m.strSymbolGroupID) > 0 Then
        bReturn = (g.SymbolPool.FieldNumForID(m.strSymbolGroupID) <> -1)
    End If
    
    If bReturn = True Then
        lStrategyID = m.lStrategyID
        strStrategyName = m.strStrategyName
        
        If mSysNav.SyncSystemInfo(strStrategyName, lStrategyID) Then
            If strStrategyName <> m.strStrategyName Then
                m.strStrategyName = strStrategyName
                bStrategyNameOrIdChanged = True
            End If
            If lStrategyID <> m.lStrategyID Then
                m.lStrategyID = lStrategyID
                bStrategyNameOrIdChanged = True
            End If
        Else
            bReturn = False
        End If
    End If
    
    Validate = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cStrategyBasketItem.Validate"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ClearID
'' Description: Clear the strategy basket ID
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ClearID()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop

    m.lID = 0
    
    For lIndex = 1 To m.Parms.Count
        m.Parms(lIndex).StrategyBasketItemID = 0
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStrategyBasketItem.ClearID"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasSameOverrides
'' Description: Does this item have the same overrides as the given item?
'' Inputs:      Item to compare
'' Returns:     True if the same overrides, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasSameOverrides(ByVal Item As cStrategyBasketItem) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim Parm As cStrategyBasketItemParm ' Strategy basket item parameter
    
    bReturn = True
    If (ContractMultiplier <> Item.ContractMultiplier) Then
        bReturn = False
    ElseIf Parms.Count <> Item.Parms.Count Then
        bReturn = False
    Else
        For lIndex = 1 To Parms.Count
            Set Parm = Parms(lIndex)
            
            If Item.Parms.Exists(Parm.Key) Then
                If Parm.IsTheSameAs(Item.Parms(Parm.Key)) = False Then
                    bReturn = False
                    Exit For
                End If
            Else
                bReturn = False
                Exit For
            End If
        Next lIndex
    End If
    
    HasSameOverrides = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cStrategyBasketItem.HasSameOverrides"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Overrides
'' Description: Get a delimited list of the overrides
'' Inputs:      None
'' Returns:     Overrides
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Overrides() As String
On Error GoTo ErrSection:

    Dim astrOverrides As cGdArray       ' Array of overrides
    Dim lIndex As Long                  ' Index into a for loop
    Dim Parm As cStrategyBasketItemParm ' Strategy basket item paramater object
    
    Set astrOverrides = New cGdArray
    astrOverrides.Create eGDARRAY_Strings
    
    For lIndex = 1 To m.Parms.Count
        Set Parm = m.Parms(lIndex)
        If Parm.IsOverridden Then
            astrOverrides.Add m.Parms(lIndex).ToString
        End If
    Next lIndex

    Overrides = astrOverrides.JoinFields(",")
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cStrategyBasketItem.Overrides"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasExposedParameters
'' Description: Determine if this item has exposed parameters
'' Inputs:      None
'' Returns:     True if has exposed parameters, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasExposedParameters() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim Parm As cStrategyBasketItemParm ' Strategy basket item paramater object
    
    bReturn = False
    
    For lIndex = 1 To m.Parms.Count
        Set Parm = m.Parms(lIndex)
        If Parm.IsExposed Then
            bReturn = True
            Exit For
        End If
    Next lIndex

    HasExposedParameters = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cStrategyBasketItem.HasExposedParameters"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MakeCopy
'' Description: Make a copy of the strategy basket item
'' Inputs:      Clear ID?
'' Returns:     Copy of the strategy basket item
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MakeCopy(Optional ByVal bClearID As Boolean = True) As cStrategyBasketItem
On Error GoTo ErrSection:

    Dim aCopy As New cStrategyBasketItem ' Copy to return
    
    aCopy.SetPrivateData m, bClearID
    Set MakeCopy = aCopy

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cStrategyBasketItem.MakeCopy"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetPrivateData
'' Description: Set the private data for copying purposes
'' Inputs:      Private members to copy from, Clear ID?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Friend Sub SetPrivateData(mCopyFrom As mPrivate, ByVal bClearID As Boolean)
On Error GoTo ErrSection:

    m = mCopyFrom
    If bClearID Then
        m.lID = 0&
    End If
    
    Set m.Parms = mCopyFrom.Parms.MakeCopy
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStrategyBasketItem.SetPrivateData"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadParmsDb
'' Description: Load the parameter overrides from the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadParmsDb()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim Parm As cStrategyBasketItemParm ' Strategy basket item parameter override
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblStrategyBasketItemParms] WHERE [StrategyBasketItemID]=" & Str(m.lID) & ";", dbOpenDynaset)
    Do While Not rs.EOF
        Set Parm = New cStrategyBasketItemParm
        If Parm.LoadDbFromRs(rs) Then
            m.Parms.Add Parm, Parm.Key
        End If
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStrategyBasketItem.LoadParmsDb"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveParmsDb
'' Description: Save the parameter overrides to the database
'' Inputs:      Database
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveParmsDb(Optional ByVal DB As Database = Nothing)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim Parm As cStrategyBasketItemParm ' Strategy basket item parameter
    Dim rs As Recordset                 ' Recordset into the database
    Dim strKey As String                ' Key into the collection
    
    If DB Is Nothing Then
        Set DB = g.dbNav
    End If

    For lIndex = 1 To m.Parms.Count
        Set Parm = m.Parms(lIndex)
        
        Parm.StrategyBasketItemID = m.lID
        Parm.SaveDb DB
        
        Set m.Parms(lIndex) = Parm
    Next lIndex

    Set rs = DB.OpenRecordset("SELECT * FROM [tblStrategyBasketItemParms] WHERE [StrategyBasketItemID]=" & Str(m.lID) & ";", dbOpenDynaset)
    Do While Not rs.EOF
        If Len(rs!RuleName) = 0 Then
            strKey = rs!ParmName
        Else
            strKey = rs!RuleName & "." & rs!ParmName
        End If
    
        If m.Parms.Exists(strKey) = False Then
            rs.Delete
        End If
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStrategyBasketItem.SaveParmsDb"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BuildKey
'' Description: Build the key for the strategy basket item
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub BuildKey()
On Error GoTo ErrSection:

    m.strKey = Str(m.lStrategyID) & "|" & m.strSymbolGroupID & "|" & Str(SymbolOrSymbolID) & "|" & m.strPeriod

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStrategyBasketItem.BuildKey"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize the class members when the object is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    m.lID = 0&
    m.lStrategyBasketID = 0&
    m.lStrategyID = 0&
    m.strStrategyName = ""
    m.strSymbolGroupID = ""
    m.strSymbolGroupName = ""
    m.lSymbolID = 0&
    m.strSymbol = ""
    m.strPeriod = ""
    m.lFromDate = 0&
    m.lToDate = 0&
    m.bToEndOfData = True
    m.strToDateDisplay = ""
    m.bSplitAdjust = True
    m.dContractMultiplier = 1#

    Set m.Parms = New cGdTree

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStrategyBasketItem.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Terminate the class members when the object is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.Parms = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStrategyBasketItem.Class_Terminate"
    
End Sub

