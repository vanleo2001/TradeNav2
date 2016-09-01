VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStrategyBasket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cStrategyBasket.cls
'' Description: Wrapper class for a strategy basket
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 04/03/2013   DAJ         Move Strategy Baskets into the database
'' 05/01/2013   DAJ         Shadow Trading
'' 05/14/2013   DAJ         Optionally allow loading a guru basket if not the owner
'' 07/23/2013   DAJ         Added the HasFilter function
'' 07/25/2013   DAJ         Send ToDate to cSystem.LoadForMultRun
'' 08/09/2013   DAJ         Set order by clause when loading items from database
'' 08/13/2013   DAJ         Make sure to tell DM_GetBars to split adjust stock data
'' 10/04/2013   DAJ         Only validate checksums on strategy basket items once
'' 05/05/2014   DAJ         Allow FractZen bars for strategy baskets
'' 05/05/2014   DAJ         Don't allow user to run FractZen bars if not enabled
'' 08/19/2014   DAJ         Expose Strategy Basket Item Inputs
'' 06/12/2015   DAJ         Don't add symbols to stream during run if more than kSN_BASKETLIMIT symbols
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    lID As Long                         ' Stratey Basket ID
    strName As String                   ' Name of the strategy basket
    strDescription As String            ' Description for the strategy basket
    dLastModified As Double             ' Last modified date/time
    strRequiredModule As String         ' Required module for this strategy basket
    lLibraryID As Long                  ' ID of the containing library
    lSecurityLevel As Long              ' Security level for the basket
    bCannotDelete As Boolean            ' User cannot delete it?
    strPassword As String               ' Password for the basket
    bIsGuru As Boolean                  ' Is this a Guru strategy basket?
    
    bStopRun As Boolean                 ' Stop the run?
    
    Items As cGdTree                    ' Collection of strategy basket items
End Type
Private m As mPrivate

Public Property Get ID() As Long
    ID = m.lID
End Property

Public Property Get Name() As String
    Name = m.strName
End Property
Public Property Let Name(ByVal strName As String)
    m.strName = strName
End Property

Public Property Get Description() As String
    Description = m.strDescription
End Property
Public Property Let Description(ByVal strDescription As String)
    m.strDescription = strDescription
End Property

Public Property Get LastModified() As Double
    LastModified = m.dLastModified
End Property

Public Property Get RequiredModule() As String
    RequiredModule = m.strRequiredModule
End Property
Public Property Let RequiredModule(ByVal strRequiredModule As String)
    m.strRequiredModule = strRequiredModule
End Property

Public Property Get LibraryID() As Long
    LibraryID = m.lLibraryID
End Property
Public Property Let LibraryID(ByVal lLibraryID As Long)
    m.lLibraryID = lLibraryID
End Property

Public Property Get SecurityLevel() As Long
    SecurityLevel = m.lSecurityLevel
End Property
Public Property Let SecurityLevel(ByVal lSecurityLevel As Long)
    m.lSecurityLevel = lSecurityLevel
End Property

Public Property Get CannotDelete() As Boolean
    CannotDelete = m.bCannotDelete
End Property
Public Property Let CannotDelete(ByVal bCannotDelete As Boolean)
    m.bCannotDelete = bCannotDelete
End Property

Public Property Get Password() As String
    Password = m.strPassword
End Property
Public Property Let Password(ByVal strPassword As String)
    m.strPassword = strPassword
End Property

Public Property Get IsGuru() As Boolean
    IsGuru = m.bIsGuru
End Property
Public Property Let IsGuru(ByVal bIsGuru As Boolean)
    m.bIsGuru = bIsGuru
End Property

Public Property Get StopRun() As Boolean
    StopRun = m.bStopRun
End Property
Public Property Let StopRun(ByVal bStopRun As Boolean)
    m.bStopRun = bStopRun
End Property

Public Property Get Items() As cGdTree
    Set Items = m.Items
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadFile
'' Description: Attempt to load the strategy basket from the given file
'' Inputs:      File to load
'' Returns:     True if succesful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LoadFile(ByVal strFileName As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim astrFile As cGdArray            ' Array of lines from the file
    Dim lIndex As Long                  ' Index into a for loop
    Dim strKey As String                ' Key in the key/value pair
    Dim strValue As String              ' Value in the key/value pair
    Dim basketItem As cStrategyBasketItem   ' Strategy basket item
    Dim astrOverrides As cGdArray       ' Overrides
    Dim lIndex2 As Long                 ' Index into a for loop
    Dim astrItem As cGdArray            ' Child item information in an array
    Dim astrFields As cGdArray          ' Fields in a string
    Dim Item2 As cStrategyBasketItem    ' Strategy basket item
    
    m.Items.Clear
    
    bReturn = False
    If FileExist(strFileName) Then
        Set astrFile = New cGdArray
        If astrFile.FromFile(strFileName) Then
            m.strName = FileBase(strFileName)
            
            For lIndex = 0 To astrFile.Size - 1
                strKey = Parse(astrFile(lIndex), "=", 1)
                strValue = Parse(astrFile(lIndex), "=", 2)
                
                If UCase(strKey) = "DESCRIPTION" Then
                    m.strDescription = strValue
                ElseIf UCase(strKey) = "REQUIRED" Then
                    m.strRequiredModule = strValue
                ElseIf UCase(strKey) <> "CHILDREN" Then
                    Set basketItem = New cStrategyBasketItem
                    basketItem.FromString astrFile(lIndex)
                    If basketItem.Validate Then
                        AddItem basketItem
                        
                        'If g.SymbolPool Is Nothing Then
                        '    FillInSymbolGroupForItemFromFile basketItem
                        'ElseIf g.SymbolPool.NumRecords = 0 Then
                        '    FillInSymbolGroupForItemFromFile basketItem
                        'Else
                            FillInSymbolGroupForItemFromPool basketItem
                        'End If
                        
                        If (lIndex + 1) <= (astrFile.Size - 1) Then
                            If UCase(Left(astrFile(lIndex + 1), 9)) = "CHILDREN=" Then
                                Set astrOverrides = New cGdArray
                                
                                astrOverrides.SplitFields Mid(astrFile(lIndex + 1), 10), "|"
                                For lIndex2 = 0 To astrOverrides.Size - 1
                                    Set astrFields = New cGdArray
                                    astrFields.SplitFields astrOverrides(lIndex2), ";"
                                    
                                    Set astrItem = New cGdArray
                                    astrItem.SplitFields astrFile(lIndex), vbTab
                                    astrItem(10) = astrFields(2)
                                    If astrFields.Size > 3 Then
                                        If Len(astrFields(3)) > 0 Then
                                            astrItem(12) = astrFields(3)
                                        End If
                                    End If
                                    
                                    Set Item2 = New cStrategyBasketItem
                                    Item2.FromString astrItem.JoinFields(vbTab)
                                    If Item2.Validate Then
                                        Item2.SymbolID = CLng(Val(astrFields(0)))
                                        Item2.Symbol = astrFields(1)
                                        
                                        AddItem Item2
                                    End If
                                Next lIndex2
                                
                                lIndex = lIndex + 1
                            End If
                        End If
                    End If
                End If
            Next lIndex
            
            bReturn = True
        End If
    End If
    
    LoadFile = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cStrategyBasket.LoadFile"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadDb
'' Description: Attempt to load the strategy basket from the database
'' Inputs:      Strategy Basket ID, Include Hidden Strategy Basket if IDE?, Include
''              even if not the owner?
'' Returns:     True if succesful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LoadDb(ByVal lStrategyBasketID As Long, Optional ByVal bIncludeHiddenIfIde As Boolean = False, Optional ByVal bIncludeGuruIfNotOwner As Boolean = False, Optional ByVal bLoadBasketItems As Boolean = True) As Boolean
On Error GoTo ErrSection:

    LoadDb = LoadDbFromRs(BasketRecord(lStrategyBasketID), bIncludeHiddenIfIde, bIncludeGuruIfNotOwner, bLoadBasketItems)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cStrategyBasket.LoadDb"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadDbFromRs
'' Description: Attempt to load the strategy basket from the database
'' Inputs:      Recordset, Include Hidden Strategy Basket if IDE?, Include
''              even if not the owner?
'' Returns:     True if succesful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LoadDbFromRs(ByVal rs As Recordset, Optional ByVal bIncludeHiddenIfIde As Boolean = False, _
        Optional ByVal bIncludeGuruIfNotOwner As Boolean = False, Optional ByVal bLoadBasketItems As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    
    bReturn = False
    
    If mSysNav.IncludeStrategyBasketsFromRecordset(rs, bIncludeHiddenIfIde, bIncludeGuruIfNotOwner) Then
        m.lID = rs!StrategyBasketID
        m.strName = rs!Name
        m.strDescription = rs!Description
        m.dLastModified = rs![tblStrategyBaskets.LastModified]
        m.strRequiredModule = rs![tblStrategyBaskets.RequiredMod]
        m.lLibraryID = rs![tblStrategyBaskets.LibraryID]
        m.lSecurityLevel = rs![tblStrategyBaskets.SecurityLevel]
        m.bCannotDelete = rs![tblStrategyBaskets.CannotDelete]
        m.strPassword = DecryptField(rs![tblStrategyBaskets.Password])
        m.bIsGuru = rs![tblStrategyBaskets.IsGuru]
        
        If bLoadBasketItems Then
            LoadItemsDb
        End If

        bReturn = True
    End If

    LoadDbFromRs = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cStrategyBasket.LoadDbFromRs"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveDb
'' Description: Save this strategy basket to the database
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

    If (m.dLastModified = 0) Or (FileExist(App.Path & "\LastMod.SKP") = False) Then
        m.dLastModified = Now()
    End If

    Set rs = BasketRecord(, DB, False)
    If (rs.BOF And rs.EOF) Then
        rs.AddNew
    Else
        rs.Edit
    End If
    
    m.lID = rs!StrategyBasketID
    rs!Name = m.strName
    rs!Description = m.strDescription
    rs!LastModified = m.dLastModified
    rs!RequiredMod = m.strRequiredModule
    rs!LibraryID = m.lLibraryID
    rs!SecurityLevel = m.lSecurityLevel
    rs!CannotDelete = m.bCannotDelete
    EncryptField rs!Password, m.strPassword
    
    If ItemExists(rs, "IsGuru") Then
        rs!IsGuru = m.bIsGuru
    End If
        
    rs!CheckSum = BuildCheckSum(rs, "tblStrategyBaskets")
    rs.Update
    
    SaveItemsDb DB
    
    If FormIsLoaded("frmTradeItems") Then
        frmTradeItems.BasketChanged Me
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStrategyBasket.SaveDb"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DeleteDb
'' Description: Delete this strategy basket from the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub DeleteDb()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database

    Set rs = BasketRecord
    If Not (rs.BOF And rs.EOF) Then
        rs.Delete
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStrategyBasket.DeleteDb"
    
End Sub

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
    
    For lIndex = 1 To m.Items.Count
        m.Items(lIndex).StrategyBasketID = 0&
        m.Items(lIndex).ClearID
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStrategyBasket.ClearID"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Run
'' Description: Run the strategy basket
'' Inputs:      Next Bar Report?, Just Selected Item, Automated trading parent ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Run(ByVal bNextBarReport As Boolean, Optional ByVal lJustSelectedID As Long = -1&, Optional lAutoTradeParentID As Long = -1&)
On Error GoTo ErrSection:

    Dim strErrMsg As String             ' Error message to display to the user
    Dim bContinue As Boolean            ' Continue?
    Dim lIndex As Long                  ' Index into a for loop
    Dim ItemsToRun As cGdTree           ' Collection of strategy basket items to run
    Dim Item As cStrategyBasketItem     ' Strategy basket item
    Dim astrParms As cGdArray           ' Parameter array
    Dim astrTrades As cGdArray          ' Trades array
    Dim dLoadThru As Double             ' Load thru date
    Dim dNextBarDate As Double          ' Date for the next bar report
    Dim bAssumeNoPosition As Boolean    ' Does the user want to assume no current position?
    Dim bIgnoreNextBarData As Boolean   ' Does the user want to ignore next bar data?
    Dim astrFiles As New cGdArray       ' Array of next bar files
    Dim bPyramid As Boolean             ' Do any of the systems pyramid?
    Dim System As New cSystem           ' System object
    Dim dLoadFrom As Double             ' Date to load from
    Dim Bars As New cGdBars             ' Temporary Bars structure
    Dim strFileName As String           ' Name of the trades file
    Dim rc As Long                      ' Return code from the optimizer
    Dim bIsOwner As Boolean             ' Does the current user have "owner" rights?
    Dim strSymbol$
    Dim dLoadTo#
    Dim strPeriod As String             ' Period to use
    Dim bHasFractZen As Boolean         ' Are there any items with FractZen bars?
    Dim TradeItem As cAutoTradeItem     ' Automated trading item
    Dim strOverrides As String          ' Overrides
    
    bContinue = True
    If FormIsLoaded("frmOptimizer") Then
        If frmOptimizer.InProgress Then
            If InStr(UCase(frmOptimizer.Caption), "BASKET") > 0 Then
                strErrMsg = "Cannot run a Strategy Basket while another Strategy Basket is running."
            Else
                strErrMsg = "Cannot run a Strategy Basket while a strategy is being optimized."
            End If
            InfBox strErrMsg, "!", , "Optimizer"
            
            bContinue = False
        End If
    End If
    
    If bContinue Then
        bHasFractZen = False
        Set ItemsToRun = New cGdTree
        For lIndex = 1 To m.Items.Count
            Set Item = m.Items(lIndex)
            If (Len(Item.Symbol) > 0) And (Item.ContractMultiplier > 0) Then
                If (lJustSelectedID = -1&) Or (Item.ID = lJustSelectedID) Then
                    ItemsToRun.Add Item
                    
                    If UCase(Item.Period) = "FRACTZEN" Then
                        bHasFractZen = True
                    End If
                End If
            End If
        Next lIndex
        
        If ItemsToRun.Count = 0 Then
            InfBox "There are no valid Symbol/Strategy|combinations in this basket to run", "!", , "Strategy Basket Error"
        ElseIf bHasFractZen And Not g.FractZen.Allowed Then
            InfBox "You are not enabled to run strategy basket items with FractZen bars", "!", , "Strategy Basket Error"
        Else
            Set astrParms = New cGdArray
            astrParms.Create eGDARRAY_Strings
            Set astrTrades = New cGdArray
            astrTrades.Create eGDARRAY_Strings
            
            dLoadThru = 0#
            
            If m.bIsGuru = True Then
                bIsOwner = IsOwnerOfGuruObject(m.lLibraryID)
            Else
                bIsOwner = True
            End If
            
            ' Run each system/symbol pair
            If bNextBarReport Then
                dNextBarDate = GetNextBarDate(bAssumeNoPosition, bIgnoreNextBarData)
                If dNextBarDate = -99999# Then
                    StatusMsg ""
                    m.bStopRun = False
                    bContinue = False
                End If
                dLoadThru = dNextBarDate
            Else
                frmOptimizer.Init ItemsToRun.Count, astrParms, eGDOptMode_StrategyBasket, , Not bIsOwner
            End If
            
            If bContinue Then
                astrFiles.Create eGDARRAY_Strings
                bPyramid = False
                    
                If bNextBarReport Then
                    Screen.MousePointer = vbHourglass
                End If
                
                For lIndex = 1 To ItemsToRun.Count
                    Set Item = ItemsToRun(lIndex)
                    Set System = New cSystem
                    strSymbol = Item.Symbol
                    strPeriod = Item.Period
                    
#If 1 Then
                    ' need to use Item.Split ???
                    With Item
                        strOverrides = .Overrides
                        If lAutoTradeParentID <> -1& Then
                            Set TradeItem = g.TradingItems.GetItemForBasketItem(lAutoTradeParentID, Item.ID)
                            If Not TradeItem Is Nothing Then
                                strOverrides = TradeItem.CombineOverrides(TradeItem.Overrides)
                            End If
                        End If
                        
                        System.LoadBarsForMultRun .StrategyID, .Symbol, .Period, .FromDate, .ToEndOfData, .ToDate, strOverrides, bNextBarReport, (ItemsToRun.Count <= kSN_BASKETLIMIT)
                    End With
#Else
                    ' Load the system first to get the "bars to load" value...
                    System.Load Item.StrategyID
                    
                    If (GetPeriodicity(Item.Period) > (ePRD_Days + 1)) And (UCase(Item.Period) <> "FRACTZEN") Then
                        dLoadFrom = 0#
                    Else
                        dLoadFrom = Int(Item.FromDate - (System.BarsLoadedBeforeTrading + 1) * 1.46)
                    End If
                    If Item.ToEndOfData Then
                        dLoadTo = dLoadThru
                    Else
                        dLoadTo = Item.ToDate
                    End If
                    
                    If Item.SymbolID <> 0 Then
                        If UCase(Item.Period) = "FRACTZEN" Then
                            If Item.ToEndOfData Then
                                strPeriod = AutoBreakoutPeriod(Item.Symbol)
                            Else
                                strPeriod = AutoBreakoutPeriod(Item.Symbol, Item.ToDate)
                            End If
                        Else
                            strPeriod = Item.Period
                        End If
                        
                        DM_GetBars Bars, Item.SymbolOrSymbolID, strPeriod, dLoadFrom, dLoadTo, , , Not Item.Split
                        g.RealTime.SpliceBars Bars
                    ElseIf InStr(strSymbol, "|") > 0 Then
                        strPeriod = Item.Period
                        
                        ' load data from external file
                        Bars.FromFile Parse(strSymbol, "|", 5), Parse(strSymbol, "|", 2), Parse(strSymbol, "|", 1), strPeriod, , dLoadFrom, dLoadTo
                        strSymbol = Parse(strSymbol, "|", 1)
                        GetMarketInfo strSymbol, Bars
                    End If
                
                    System.LoadForMultRun Item.StrategyID, Bars, Item.FromDate, Item.ToEndOfData, Item.ToDate, Item.Overrides
#End If
                    If bNextBarReport = False Then
                        If System.Bars.Size > 0 Then
                            System.Test , , astrTrades, False
            
                            ' TLB: Dave, call this for the contract multiplier for each item ...
                            MultiplyTrades astrTrades, Item.ContractMultiplier
            
                            strFileName = Replace(System.NextBarFile(eGDNextBarMode_RunMult), "\NB", "\S")
                            astrTrades.ToFile strFileName
                            astrFiles.Add strFileName
                            If System.Pyramid = True Then
                                bPyramid = True
                            End If
                        Else
                            astrTrades.Clear
                        End If
                        
                        astrParms(0) = vbTab & "<system>" & vbTab & System.SystemName & vbTab & "true"
                        'astrParms(1) = vbTab & "<symbol>" & vbTab & Item.Symbol & vbTab & "true"
                        astrParms(1) = vbTab & "<symbol>" & vbTab & strSymbol & vbTab & "true"
                        astrParms(2) = vbTab & "<period>" & vbTab & strPeriod & vbTab & "true"
                        astrParms(3) = vbTab & "<fromdate>" & vbTab & Str(Item.FromDate) & vbTab & "true"
                        astrParms(4) = vbTab & "<todate>" & vbTab & Str(Item.ToDate) & vbTab & "true"
                        astrParms(5) = vbTab & "<toend>" & vbTab & Str(Item.ToEndOfData) & vbTab & "true"
                        astrParms(6) = vbTab & "<overrides>" & vbTab & strOverrides & vbTab & "true"
                        astrParms(7) = vbTab & "<filenames>" & vbTab & strFileName & vbTab & "true"
                        
                        rc = frmOptimizer.Add(lIndex, astrTrades, astrParms)
                        If rc <> kSN_OPTIMIZATION_IN_PROGRESS Then
                            Exit For
                        End If
                    Else
                        StatusMsg "Testing " & Replace(System.SystemName, "&", "&&") & " on " & strSymbol ' Item.Symbol
                        System.NextBarReport eGDNextBarMode_RunMult, dNextBarDate, bAssumeNoPosition, bIgnoreNextBarData
                        astrFiles.Add System.NextBarFile(eGDNextBarMode_RunMult)
                        DoEvents
                        If m.bStopRun Then
                            Exit For
                        End If
                    End If
                Next lIndex
                
                If bNextBarReport Then
                    Screen.MousePointer = vbDefault
                    m.bStopRun = False
                    StatusMsg ""
                    frmNextBar.ShowMeMult astrFiles, , Name
                Else
                    frmOptimizer.SetUpMergedRun Name, bPyramid, astrFiles.ArrayHandle
                End If
            End If
        End If
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    If FormIsLoaded("frmOptimizer") Then
        frmOptimizer.StopRun
    End If
    RaiseError "cStrategyBasket.Run"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HasFilter
'' Description: Is there an item that contains a filter?
'' Inputs:      None
'' Returns:     True if item with a filter, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function HasFilter() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim Item As cStrategyBasketItem     ' Strategy basket item
    
    bReturn = False
    For lIndex = 1 To m.Items.Count
        Set Item = m.Items(lIndex)
        If Len(Item.SymbolGroupID) > 0 Then
            If Left(Item.SymbolGroupID, 4) = "FIL:" Then
                bReturn = True
                Exit For
            End If
        End If
    Next lIndex
    
    HasFilter = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cStrategyBasket.HasFilter"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    MakeCopy
'' Description: Make a copy of the strategy basket
'' Inputs:      None
'' Returns:     Copy of the strategy basket
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MakeCopy() As cStrategyBasket
On Error GoTo ErrSection:

    Dim aCopy As New cStrategyBasket    ' Copy to return
    
    aCopy.SetPrivateData m
    Set MakeCopy = aCopy

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cStrategyBasket.MakeCopy"

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

    m = mCopyFrom
    m.lID = 0&
    
    Set m.Items = mCopyFrom.Items.MakeCopy
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStrategyBasket.SetPrivateData"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BasketRecord
'' Description: Return the recordset for the given strategy basket ID
'' Inputs:      Strategy Basket ID, Database, Include Libraries?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function BasketRecord(Optional ByVal lStrategyBasketID As Long = kNullData, Optional DB As Database = Nothing, Optional ByVal bIncludeLibraries As Boolean = True) As Recordset
On Error GoTo ErrSection:

    If lStrategyBasketID = kNullData Then
        lStrategyBasketID = m.lID
    End If

    Set BasketRecord = mSysNav.LoadStrategyBasketsRecordset(lStrategyBasketID, DB, bIncludeLibraries)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cStrategyBasket.BasketRecord"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadItemsDb
'' Description: Load the strategy basket items from the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadItemsDb()
On Error GoTo ErrSection:

    Dim rs As Recordset                 ' Recordset into the database
    Dim basketItem As cStrategyBasketItem   ' Strategy basket item
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblStrategyBasketItems] WHERE [StrategyBasketID]=" & Str(m.lID) & " ORDER BY [SymbolGroupName],[Symbol];", dbOpenDynaset)
    ValidateCheckSums rs, "tblStrategyBasketItems"
    
    Do While Not rs.EOF
        If rs!CheckSum <> 0.5 Then
            Set basketItem = New cStrategyBasketItem
            If basketItem.LoadDbFromRs(rs, False) Then
                AddItem basketItem
            End If
        End If
        
        rs.MoveNext
    Loop
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStrategyBasket.LoadItemsDb"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveItemsDb
'' Description: Save the strategy basket items to the database
'' Inputs:      Database
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SaveItemsDb(Optional DB As Database = Nothing)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim Item As cStrategyBasketItem     ' Strategy basket item object
    Dim rs As Recordset                 ' Recordset into the database
    Dim bFound As Boolean               ' Was the item found?
    
    If DB Is Nothing Then
        Set DB = g.dbNav
    End If

    For lIndex = 1 To m.Items.Count
        Set Item = m.Items(lIndex)
        
        Item.StrategyBasketID = m.lID
        Item.SaveDb DB
        
        Set m.Items(lIndex) = Item
    Next lIndex
    
    Set rs = DB.OpenRecordset("SELECT * FROM [tblStrategyBasketItems] WHERE [StrategyBasketID]=" & Str(m.lID) & ";", dbOpenDynaset)
    Do While Not rs.EOF
        bFound = False
        For lIndex = 1 To m.Items.Count
            If m.Items(lIndex).ID = rs!StrategyBasketItemID Then
                bFound = True
                Exit For
            End If
        Next lIndex
        
        If bFound = False Then
            rs.Delete
        End If
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStrategyBasket.SaveItemsDb"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillInSymbolGroupForItemFromPool
'' Description: Fill in the items for a symbol group based on the given item
'' Inputs:      Strategy Basket Item
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FillInSymbolGroupForItemFromPool(ByVal Item As cStrategyBasketItem)
On Error GoTo ErrSection:

    Dim lFieldNum As Long               ' Field number for the group in the symbol pool
    Dim aIndex As cGdArray              ' Indexed list of items in the group
    Dim lIndex As Long                  ' Index into a for loop
    Dim lSymbolID As Long               ' Symbol ID
    Dim ChildItem As cStrategyBasketItem    ' Strategy basket item for the child

    If (Len(Item.SymbolGroupID) > 0) And (Len(Item.Symbol) = 0) Then
        lFieldNum = g.SymbolPool.FieldNumForID(Item.SymbolGroupID)
        If lFieldNum >= 0 Then
            Set aIndex = g.SymbolPool.ArrayTable.CreateIndex(lFieldNum)
            For lIndex = 0 To aIndex.Size - 1
                lSymbolID = g.SymbolPool.SymbolID(aIndex(lIndex))
                
                Set ChildItem = Item.MakeCopy
                ChildItem.SymbolOrSymbolID = lSymbolID
                
                If m.Items.Exists(ChildItem.Key) = False Then
                    AddItem ChildItem
                End If
            Next lIndex
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStrategyBasket.FillInSymbolGroupForItemFromPool"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FillInSymbolGroupForItemFromFile
'' Description: Fill in the items for a symbol group based on the given item
'' Inputs:      Strategy Basket Item
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FillInSymbolGroupForItemFromFile(ByVal Item As cStrategyBasketItem)
On Error GoTo ErrSection:

    Dim lFieldNum As Long               ' Field number for the group in the symbol pool
    Dim aIndex As cGdArray              ' Indexed list of items in the group
    Dim lIndex As Long                  ' Index into a for loop
    Dim lSymbolID As Long               ' Symbol ID
    Dim ChildItem As cStrategyBasketItem    ' Strategy basket item for the child
    Dim Group As cSymbolGroup           ' Symbol group object

    If (Len(Item.SymbolGroupID) > 0) And (Len(Item.Symbol) = 0) Then
        Set Group = New cSymbolGroup
        
        If Group.FromFile("", Parse(Item.SymbolGroupID, ":", 2), True) Then
            Set aIndex = Group.SymbolIDs
            For lIndex = 0 To aIndex.Size - 1
                lSymbolID = CLng(Val(aIndex(lIndex)))
                
                Set ChildItem = Item.MakeCopy
                ChildItem.SymbolOrSymbolID = lSymbolID
                
                If m.Items.Exists(ChildItem.Key) = False Then
                    AddItem ChildItem
                End If
            Next lIndex
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStrategyBasket.FillInSymbolGroupForItemFromFile"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddItem
'' Description: Add the given strategy basket item to the collection if it
''              doesn't exist
'' Inputs:      Item, Replace if Exist?
'' Returns:     Index for the item ( -1 if not added )
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AddItem(ByVal Item As cStrategyBasketItem, Optional ByVal bReplaceIfExist As Boolean = True) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    
    lReturn = -1&
    If m.Items.Exists(Item.Key) Then
        If bReplaceIfExist Then
            Set m.Items(Item.Key) = Item
            lReturn = m.Items.Index(Item.Key)
        End If
    Else
        lReturn = m.Items.Add(Item, Item.Key)
    End If
    
    AddItem = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cStrategyBasket.AddItem"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetNextBarDate
'' Description: Ask the user for the date to use for next bar reports
'' Inputs:      Assume No Position, Ignore Next Bar Data
'' Returns:     Next Bar Date
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetNextBarDate(bAssumeNoPosition As Boolean, bIgnoreNextBarData As Boolean) As Double
On Error GoTo ErrSection:

    Dim dNewYorkTime As Double          ' Current date and time in New York
    Dim dNextBarDate As Double          ' Date (and time) of the next bar report
    Dim lMousePointer As Long           ' Current state of the mouse pointer

    lMousePointer = Screen.MousePointer
    Screen.MousePointer = vbDefault

    ' Come up with an educated guess as to the next bar date...
    dNewYorkTime = ConvertTimeZone(Now)
    If Hour(dNewYorkTime) < 14 Then
        dNextBarDate = Int(dNewYorkTime)
    Else
        dNextBarDate = Int(dNewYorkTime) + 1
    End If
    Do While Not IsWeekday(dNextBarDate)
        dNextBarDate = dNextBarDate + 1
    Loop
    
    ' Verify our educated guess with the user...
    If frmNextBarOpt.ShowMe(dNextBarDate, False, False, bAssumeNoPosition, bIgnoreNextBarData) Then
        GetNextBarDate = dNextBarDate
    Else
        GetNextBarDate = -99999#
    End If

ErrExit:
    Screen.MousePointer = lMousePointer
    Exit Function
    
ErrSection:
    Screen.MousePointer = lMousePointer
    RaiseError "cStrategyBasket.GetNextBarDate"
    
End Function

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

    m.lID = 0&
    m.strName = ""
    m.strDescription = ""
    m.strRequiredModule = ""
    m.lLibraryID = mSysNav.kSN_UserLibrary
    m.lSecurityLevel = 0&
    m.bCannotDelete = False
    m.strPassword = ""
    m.bIsGuru = False
    
    Set m.Items = New cGdTree

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStrategyBasket.Class_Initialize"
    
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

    Set m.Items = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStrategyBasket.Class_Terminate"
    
End Sub

