VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTradeRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cTradeRules
'' Description: Wrapper class for trade entry rules and trade exit rules
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    TradeRules As cGdTree               ' Collection of Trade Rule objects
    
    cboEntryRules As ctlUniComboImageXP   'RH was ComboBox           ' Entry rules combo box
    cboExitRules As ctlUniComboImageXP            ' Exit rules combo box
    
    tblRules As cGdTable                ' Table of rule information
End Type
Private m As mPrivate

Public Property Get RulesTable() As cGdTable
    Set RulesTable = m.tblRules
End Property

Private Property Get Item(ByVal vIndexOrKey As Variant) As cTradeRule
    Set Item = m.TradeRules(vIndexOrKey)
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Load
'' Description: Load the trade rules from the appropriate files
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Load()
On Error GoTo ErrSection:
    
    ' Clear out the trade rules collection first...
    m.TradeRules.Clear
    
    ' Add the entries to the collection starting with None, then the Provided,
    ' then the Custom ones...
    AddRuleToCollection "0" & vbTab & "None" & vbTab & "" & vbTab & "", True, eGDTradeRuleType_Entry
    LoadFile AddSlash(App.Path) & "Provided\ErFilter.TXT", True, eGDTradeRuleType_Entry
    LoadFile AddSlash(App.Path) & "Custom\ErFilter.TXT", False, eGDTradeRuleType_Entry
    
    ' Add the exits to the collection starting with None, then the Provided,
    ' then the Custom ones...
    AddRuleToCollection "0" & vbTab & "None" & vbTab & "" & vbTab & "", True, eGDTradeRuleType_Exit
    LoadFile AddSlash(App.Path) & "Provided\XrFilter.TXT", True, eGDTradeRuleType_Exit
    LoadFile AddSlash(App.Path) & "Custom\XrFilter.TXT", False, eGDTradeRuleType_Exit
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeRules.Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadEntryCombo
'' Description: Load the given combo box with the entry rules
'' Inputs:      Combo Box
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadEntryCombo(cboEntryRules As ctlUniComboImageXP)
On Error GoTo ErrSection:

    If Not cboEntryRules Is Nothing Then
        Set m.cboEntryRules = cboEntryRules
        LoadCombo m.cboEntryRules, eGDTradeRuleType_Entry
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeRules.LoadEntryCombo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetEntryCombo
'' Description: Attempt to set the combo to the entry rule ID provided
'' Inputs:      Entry Rule ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetEntryCombo(ByVal lEntryRuleID As Long)
On Error GoTo ErrSection:

    SetCombo m.cboEntryRules, lEntryRuleID

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeRules.SetEntryCombo"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EntryRuleForID
'' Description: Return the entry rule for the given ID
'' Inputs:      Entry Rule ID, Abbreviate Name?
'' Returns:     Name of Rule if found, Blank if not found
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function EntryRuleForID(ByVal lEntryRuleID As Long, Optional ByVal bAbbreviation As Boolean = False) As String
On Error GoTo ErrSection:

    EntryRuleForID = RuleForID(lEntryRuleID, eGDTradeRuleType_Entry, bAbbreviation)
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeRules.EntryRuleForID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EntryComboString
'' Description: Build a string for a grid combo box from the entry rules
'' Inputs:      Selected ID
'' Returns:     Combo String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function EntryComboString(Optional ByVal lRuleID As Long = -1&) As String
On Error GoTo ErrSection:

    EntryComboString = ComboString(eGDTradeRuleType_Entry, lRuleID)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeRules.EntryComboString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadExitCombo
'' Description: Load the given combo box with the exit rules
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LoadExitCombo(cboExitRules As ctlUniComboImageXP)
On Error GoTo ErrSection:

    If Not cboExitRules Is Nothing Then
        Set m.cboExitRules = cboExitRules
        LoadCombo m.cboExitRules, eGDTradeRuleType_Exit
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeRules.LoadExitCombo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetExitCombo
'' Description: Attempt to set the combo to the exit rule ID provided
'' Inputs:      Exit Rule ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetExitCombo(ByVal lExitRuleID As Long)
On Error GoTo ErrSection:

    SetCombo m.cboExitRules, lExitRuleID

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeRules.SetExitCombo"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExitRuleForID
'' Description: Return the Exit rule for the given ID
'' Inputs:      Exit Rule ID, Abbreviate Name?
'' Returns:     Name of Rule if found, Blank if not found
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ExitRuleForID(ByVal lExitRuleID As Long, Optional ByVal bAbbreviation As Boolean = False) As String
On Error GoTo ErrSection:

    ExitRuleForID = RuleForID(lExitRuleID, eGDTradeRuleType_Exit, bAbbreviation)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeRules.ExitRuleForID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ExitComboString
'' Description: Build a string for a grid combo box from the exit rules
'' Inputs:      Selected ID
'' Returns:     Combo String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ExitComboString(Optional ByVal lRuleID As Long = -1&) As String
On Error GoTo ErrSection:

    ExitComboString = ComboString(eGDTradeRuleType_Exit, lRuleID)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeRules.ExitComboString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DescriptionForID
'' Description: Retrieve the description for the given ID
'' Inputs:      ID, Type
'' Returns:     Description
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DescriptionForID(ByVal lRuleID As Long, ByVal nType As eGDTradeRuleTypes) As String
On Error GoTo ErrSection:

    Dim strKey As String                ' Key into the collection
    Dim strReturn As String             ' Return value for the function
    
    strReturn = ""
    If nType = eGDTradeRuleType_Entry Then
        strKey = "Entry" & vbTab & Str(lRuleID)
    ElseIf nType = eGDTradeRuleType_Exit Then
        strKey = "Exit" & vbTab & Str(lRuleID)
    End If
    
    If m.TradeRules.Exists(strKey) Then
        strReturn = Item(strKey).Description
    End If
    
    DescriptionForID = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeRules.DescriptionForID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadFile
'' Description: Load a file into the collection and table
'' Inputs:      File Name, Provided?, Rule Type
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadFile(ByVal strFile As String, ByVal bProvided As Boolean, ByVal nType As eGDTradeRuleTypes)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim astrFile As cGdArray            ' Array of information read in from a file
    
    Set astrFile = New cGdArray
    
    If astrFile.FromFile(strFile) Then
        For lIndex = 0 To astrFile.Size - 1
            AddRuleToCollection astrFile(lIndex), bProvided, nType
        Next lIndex
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeRules.LoadFile"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddRuleToCollection
'' Description: Add a rule to the collection and table
'' Inputs:      Rule String, Provided?, Rule Type
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddRuleToCollection(ByVal strRule As String, ByVal bProvided As Boolean, ByVal nType As eGDTradeRuleTypes)
On Error GoTo ErrSection:

    Dim TradeRule As cTradeRule         ' Trade rule object
    Dim strKey As String                ' Key into the collection
    Dim lAddToID As Long                ' Constant amount to add to the ID in the table

    If Len(strRule) > 0 Then
        Set TradeRule = New cTradeRule
        TradeRule.FromString strRule, bProvided, nType
        
        If nType = eGDTradeRuleType_Entry Then
            strKey = "Entry" & vbTab & TradeRule.ID
            lAddToID = kEntryTradeRuleID
        ElseIf nType = eGDTradeRuleType_Exit Then
            strKey = "Exit" & vbTab & TradeRule.ID
            lAddToID = kExitTradeRuleID
        End If
        
        m.TradeRules.Add TradeRule, strKey
        
        ' Need to add the entry rule constant to the ID since the entry and exit ID's
        ' are only unique within themselves.  The ToString method of the cTradeLine
        ' class will also add the constant on because this table gets passed to the
        ' reports module and the table ID's have to be unique...
        m.tblRules.AddRecord Str(lAddToID + TradeRule.ID) & vbTab & "0" & vbTab & TradeRule.FullName, , vbTab
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeRules.AddRuleToCollection"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadCombo
'' Description: Load up the combo box
'' Inputs:      Combo Box, Type
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadCombo(cbo As ctlUniComboImageXP, ByVal nRuleType As eGDTradeRuleTypes)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim TradeRule As cTradeRule         ' Trade rule object
    
    With cbo
        .Clear
        
        For lIndex = 1 To m.TradeRules.Count
            Set TradeRule = Item(lIndex)
            If TradeRule.RuleType = nRuleType Then
                If HasModule(TradeRule.RequiredModule) Then
                    .AddItem TradeRule.FullName
                    .ItemData(.NewIndex) = TradeRule.ID
                End If
            End If
        Next lIndex
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeRules.LoadCombo"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetCombo
'' Description: Attempt to set the combo to the rule ID provided
'' Inputs:      Combo, Rule ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetCombo(cbo As ctlUniComboImageXP, ByVal lRuleID As Long)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bFound As Boolean               ' Was the entry rule found in the combo?
    
    If Not cbo Is Nothing Then
        With cbo
            bFound = False
            For lIndex = 0 To .ListCount - 1
                If .ItemData(lIndex) = lRuleID Then
                    .ListIndex = lIndex
                    bFound = True
                    Exit For
                End If
            Next lIndex
            
            If bFound = False Then
                If .ListCount > 0 Then
                    .ListIndex = 0
                Else
                    .ListIndex = -1
                End If
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeRules.SetCombo"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RuleForID
'' Description: Return the rule for the given ID
'' Inputs:      Rule ID, Type, Abbreviate Name?
'' Returns:     Name of Rule if found, Blank if not found
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RuleForID(ByVal lRuleID As Long, ByVal nRuleType As eGDTradeRuleTypes, Optional ByVal bAbbreviation As Boolean = False) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim strKey As String                ' Key into the collection
    
    strReturn = ""
    If nRuleType = eGDTradeRuleType_Entry Then
        strKey = "Entry" & vbTab & Str(lRuleID)
    ElseIf nRuleType = eGDTradeRuleType_Exit Then
        strKey = "Exit" & vbTab & Str(lRuleID)
    End If
    
    If m.TradeRules.Exists(strKey) Then
        If bAbbreviation Then
            strReturn = Item(strKey).Abbreviation
        Else
            strReturn = Item(strKey).FullName
        End If
    End If
    
    RuleForID = strReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeRules.RuleForID"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ComboString
'' Description: Build a string for a grid combo box from the rules
'' Inputs:      Type, Selected ID
'' Returns:     Combo String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ComboString(ByVal nRuleType As eGDTradeRuleTypes, Optional ByVal lRuleID As Long = -1&) As String
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim TradeRule As cTradeRule         ' Trade rule object
    Dim astrReturn As cGdArray          ' Array of information to put into the return string
    
    Set astrReturn = New cGdArray
    astrReturn.Create eGDARRAY_Strings
    
    For lIndex = 1 To m.TradeRules.Count
        Set TradeRule = Item(lIndex)
        
        If TradeRule.RuleType = nRuleType Then
            ' Only include the rule if the user has the associated module, there
            ' is no association module, or they have the rule already selected...
            If (HasModule(TradeRule.RequiredModule)) Or (TradeRule.ID = lRuleID) Then
                astrReturn.Add TradeRule.ComboString
            End If
        End If
    Next lIndex
    
    ComboString = astrReturn.JoinFields("|")

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeRules.ComboString"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize class members when the class is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.cboEntryRules = Nothing
    Set m.cboExitRules = Nothing
    
    Set m.TradeRules = New cGdTree
    
    Set m.tblRules = New cGdTable
    m.tblRules.CreateField eGDARRAY_Longs, 0, "RuleID"
    m.tblRules.CreateField eGDARRAY_Shorts, 1, "OrderTypeID"
    m.tblRules.CreateField eGDARRAY_Strings, 2, "RuleName"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeRules.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Terminate class members when the class is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.TradeRules = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeRules.Class_Terminate"
    
End Sub
