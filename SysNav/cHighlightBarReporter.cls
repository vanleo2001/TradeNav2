VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHighlightBarReporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cHighlightBarReporter.cls
'' Description: Class to handle the highlight bar reporter
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 06/22/2011   DAJ         Always set optimization info on strategy and allow override of max bars
'' 06/27/2011   DAJ         Utilize the VerifyHighlightBarReport function in frmRule
'' 12/27/2011   DAJ         Change number of bars input on both exit rules
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kFunctionName As String = "TNHBR Condition"
Private Const kLongEntryName As String = "TNHBR Long Entry"
Private Const kShortEntryName As String = "TNHBR Short Entry"
Private Const kStrategyName As String = "Trade Navigator Highlight Bar Reporter"

Private Type mPrivate
    Strategy As cSystem                 ' Strategy to run through the engine
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RunFromBars
'' Description: Run the hightlight bar reporter
'' Inputs:      Condition, Long/Short, Bars, Max Bars
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RunFromBars(ByVal strCondition As String, ByVal bLong As Boolean, ByVal Bars As cGdBars, Optional ByVal lMaxBars As Long = 30&) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lEntryRuleID As Long            ' Entry rule ID
    
    bReturn = False
    If VerifyCondition(strCondition) Then
        lEntryRuleID = ChangeEntryRule(strCondition, bLong)
        If lEntryRuleID > 0 Then
            If LoadStrategy(Bars) Then
                SetOptimizationValues bLong, lMaxBars
                m.Strategy.HighlightBarReport = True
                If bLong Then
                    m.Strategy.HighlightBarInfo = "Long|" & strCondition
                Else
                    m.Strategy.HighlightBarInfo = "Short|" & strCondition
                End If
                m.Strategy.Test False, lEntryRuleID
            Else
                InfBox "Could not load the Hightlight Bar Reporter strategy", "!", , "Highlight Bar Reporter"
            End If
        End If
    End If
    
    RunFromBars = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cHighlightBarReporter.RunFromBars"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    RunFromSymbol
'' Description: Run the hightlight bar reporter
'' Inputs:      Condition, Long/Short, Symbol, Period, Start Date, End Date, Max Bars
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function RunFromSymbol(ByVal strCondition As String, ByVal bLong As Boolean, ByVal strSymbol As String, ByVal strPeriod As String, Optional ByVal lStartDate As Long = 0&, Optional ByVal lEndDate As Long = 0&, Optional ByVal lMaxBars As Long = 30&) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim Bars As cGdBars                 ' Bars object
    
    bReturn = False
    Set Bars = New cGdBars
    If DM_GetBars(Bars, strSymbol, strPeriod, lStartDate, lEndDate) Then
        bReturn = RunFromBars(strCondition, bLong, Bars, lMaxBars)
    Else
        InfBox "Cannot load '" & strPeriod & "' data for '" & strSymbol & "'", "!", , "Highlight Bar Reporter"
    End If
    
    RunFromSymbol = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cHighlightBarReporter.RunFromSymbol"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    VerifyCondition
'' Description: Verify that the given condition is valid for this purpose
'' Inputs:      Condition, Show Message?
'' Returns:     True if valid, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function VerifyCondition(ByVal strCondition As String, Optional ByVal bShowMessage As Boolean = True) As Boolean
On Error GoTo ErrSection:

    Dim strMessage As String            ' Message to display to the user
    
    strMessage = ""
    If IsBooleanExpression(strCondition) = False Then
        strMessage = "Cannot run the Hightlight Bar Reporter on a non-boolean expression"
    End If
    
    If (Len(strMessage) > 0) And (bShowMessage = True) Then
        InfBox strMessage, "!", , "Highlight Bar Reporter"
    End If
    
    VerifyCondition = (Len(strMessage) = 0)

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cHighlightBarReporter.VerifyCondition"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadStrategy
'' Description: Load the strategy
'' Inputs:      Bars
'' Returns:     True if successful, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LoadStrategy(ByVal Bars As cGdBars) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lStrategyID As Long             ' Database ID for the strategy
    
    bReturn = False
    lStrategyID = SystemIDForName(kStrategyName)
    If lStrategyID > 0 Then
        Set m.Strategy = New cSystem
        m.Strategy.Load lStrategyID
        m.Strategy.Bars = Bars
        
        bReturn = True
    End If
    
    LoadStrategy = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cHighlightBarReporter.LoadStrategy"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeEntryRule
'' Description: Change the appropriate entry rule
'' Inputs:      Condition, Long/Short
'' Returns:     Rule ID (Zero if not found or loaded)
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ChangeEntryRule(ByVal strCondition As String, ByVal bLong As Boolean) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim strName As String               ' Rule name
    Dim RuleForm As frmRule             ' Rule form
    Dim rs As Recordset                 ' Recordset into the database
    Dim strError As String              ' Rule verification error
    
    lReturn = 0&
    
    If bLong Then
        strName = kLongEntryName
    Else
        strName = kShortEntryName
    End If
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] WHERE [Name]='" & strName & "';", dbOpenDynaset)
    If Not (rs.BOF And rs.EOF) Then
        Set RuleForm = New frmRule
        Load RuleForm
        If RuleForm.LoadRec(rs!RuleID, rs!SystemNumber, True) Then
            RuleForm.Editor1.Text = strCondition
            If RuleForm.VerifyHighlightBarReport(strError) = True Then
                If RuleForm.Save("ID_Save", False) Then
                    lReturn = rs!RuleID
                End If
            Else
                InfBox "Error verifying condition:|" & strError, "!", , "Highlight Bar Reporter"
            End If
        Else
            InfBox "Could not change the Hightlight Bar Reporter rule", "!", , "Highlight Bar Reporter"
        End If
    End If
    
    ChangeEntryRule = lReturn
            
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cHighlightBarReporter.ChangeEntryRule"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetOptimizationValues
'' Description: Set the optimization values on the input in the strategy
'' Inputs:      Long/Short, Max Value
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetOptimizationValues(ByVal bLong As Boolean, Optional ByVal lMaxBars As Long = 30&)
On Error GoTo ErrSection:

    Dim lRule As Long                   ' Index into a for loop
    Dim lInput As Long                  ' Index into a for loop
    
    If Not m.Strategy Is Nothing Then
        For lRule = 1 To m.Strategy.Rules.Count
            ' 12/27/2011 DAJ: Since the exit rules are linked, we can change both (when we
            ' do the following line and the user runs the "short" option with a custom
            ' number of bars, it only does the stock 30 days -- Issue #6556)...
            'If m.Strategy.Rules.Item(lRule).BuySell = Not bLong Then
                If m.Strategy.Rules.Item(lRule).RuleType = 1 Then
                    For lInput = 1 To m.Strategy.Rules.Item(lRule).Inputs.Count
                        If m.Strategy.Rules.Item(lRule).Inputs.Item(lInput).ParmTypeID <> 5 Then
                            m.Strategy.Rules.Item(lRule).Inputs.Item(lInput).IfOptimize = True
                            m.Strategy.Rules.Item(lRule).Inputs.Item(lInput).OptFromValue = 1
                            m.Strategy.Rules.Item(lRule).Inputs.Item(lInput).OptToValue = lMaxBars
                            m.Strategy.Rules.Item(lRule).Inputs.Item(lInput).OptStepValue = 1
                            
                            Exit For
                        End If
                    Next lInput
                End If
            'End If
        Next lRule
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cHighlightBarReporter.SetOptimizationValues"
    
End Sub
