VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStrategyBasketItemParmGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cStrategyBasketItemParmGrid.cls
'' Description: Class to handle a grid of strategy basket item parameters
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private WithEvents m_fgParms As VSFlexGrid
Attribute m_fgParms.VB_VarHelpID = -1

Private Enum eGDCols
    eGDCol_RuleID = 0
    eGDCol_ParmID
    eGDCol_Expose
    eGDCol_RuleName
    eGDCol_ParmName
    eGDCol_Override
    eGDCol_Value
    eGDCol_DefaultValue
    eGDCol_Lookup
    eGDCol_BasketValue
    eGDCol_NumCols
End Enum

Private Type mPrivate
   bStrategyBasketItem As Boolean       ' Is this being used for a strategy basket item form?
   bLoading As Boolean                  ' Are we currently loading?
   
   basketItem As cStrategyBasketItem    ' Strategy basket item
End Type
Private m As mPrivate

Private Function GDCol(ByVal Col As eGDCols) As Long
    GDCol = Col
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitForStrategyBasketItem
'' Description: Initialize for a strategy basket item
'' Inputs:      Grid, Strategy Basket Item
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub InitForStrategyBasketItem(Grid As VSFlexGrid, basketItem As cStrategyBasketItem)
On Error GoTo ErrSection:

    m.bStrategyBasketItem = True
    m.bLoading = True
    
    Set m_fgParms = Grid
    Set m.basketItem = basketItem
    
    InitGrid
    LoadGrid m.basketItem.StrategyID
    SetOverrides m.basketItem.Parms
    
    m.bLoading = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStrategyBasketItemParmGrid.InitForStrategyBasketItem"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitForAutoTradeItem
'' Description: Initialize for an automated trading item
'' Inputs:      Grid, Automated trading item
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub InitForAutoTradeItem(Grid As VSFlexGrid, AutoTradeItem As cAutoTradeItem)
On Error GoTo ErrSection:

    m.bStrategyBasketItem = False
    m.bLoading = True
    
    Set m_fgParms = Grid
    Set m.basketItem = New cStrategyBasketItem
    m.basketItem.LoadDb AutoTradeItem.StrategyBasketItemID
    
    InitGrid
    LoadGrid m.basketItem.StrategyID
    SetOverrides m.basketItem.Parms, Parse(AutoTradeItem.Overrides, "|", 2)
    
    m.bLoading = False

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStrategyBasketItemParmGrid.InitForAutoTradeItem"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeStrategy
'' Description: Change the strategy
'' Inputs:      New Strategy ID
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ChangeStrategy(ByVal lNewStrategyID As Long)
On Error GoTo ErrSection:

    LoadGrid lNewStrategyID

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStrategyBasketItemParmGrid.ChangeStrategy"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    BasketItemParmsFromGrid
'' Description: Set the basket item parameters based on the grid
'' Inputs:      Basket Item Parameters
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub BasketItemParmsFromGrid(Parms As cGdTree)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim bLinkInputs As Boolean          ' Are the inputs linked?
    Dim Parm As cStrategyBasketItemParm ' Strategy basket item parameter

    If Not Parms Is Nothing Then
        Parms.Clear
        
        With m_fgParms
            bLinkInputs = .ColHidden(GDCol(eGDCol_RuleName))
            
            For lIndex = .FixedRows To .Rows - 1
                If .RowHidden(lIndex) = False Then
                    If CheckedCell(m_fgParms, lIndex, GDCol(eGDCol_Override)) Or CheckedCell(m_fgParms, lIndex, GDCol(eGDCol_Expose)) Then
                        Set Parm = New cStrategyBasketItemParm
                        
                        If bLinkInputs Then
                            Parm.RuleName = ""
                        Else
                            Parm.RuleName = .TextMatrix(lIndex, GDCol(eGDCol_RuleName))
                        End If
                        Parm.ParmName = .TextMatrix(lIndex, GDCol(eGDCol_ParmName))
                        Parm.Value = .TextMatrix(lIndex, GDCol(eGDCol_Value))
                        Parm.IsOverridden = CheckedCell(m_fgParms, lIndex, GDCol(eGDCol_Override))
                        Parm.IsExposed = CheckedCell(m_fgParms, lIndex, GDCol(eGDCol_Expose))
                        
                        Parms.Add Parm, Parm.Key
                    End If
                End If
            Next lIndex
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStrategyBasketItemParmGrid.BasketItemParmsFromGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetOverrides
'' Description: Get the overrides from the grid
'' Inputs:      None
'' Returns:     Overrides string
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetOverrides() As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim astrBasket As cGdArray          ' Array of basket overrides
    Dim astrAutoTrade As cGdArray       ' Array of automated trading overrides
    Dim lIndex As Long                  ' Index into a for loop
    Dim lPos As Long                    ' Position in the array
    Dim strKey As String                ' Key into the array
    
    Set astrBasket = New cGdArray
    astrBasket.Create eGDARRAY_Strings
    
    Set astrAutoTrade = New cGdArray
    astrAutoTrade.Create eGDARRAY_Strings
    
    With m_fgParms
        For lIndex = .FixedRows To .Rows - 1
            strKey = .TextMatrix(lIndex, GDCol(eGDCol_Lookup)) & "="
            
            If m.bStrategyBasketItem = True Then
                If .TextMatrix(lIndex, GDCol(eGDCol_Value)) <> .TextMatrix(lIndex, GDCol(eGDCol_DefaultValue)) Then
                    If astrBasket.BinarySearch(strKey, lPos, eGdSort_MatchUsingSearchStringLength) = False Then
                        astrBasket.Add strKey & .TextMatrix(lIndex, GDCol(eGDCol_Value)), lPos
                    End If
                End If
            Else
                If .TextMatrix(lIndex, GDCol(eGDCol_BasketValue)) <> .TextMatrix(lIndex, GDCol(eGDCol_DefaultValue)) Then
                    If astrBasket.BinarySearch(strKey, lPos, eGdSort_MatchUsingSearchStringLength) = False Then
                        astrBasket.Add strKey & .TextMatrix(lIndex, GDCol(eGDCol_BasketValue)), lPos
                    End If
                End If
            
                If .TextMatrix(lIndex, GDCol(eGDCol_Value)) <> .TextMatrix(lIndex, GDCol(eGDCol_BasketValue)) Then
                    If astrAutoTrade.BinarySearch(strKey, lPos, eGdSort_MatchUsingSearchStringLength) = False Then
                        astrAutoTrade.Add strKey & .TextMatrix(lIndex, GDCol(eGDCol_Value)), lPos
                    End If
                End If
            End If
        Next lIndex
    End With
    
    strReturn = astrBasket.JoinFields(",")
    If m.bStrategyBasketItem = False Then
        strReturn = strReturn & "|" & astrAutoTrade.JoinFields(",")
    End If
    
    GetOverrides = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cStrategyBasketItemParmGrid.GetOverrides"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgParms_AfterEdit
'' Description: Do some processing after a user edit
'' Inputs:      Row and Column of Edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgParms_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrSection:

    If m.bLoading = False Then
        Select Case Col
            Case GDCol(eGDCol_Override)
                If CheckedCell(m_fgParms, Row, Col) = False Then
                    m_fgParms.TextMatrix(Row, GDCol(eGDCol_Value)) = m_fgParms.TextMatrix(Row, GDCol(eGDCol_DefaultValue))
                End If
            
            Case GDCol(eGDCol_Value)
                If m_fgParms.TextMatrix(Row, GDCol(eGDCol_Value)) = m_fgParms.TextMatrix(Row, GDCol(eGDCol_DefaultValue)) Then
                    CheckedCell(m_fgParms, Row, GDCol(eGDCol_Override)) = False
                Else
                    CheckedCell(m_fgParms, Row, GDCol(eGDCol_Override)) = True
                End If
            
        End Select
        
        SyncParms Row
    End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStrategyBasketItemParmGrid.m_fgParms_AfterEdit"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgParms_AfterRowColChange
'' Description: Put user in edit mode if changed to the value column
'' Inputs:      Old Row and Column, New Row and Column
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgParms_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
On Error GoTo ErrSection:

    If NewCol = GDCol(eGDCol_Value) Then
        m_fgParms.EditCell
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStrategyBasketItemParmGrid.m_fgParms_AfterRowColChange"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgParms_BeforeEdit
'' Description: Only allow the user to edit the override and value columns
'' Inputs:      Row and Column of Edit, Whether to Cancel the Edit
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgParms_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
On Error GoTo ErrSection:

    Select Case Col
        Case GDCol(eGDCol_Expose), GDCol(eGDCol_Override), GDCol(eGDCol_Value)
        
        Case Else
            Cancel = True
            
    End Select

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStrategyBasketItemParmGrid.m_fgParms_BeforeEdit"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    m_fgParms_MouseUp
'' Description: If the user just turned on the override column, move the focus
''              to the value column
'' Inputs:      Button Pressed, Shift/Ctrl/Alt status, Location of Click
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub m_fgParms_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrSection:

    Dim lMouseRow As Long               ' Current grid row clicked on
    Dim lMouseCol As Long               ' Current grid column clicked on

    With m_fgParms
        If Button = vbLeftButton Then
            lMouseRow = .MouseRow
            lMouseCol = .MouseCol
            
            If lMouseCol = GDCol(eGDCol_Override) Then
                If CheckedCell(m_fgParms, lMouseRow, lMouseCol) = True Then
                    .Row = lMouseRow
                    .Col = GDCol(eGDCol_Value)
                End If
            End If
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStrategyBasketItemParmGrid.m_fgParms_MouseUp"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    InitGrid
'' Description: Initialize the grid
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub InitGrid()
On Error GoTo ErrSection:

    With m_fgParms
        .Redraw = flexRDNone
        
        .AllowBigSelection = False
        .AllowSelection = False
        .BackColorAlternate = ALT_GRID_ROW_COLOR
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExNone
        .ExtendLastCol = True
        .HighLight = flexHighlightNever
        .ScrollTrack = True
        .SelectionMode = flexSelectionListBox
        .SheetBorder = RGB(128, 128, 128)
        
        .Rows = 1
        .FixedRows = 1
        .Cols = GDCol(eGDCol_NumCols)
        .FixedCols = 0
        
        .TextMatrix(0, GDCol(eGDCol_RuleID)) = "Rule ID"
        .TextMatrix(0, GDCol(eGDCol_ParmID)) = "Parm ID"
        .TextMatrix(0, GDCol(eGDCol_Expose)) = "Expose"
        .TextMatrix(0, GDCol(eGDCol_RuleName)) = "Rule"
        .TextMatrix(0, GDCol(eGDCol_ParmName)) = "Input"
        .TextMatrix(0, GDCol(eGDCol_Override)) = "Customize"
        .TextMatrix(0, GDCol(eGDCol_Value)) = "Value"
        .TextMatrix(0, GDCol(eGDCol_DefaultValue)) = "Default"
        .TextMatrix(0, GDCol(eGDCol_Lookup)) = "Lookup"
        .TextMatrix(0, GDCol(eGDCol_BasketValue)) = "Basket Value"
        
        .ColHidden(GDCol(eGDCol_RuleID)) = True
        .ColHidden(GDCol(eGDCol_ParmID)) = True
        .ColHidden(GDCol(eGDCol_Expose)) = Not m.bStrategyBasketItem
        .ColHidden(GDCol(eGDCol_DefaultValue)) = True
        .ColHidden(GDCol(eGDCol_Lookup)) = True
        .ColHidden(GDCol(eGDCol_BasketValue)) = True
        
        .ColDataType(GDCol(eGDCol_Override)) = flexDTBoolean
        .ColDataType(GDCol(eGDCol_Expose)) = flexDTBoolean
        
        .ColAlignment(GDCol(eGDCol_Value)) = flexAlignLeftTop
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignLeftTop
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = flexRDBuffered
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStrategyBasketItemParmGrid.InitGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGrid
'' Description: Load the grid based on the strategy that was selected
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadGrid(ByVal lStrategyID As Long)
On Error GoTo ErrSection:

    Dim Strategy As cSystem             ' Strategy object
    Dim nRedraw As RedrawSettings       ' Redraw settings for the grid
    Dim lRule As Long                   ' Index into a for loop
    Dim Rule As cRule                   ' Rule object
    Dim lInput As Long                  ' Index into a for loop
    Dim Parm As cInput                  ' Input object

    With m_fgParms
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Rows = .FixedRows
    
        Set Strategy = New cSystem
        Strategy.Load lStrategyID
        
        For lRule = 1 To Strategy.Rules.Count
            Set Rule = Strategy.Rules.Item(lRule)
            
            If (Rule.Alternate = True) Or (Rule.Selected = True) Then
                For lInput = 1 To Rule.Inputs.Count
                    Set Parm = Rule.Inputs.Item(lInput)
                    
                    If Parm.ParmTypeID <> 5 Then
                        .Rows = .Rows + 1
                        
                        .TextMatrix(.Rows - 1, GDCol(eGDCol_ParmID)) = Str(Parm.ParmID)
                        .TextMatrix(.Rows - 1, GDCol(eGDCol_RuleID)) = Str(Parm.RuleID)
                        .TextMatrix(.Rows - 1, GDCol(eGDCol_RuleName)) = Parm.RuleName
                        .TextMatrix(.Rows - 1, GDCol(eGDCol_ParmName)) = Parm.ParmName
                        .TextMatrix(.Rows - 1, GDCol(eGDCol_Value)) = Str(Parm.Value)
                        .TextMatrix(.Rows - 1, GDCol(eGDCol_DefaultValue)) = Str(Parm.Value)
                        .TextMatrix(.Rows - 1, GDCol(eGDCol_BasketValue)) = Str(Parm.Value)
                        
                        If Strategy.LinkInputs Then
                            .TextMatrix(.Rows - 1, GDCol(eGDCol_Lookup)) = Parm.ParmName
                        Else
                            .TextMatrix(.Rows - 1, GDCol(eGDCol_Lookup)) = Parm.RuleName & "." & Parm.ParmName
                        End If
                        
                        CheckedCell(m_fgParms, .Rows - 1, GDCol(eGDCol_Override)) = False
                        CheckedCell(m_fgParms, .Rows - 1, GDCol(eGDCol_Expose)) = False
                    End If
                Next lInput
            End If
            
            .ColHidden(GDCol(eGDCol_RuleName)) = Strategy.LinkInputs
        Next lRule
        
        If Strategy.LinkInputs Then
            HideInputs
        End If
        
        .AutoSize 0, .Cols - 1, False, 75
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStrategyBasketItemParmGrid.LoadGrid"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetOverrides
'' Description: Set the overrides in the grid
'' Inputs:      Strategy Basket Item Parameters, Overrides
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetOverrides(basketItemParms As cGdTree, Optional ByVal strOverrides As String = "")
On Error GoTo ErrSection:

    Dim nRedraw As Long                 ' Current state of the grid's redraw
    Dim lIndex As Long                  ' Index into a for loop
    Dim ItemParm As cStrategyBasketItemParm ' Strategy basket item parameter
    Dim strKey As String                ' Key into the collection
    Dim astrOverrides As New cGdArray   ' Array of user overrides
    Dim lPos As Long                    ' Position in the array of the lookup string
    Dim bHideRow As Boolean             ' Hide the row?
    
    astrOverrides.SplitFields strOverrides, ","
    astrOverrides.Sort
    
    With m_fgParms
        nRedraw = .Redraw
        .Redraw = flexRDNone
        
        For lIndex = .FixedRows To .Rows - 1
            strKey = .TextMatrix(lIndex, GDCol(eGDCol_Lookup))
            
            If basketItemParms.Exists(strKey) Then
                Set ItemParm = basketItemParms(strKey)
                                                
                CheckedCell(m_fgParms, lIndex, GDCol(eGDCol_Override)) = ItemParm.IsOverridden
                If ItemParm.IsOverridden Then
                    .TextMatrix(lIndex, GDCol(eGDCol_Value)) = Str(ItemParm.Value)
                    .TextMatrix(lIndex, GDCol(eGDCol_BasketValue)) = Str(ItemParm.Value)
                End If
                CheckedCell(m_fgParms, lIndex, GDCol(eGDCol_Expose)) = ItemParm.IsExposed
                
                bHideRow = ((m.bStrategyBasketItem = False) And (ItemParm.IsExposed = False))
            Else
                bHideRow = (m.bStrategyBasketItem = False)
            End If
            
            If .RowHidden(lIndex) = False Then
                .RowHidden(lIndex) = bHideRow
            End If
            
            If astrOverrides.Size > 0 Then
                strKey = strKey & "="
                If astrOverrides.BinarySearch(strKey, lPos, eGdSort_MatchUsingSearchStringLength) = True Then
                    .TextMatrix(lIndex, GDCol(eGDCol_Value)) = Parse(astrOverrides(lPos), "=", 2)
                    CheckedCell(m_fgParms, lIndex, GDCol(eGDCol_Override)) = True
                End If
            End If
        Next lIndex
        
        .Redraw = nRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStrategyBasketItemParmGrid.SetOverrides"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    HideInputs
'' Description: If the inputs are linked, only show one of each
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HideInputs()
On Error GoTo ErrSection:

    Dim lRedraw As Long                 ' Current state of the grid's redraw
    Dim lIndex As Long                  ' Index into a for loop
    Dim strLastParm As String           ' Name of the last parameter
    
    With m_fgParms
        lRedraw = .Redraw
        .Redraw = flexRDNone
        
        .Col = GDCol(eGDCol_ParmName)
        .Sort = flexSortStringAscending
        
        strLastParm = ""
        For lIndex = .FixedRows To .Rows - 1
            If .TextMatrix(lIndex, GDCol(eGDCol_ParmName)) = strLastParm Then
                .RowHidden(lIndex) = True
            Else
                strLastParm = .TextMatrix(lIndex, GDCol(eGDCol_ParmName))
                .RowHidden(lIndex) = False
            End If
        Next lIndex
        
        .AutoSize 0, .Cols - 1, False, 75
        SetBackColors m_fgParms
        .Redraw = lRedraw
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStrategyBasketItemParmGrid.HideInputs"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SyncParms
'' Description: Synchronize parameters if in Linked Inputs mode
'' Inputs:      Row to Synchronize with
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SyncParms(ByVal Row As Long)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    With m_fgParms
        If .ColHidden(GDCol(eGDCol_RuleName)) = True Then
            For lIndex = .FixedRows To .Rows - 1
                If lIndex <> Row Then
                    If .TextMatrix(lIndex, GDCol(eGDCol_ParmName)) = .TextMatrix(Row, GDCol(eGDCol_ParmName)) Then
                        CheckedCell(m_fgParms, lIndex, GDCol(eGDCol_Override)) = CheckedCell(m_fgParms, Row, GDCol(eGDCol_Override))
                        .TextMatrix(lIndex, GDCol(eGDCol_Value)) = .TextMatrix(Row, GDCol(eGDCol_Value))
                    End If
                End If
            Next lIndex
        End If
    End With

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cStrategyBasketItemParmGrid.SyncParms"
    
End Sub

