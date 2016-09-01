VERSION 5.00
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.16#0"; "gdOCX.ocx"
Begin VB.Form frmBracketOrder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkProfitTarget 
      Caption         =   "&Profit Target: Place a Limit order above the average entry price"
      Height          =   255
      Left            =   300
      TabIndex        =   0
      Top             =   120
      Width           =   4755
   End
   Begin VB.CheckBox chkStopLoss 
      Caption         =   "&Stop Loss: Place a Stop order below the average entry price"
      Height          =   255
      Left            =   300
      TabIndex        =   6
      Top             =   1200
      Width           =   4575
   End
   Begin VB.Frame fraStopLoss 
      Height          =   1155
      Left            =   180
      TabIndex        =   7
      Top             =   1200
      Width           =   5055
      Begin VB.Frame fraStopOptions 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   255
         Left            =   840
         TabIndex        =   12
         Top             =   780
         Width           =   3015
         Begin VB.OptionButton optFixedStop 
            Caption         =   "&Fixed Stop Loss"
            Height          =   255
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton optTrailingStop 
            Caption         =   "&Trailing Stop"
            Height          =   255
            Left            =   1740
            TabIndex        =   14
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.TextBox txtStopDollars 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         TabIndex        =   11
         Text            =   "0.00"
         Top             =   375
         Width           =   915
      End
      Begin VB.TextBox txtStopPoints 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         TabIndex        =   8
         Text            =   "0.00"
         Top             =   375
         Width           =   915
      End
      Begin gdOCX.gdScrollBar sbStopPoints 
         Height          =   360
         Left            =   1740
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   337
         Width           =   210
         _ExtentX        =   370
         _ExtentY        =   635
      End
      Begin VB.Label lblStopPoints 
         Caption         =   "points  =   $"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1980
         TabIndex        =   10
         Top             =   390
         Width           =   855
      End
   End
   Begin VB.Frame fraProfitTarget 
      Height          =   855
      Left            =   180
      TabIndex        =   1
      Top             =   120
      Width           =   5055
      Begin VB.TextBox txtTargetDollars 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         TabIndex        =   5
         Text            =   "0.00"
         Top             =   330
         Width           =   915
      End
      Begin VB.TextBox txtTargetPoints 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         TabIndex        =   2
         Text            =   "0.00"
         Top             =   330
         Width           =   915
      End
      Begin gdOCX.gdScrollBar sbTargetPoints 
         Height          =   360
         Left            =   1740
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   292
         Width           =   210
         _ExtentX        =   370
         _ExtentY        =   635
      End
      Begin VB.Label lblTargetPoints 
         Caption         =   "points  =   $"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1980
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   1440
      TabIndex        =   15
      Top             =   2520
      Width           =   2595
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   495
         Left            =   1380
         TabIndex        =   17
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   495
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmBracketOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmBracketOrder.frm
'' Description: Allow the user to easily setup exit orders for a position or
''              a current entry order
''
'' Author:      Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    bOK As Boolean                      ' Did the user press OK?
    bChanging As Boolean                ' Are we currently changing a price?
    
    lSymbolID As Long                   ' Symbol ID
    strSymbol As String                 ' Symbol
    
    Bars As cGdBars                     ' Bars
    StopPoints As cPriceEditor          ' Stop points price editor
    TargetPoints As cPriceEditor        ' Target points price editor
End Type
Private m As mPrivate

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMeForOrder
'' Description: Initialize from the given order and show the form
'' Inputs:      Order
'' Returns:     True if OK pressed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMeForOrder(ByVal Order As cPtOrder) As Boolean
On Error GoTo ErrSection:

    LoadBarProperties Order.SymbolID, Order.Symbol

    EnableControls
    ShowForm Me, eForm_Modal, frmMain

    If m.bOK = True Then CreateOrdersFromOrder Order
    ShowMeForOrder = m.bOK
    
ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmBracketOrder.ShowMeForOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMeForPosition
'' Description: Initialize from the given position and show the form
'' Inputs:      Position
'' Returns:     True if OK pressed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMeForPosition(ByVal Position As cPtPosition) As Boolean
On Error GoTo ErrSection:

    LoadBarProperties Position.SymbolID, Position.Symbol

    EnableControls
    ShowForm Me, eForm_Modal, frmMain

    If m.bOK Then CreateOrdersFromPosition Position
    ShowMeForPosition = m.bOK
    
ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmBracketOrder.ShowMeForPosition"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMeForAccountPosition
'' Description: Initialize from the given account position and show the form
'' Inputs:      Account Position
'' Returns:     True if OK pressed, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMeForAccountPosition(ByVal AcctPos As cAccountPosition) As Boolean
On Error GoTo ErrSection:

    LoadBarProperties AcctPos.SymbolID, AcctPos.Symbol

    EnableControls
    ShowForm Me, eForm_Modal, frmMain

    If m.bOK Then CreateOrdersFromAccountPosition AcctPos
    ShowMeForAccountPosition = m.bOK
    
ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmBracketOrder.ShowMeForAccountPosition"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkProfitTarget_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkProfitTarget_Click()
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBracketOrder.chkProfitTarget_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    chkStopLoss_Click
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub chkStopLoss_Click()
On Error GoTo ErrSection:

    EnableControls

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBracketOrder.chkStopLoss_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Allow the ShowMe routine to unload the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdCancel_Click()
On Error GoTo ErrSection:

    m.bOK = False
    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBracketOrder.cmdCancel_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: Allow the ShowMe routine to create orders and unload the form
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdOK_Click()
On Error GoTo ErrSection:

    m.bOK = True
    Hide

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBracketOrder.cmdOK_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Initialize and setup the form upon loading
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    CenterTheForm Me
    Caption = "Exit Order Setup"
    Icon = Picture16("kBlank")
    
    chkProfitTarget.Value = vbChecked
    chkStopLoss.Value = vbChecked
    
    fraStopOptions.Visible = False
    optFixedStop.Value = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBracketOrder.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: If the user clicks on the X, allow ShowMe to unload the form
'' Inputs:      Whether to Cancel the Unload, Mode of the Unload
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrSection:

    If UnloadMode <> vbFormCode Then
        m.bOK = False
        Cancel = True
        Hide
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBracketOrder.Form_QueryUnload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Resize and move controls as the form is resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    With fraButtons
        .Move (ScaleWidth / 2) - (.Width / 2)
    End With

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EnableControls
'' Description: Enable/Disable controls as appropriate
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub EnableControls()
On Error GoTo ErrSection:

    Enable txtTargetPoints, (chkProfitTarget.Value = vbChecked)
    Enable lblTargetPoints, (chkProfitTarget.Value = vbChecked)
    Enable txtTargetDollars, (chkProfitTarget.Value = vbChecked)

    Enable txtStopPoints, (chkStopLoss.Value = vbChecked)
    Enable sbStopPoints, (chkStopLoss.Value = vbChecked)
    Enable lblStopPoints, (chkStopLoss.Value = vbChecked)
    Enable txtStopDollars, (chkStopLoss.Value = vbChecked)

    Enable optFixedStop, (chkStopLoss.Value = vbChecked)
    Enable optTrailingStop, (chkStopLoss.Value = vbChecked)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBracketOrder.EnableControls"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadBarProperties
'' Description: Load bar properties for the given symbol id or symbol
'' Inputs:      Symbol ID, Symbol
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadBarProperties(ByVal lSymbolID As Long, ByVal strSymbol As String)
On Error GoTo ErrSection:

    Dim dPrice As Double                ' Price from the ini file

    m.lSymbolID = lSymbolID
    m.strSymbol = strSymbol
    
    Set m.Bars = New cGdBars
    If m.lSymbolID = 0 Then
        SetBarProperties m.Bars, m.strSymbol
    Else
        SetBarProperties m.Bars, m.lSymbolID
    End If
    
    dPrice = GetIniFileProperty(m.Bars.Prop(eBARS_BaseSymbol) & ".Stop", m.Bars.Prop(eBARS_TickMove) * 10, "BracketOrder", g.strIniFile)
    Set m.StopPoints = New cPriceEditor
    m.StopPoints.Init sbStopPoints, txtStopPoints, m.Bars, dPrice, m.Bars.Prop(eBARS_TickMove)
    
    dPrice = GetIniFileProperty(m.Bars.Prop(eBARS_BaseSymbol) & ".Target", m.Bars.Prop(eBARS_TickMove) * 10, "BracketOrder", g.strIniFile)
    Set m.TargetPoints = New cPriceEditor
    m.TargetPoints.Init sbTargetPoints, txtTargetPoints, m.Bars, dPrice, m.Bars.Prop(eBARS_TickMove)

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBracketOrder.LoadBarProperties"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DollarsToPoints
'' Description: Calculate the number of points for the dollar amount given
'' Inputs:      Dollar Amount
'' Returns:     Number of Points
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DollarsToPoints(ByVal dDollars As Double) As Double
On Error GoTo ErrSection:

    If m.Bars.Prop(eBARS_TickValue) <> 0 Then
        DollarsToPoints = (dDollars / m.Bars.Prop(eBARS_TickValue)) * m.Bars.Prop(eBARS_TickMove)
    Else
        DollarsToPoints = 0#
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmBracketOrder.DollarsToPoints"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PointsToDollars
'' Description: Calculate the number of dollars for the point amount given
'' Inputs:      Number of Points
'' Returns:     Dollar Amount
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function PointsToDollars(ByVal dPoints As Double) As Double
On Error GoTo ErrSection:

    If m.Bars.Prop(eBARS_TickMove) <> 0 Then
        PointsToDollars = m.Bars.Prop(eBARS_TickValue) * (dPoints / m.Bars.Prop(eBARS_TickMove))
    Else
        PointsToDollars = 0#
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmBracketOrder.PointsToDollars"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtStopDollars_Change
'' Description: When the stop loss dollars change, change the stop points amount
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtStopDollars_Change()
On Error GoTo ErrSection:

    If m.bChanging = False Then
        m.bChanging = True
        m.StopPoints.Price = DollarsToPoints(Val(Trim(txtStopDollars.Text)))
        m.bChanging = False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBracketOrder.txtStopDollars_Change"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtStopDollars_GotFocus
'' Description: When the control gets focus, highlight all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtStopDollars_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtStopDollars

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBracketOrder.txtStopDollars_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtStopDollars_LostFocus
'' Description: When the control loses focus, resync the points and dollars
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtStopDollars_LostFocus()
On Error GoTo ErrSection:

    If m.bChanging = False Then
        m.bChanging = True
        txtStopDollars.Text = PointsToDollars(Val(Trim(txtStopPoints.Text)))
        m.bChanging = False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBracketOrder.txtStopDollars_LostFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtStopPoints_Change
'' Description: When the stop loss points change, change the stop dollar amount
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtStopPoints_Change()
On Error GoTo ErrSection:

    If m.bChanging = False Then
        m.bChanging = True
        txtStopDollars.Text = PointsToDollars(Val(Trim(txtStopPoints.Text)))
        m.bChanging = False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBracketOrder.txtStopPoints_Change"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtStopPoints_GotFocus
'' Description: When the control gets focus, highlight all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtStopPoints_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtStopPoints

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBracketOrder.txtStopPoints_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtTargetDollars_Change
'' Description: When the profit target dollars change, change the profit target
''              points amount
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtTargetDollars_Change()
On Error GoTo ErrSection:

    If m.bChanging = False Then
        m.bChanging = True
        m.TargetPoints.Price = DollarsToPoints(Val(Trim(txtTargetDollars.Text)))
        m.bChanging = False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBracketOrder.txtTargetDollars_Change"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtTargetDollars_GotFocus
'' Description: When the control gets focus, highlight all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtTargetDollars_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtTargetDollars

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBracketOrder.txtTargetDollars_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtTargetDollars_LostFocus
'' Description: When the control loses focus, resync the points and dollars
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtTargetDollars_LostFocus()
On Error GoTo ErrSection:

    If m.bChanging = False Then
        m.bChanging = True
        txtTargetDollars.Text = PointsToDollars(m.TargetPoints.Price)
        m.bChanging = False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBracketOrder.txtTargetDollars_LostFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtTargetPoints_Change
'' Description: When the profit target points change, change the profit target
''              dollar amount
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtTargetPoints_Change()
On Error GoTo ErrSection:

    If m.bChanging = False Then
        m.bChanging = True
        txtTargetDollars.Text = PointsToDollars(Val(Trim(txtTargetPoints.Text)))
        m.bChanging = False
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBracketOrder.txtTargetPoints_Change"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtTargetPoints_GotFocus
'' Description: When the control gets focus, highlight all of the text
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtTargetPoints_GotFocus()
On Error GoTo ErrSection:

    SelectAll txtTargetPoints

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBracketOrder.txtTargetPoints_GotFocus"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CreateOrdersFromOrders
'' Description: Create the orders from the user interface information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CreateOrdersFromOrder(ByVal Order As cPtOrder)
On Error GoTo ErrSection:

    Dim Order1 As New cPtOrder          ' First order to create
    Dim Order2 As New cPtOrder          ' Second order to create
    Dim lCurPos As Long                 ' Current position
    Dim strPos As String                ' Position info from the Trade Console

    strPos = frmTTSummary.PosEquityStr(Order.SymbolID, Order.AccountID)
    If Len(strPos) > 0 Then
        lCurPos = CLng(ValOfText(Parse(strPos, "|", 2)))
        If UCase(Parse(strPos, "|", 1)) = "LONG" Then
            If Order.Buy = False Then
                'lCurPos = Order.Quantity - lCurPos
                lCurPos = -lCurPos
            End If
        Else
            If Order.Buy = True Then
                'lCurPos = Order.Quantity - lCurPos
                lCurPos = -lCurPos
            End If
        End If
    Else
        lCurPos = 0
    End If

    'If lCurPos >= 0 Then
        If chkProfitTarget.Value = vbChecked Then
            With Order1
                .AccountID = Order.AccountID
                .AutoTradeItemID = Order.AutoTradeItemID
                .Buy = Not Order.Buy
                .Enter = False
                .GenesisOrderID = NextGenesisOrderID(g.Broker.AccountNumberForID(Order.AccountID))
                .Expiration = Order.Expiration
                .ExitPos = 100
                .LimitPrice = 0
                .OrderDate = ConvertTimeZone(Now, "", m.Bars.Prop(eBARS_ExchangeTimeZoneInf))
                .OrderType = eTT_OrderType_Limit
                .Quantity = lCurPos
                .Status = eTT_OrderStatus_TriggerPending
                .StopPrice = 0#
                .SymbolOrSymbolID = Order.Symbol
                .TriggerOrderID = Order.OrderID
                .TriggerOptions = "1,0," & Str(m.TargetPoints.Price)
                
                .Save
            End With
            
            RefreshOrder Order1
            
            SetIniFileProperty m.Bars.Prop(eBARS_BaseSymbol) & ".Target", m.TargetPoints.Price, "BracketOrder", g.strIniFile
        End If
        
        If chkStopLoss.Value = vbChecked Then
            With Order2
                .AccountID = Order.AccountID
                .AutoTradeItemID = Order.AutoTradeItemID
                .Buy = Not Order.Buy
                .CancelOrderID = Order1.OrderID
                '.CancelOrderID = 0&
                .Enter = False
                .ExitPos = 100
                .Expiration = Order.Expiration
                .GenesisOrderID = NextGenesisOrderID(g.Broker.AccountNumberForID(Order.AccountID))
                .LimitPrice = 0
                .OrderDate = ConvertTimeZone(Now, "", m.Bars.Prop(eBARS_ExchangeTimeZoneInf))
                .OrderType = eTT_OrderType_Stop
                .Quantity = lCurPos
                .Status = eTT_OrderStatus_TriggerPending
                .StopPrice = 0#
                .SymbolOrSymbolID = Order.Symbol
                .TriggerOrderID = Order.OrderID
                .TriggerOptions = "1,0," & Str(m.StopPoints.Price)
                If optFixedStop.Value = True Then
                    .TrailAmount = 0
                    .TrailOptions = ""
                Else
                    .TrailAmount = m.StopPoints.Price
                    .TrailOptions = "1"
                End If
                
                .Save
            End With
    
            RefreshOrder Order2
        
            SetIniFileProperty m.Bars.Prop(eBARS_BaseSymbol) & ".Stop", m.StopPoints.Price, "BracketOrder", g.strIniFile
        End If
    'Else
    'End If
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBracketOrder.CreateOrdersFromOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CreateOrdersFromPosition
'' Description: Create the orders from the user interface information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CreateOrdersFromPosition(ByVal Position As cPtPosition)
On Error GoTo ErrSection:

    Dim Order1 As New cPtOrder          ' First order to create
    Dim Order2 As New cPtOrder          ' Second order to create

    If chkProfitTarget.Value = vbChecked Then
        With Order1
            .AccountID = Position.AccountID
            .AutoTradeItemID = Position.AutoTradingItemID
            .Buy = (Position.Position = eTT_Position_Short)
            .Enter = False
            .ExitPos = 100
            .Expiration = -1&
            .GenesisOrderID = NextGenesisOrderID(g.Broker.AccountNumberForID(Position.AccountID))
            If .Buy Then
                .LimitPrice = Position.EntryPrice - m.TargetPoints.Price
            Else
                .LimitPrice = Position.EntryPrice + m.TargetPoints.Price
            End If
            .OrderDate = ConvertTimeZone(Now, "", m.Bars.Prop(eBARS_ExchangeTimeZoneInf))
            .OrderType = eTT_OrderType_Limit
            .Quantity = Position.EntryQuantity - Position.ExitQuantity
            .Status = eTT_OrderStatus_Open
            .StopPrice = 0#
            .SymbolOrSymbolID = Position.Symbol
            
            .Save
        End With
        SetIniFileProperty m.Bars.Prop(eBARS_BaseSymbol) & ".Target", m.TargetPoints.Price, "BracketOrder", g.strIniFile
    End If
    
    If chkStopLoss.Value = vbChecked Then
        With Order2
            .AccountID = Position.AccountID
            .AutoTradeItemID = Position.AutoTradingItemID
            .Buy = (Position.Position = eTT_Position_Short)
            .CancelOrderID = Order1.OrderID
            '.CancelOrderID = 0&
            .Enter = False
            .ExitPos = 100
            .Expiration = -1&
            .GenesisOrderID = NextGenesisOrderID(g.Broker.AccountNumberForID(Position.AccountID))
            .LimitPrice = 0
            .OrderDate = ConvertTimeZone(Now, "", m.Bars.Prop(eBARS_ExchangeTimeZoneInf))
            .OrderType = eTT_OrderType_Stop
            .Quantity = Position.EntryQuantity - Position.ExitQuantity
            .Status = eTT_OrderStatus_TriggerPending
            If .Buy Then
                .StopPrice = Position.EntryPrice + m.StopPoints.Price
            Else
                .StopPrice = Position.EntryPrice - m.StopPoints.Price
            End If
            .SymbolOrSymbolID = Position.Symbol
            If optFixedStop.Value = True Then
                .TrailAmount = 0
                .TrailOptions = ""
            Else
                .TrailAmount = m.StopPoints.Price
                .TrailOptions = "1"
            End If
            
            .Save
        End With
        SetIniFileProperty m.Bars.Prop(eBARS_BaseSymbol) & ".Stop", m.StopPoints.Price, "BracketOrder", g.strIniFile
    End If
    
    If Order1.OrderID > 0 Then SubmitOrder Order1
    If Order2.OrderID > 0 Then SubmitOrder Order2
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBracketOrder.CreateOrdersFromPosition"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CreateOrdersFromAccountPosition
'' Description: Create the orders from the user interface information
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub CreateOrdersFromAccountPosition(ByVal AcctPos As cAccountPosition)
On Error GoTo ErrSection:

    Dim Order1 As New cPtOrder          ' First order to create
    Dim Order2 As New cPtOrder          ' Second order to create

    If chkProfitTarget.Value = vbChecked Then
        With Order1
            .AccountID = AcctPos.AccountID
            .AutoTradeItemID = AcctPos.AutoTradeItemID
            .Buy = (AcctPos.CurrentPosition < 0)
            .Enter = False
            .ExitPos = 100
            .Expiration = -1&
            .GenesisOrderID = NextGenesisOrderID(g.Broker.AccountNumberForID(.AccountID))
            If .Buy Then
                .LimitPrice = AcctPos.AverageEntry - m.TargetPoints.Price
            Else
                .LimitPrice = AcctPos.AverageEntry + m.TargetPoints.Price
            End If
            .OrderDate = ConvertTimeZone(Now, "", m.Bars.Prop(eBARS_ExchangeTimeZoneInf))
            .OrderType = eTT_OrderType_Limit
            .Quantity = Abs(AcctPos.CurrentPosition)
            .Status = eTT_OrderStatus_Open
            .StopPrice = 0#
            .SymbolOrSymbolID = AcctPos.SymbolOrSymbolID
            
            .Save
        End With
        SetIniFileProperty m.Bars.Prop(eBARS_BaseSymbol) & ".Target", m.TargetPoints.Price, "BracketOrder", g.strIniFile
    End If
    
    If chkStopLoss.Value = vbChecked Then
        With Order2
            .AccountID = AcctPos.AccountID
            .AutoTradeItemID = AcctPos.AutoTradeItemID
            .Buy = (AcctPos.CurrentPosition < 0)
            .CancelOrderID = Order1.OrderID
            .Enter = False
            .ExitPos = 100
            .Expiration = -1&
            .GenesisOrderID = NextGenesisOrderID(g.Broker.AccountNumberForID(.AccountID))
            .LimitPrice = 0
            .OrderDate = ConvertTimeZone(Now, "", m.Bars.Prop(eBARS_ExchangeTimeZoneInf))
            .OrderType = eTT_OrderType_Stop
            .Quantity = Abs(AcctPos.CurrentPosition)
            .Status = eTT_OrderStatus_TriggerPending
            If .Buy Then
                .StopPrice = AcctPos.AverageEntry + m.StopPoints.Price
            Else
                .StopPrice = AcctPos.AverageEntry - m.StopPoints.Price
            End If
            .SymbolOrSymbolID = AcctPos.SymbolOrSymbolID
            If optFixedStop.Value = True Then
                .TrailAmount = 0
                .TrailOptions = ""
            Else
                .TrailAmount = m.StopPoints.Price
                .TrailOptions = "1"
            End If
            
            .Save
        End With
        SetIniFileProperty m.Bars.Prop(eBARS_BaseSymbol) & ".Stop", m.StopPoints.Price, "BracketOrder", g.strIniFile
    End If
    
    If Order1.OrderID > 0 Then SubmitOrder Order1
    If Order2.OrderID > 0 Then SubmitOrder Order2
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmBracketOrder.CreateOrdersFromAccountPosition"
    
End Sub
