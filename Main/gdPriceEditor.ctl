VERSION 5.00
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.UserControl gdPriceEditor 
   BackStyle       =   0  'Transparent
   ClientHeight    =   510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1365
   ScaleHeight     =   510
   ScaleWidth      =   1365
   Begin HexUniControls.ctlUniFrameWL fraPriceEditor 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1275
      _ExtentX        =   0
      _ExtentY        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "gdPriceEditor.ctx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "gdPriceEditor.ctx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "gdPriceEditor.ctx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniTextBoxXP txtPriceEditor 
         Height          =   285
         Left            =   0
         TabIndex        =   1
         Top             =   38
         Width           =   1020
         _ExtentX        =   0
         _ExtentY        =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Enabled         =   -1  'True
         Locked          =   0   'False
         Text            =   "gdPriceEditor.ctx":0068
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   0
         MultiLine       =   0   'False
         Alignment       =   2
         ScrollBars      =   0
         PasswordChar    =   ""
         TrapTab         =   0   'False
         EnableContextMenu=   -1  'True
         RaiseChangeEvent=   -1  'True
         Tip             =   "gdPriceEditor.ctx":0098
         HideSelection   =   -1  'True
         RightToLeft     =   0   'False
         ManualStart     =   0   'False
         RoundedBorders  =   0   'False
         MousePointer    =   0
         MouseIcon       =   "gdPriceEditor.ctx":00B8
      End
      Begin gdOCX.gdScrollBar sbPriceEditor 
         Height          =   360
         Left            =   1020
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Width           =   210
         _ExtentX        =   370
         _ExtentY        =   635
      End
   End
End
Attribute VB_Name = "gdPriceEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        gdPriceEditor.ctl
'' Description: Custom control for a price editor
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Public Event Changed()

Private Type mPrivate
    Bars As cGdBars                     ' Internal Bars object for price display / min-move
    bAsTradingUnits As Boolean          ' Display as trading units?
    dPrice As Double                    ' Current price
    dMinMove As Double                  ' Minimum movement
    bChanged As Boolean                 ' Has the price changed?
    bShowIfZero As Boolean              ' Show the price in the text box if zero?
    bScrollbarValueSetInCode As Boolean ' Was the scrollbar value set in code?
    dPrevScrollBarValue As Double       ' Previous scrollbar value
End Type
Private m As mPrivate

Public Property Get Price() As Variant
    Price = m.dPrice
End Property
Public Property Let Price(ByVal vNewPrice As Variant)
    FixPrice vNewPrice
    PropertyChanged "Price"
End Property

Public Property Get AsTradingUnits() As Boolean
    AsTradingUnits = m.bAsTradingUnits
End Property
Public Property Let AsTradingUnits(ByVal bAsTradingUnits As Boolean)
    m.bAsTradingUnits = bAsTradingUnits
    FixPrice
    PropertyChanged "AsTradingUnits"
End Property

Public Property Get MinMove() As Double
    MinMove = m.dMinMove
End Property
Public Property Let MinMove(ByVal dMinMove As Double)
    m.dMinMove = dMinMove
    
    If m.dMinMove = 0# Then
        m.dMinMove = 0.01
    End If
    
    If m.dMinMove = 1 Then
        Set m.Bars = Nothing
        m.dMinMove = 1#
    Else
        Set m.Bars = New cGdBars
        m.Bars.Prop(eBARS_TickMove) = m.dMinMove
        m.Bars.Prop(eBARS_MinMoveInTicks) = 1
    End If
    
    FixPrice txtPriceEditor.Text
    PropertyChanged "MinMove"
End Property

Public Property Get Min() As Double
    Min = -sbPriceEditor.Max
End Property
Public Property Let Min(ByVal dMin As Double)
    sbPriceEditor.Max = -Round(dMin / m.dMinMove)
    PropertyChanged "Min"
End Property

Public Property Get Max() As Double
    Max = -sbPriceEditor.Min
End Property
Public Property Let Max(ByVal dMax As Double)
    sbPriceEditor.Min = -Round(dMax / m.dMinMove)
    PropertyChanged "Max"
End Property

Public Property Get ShowIfZero() As Boolean
    ShowIfZero = m.bShowIfZero
End Property
Public Property Let ShowIfZero(ByVal bShowIfZero As Boolean)
    m.bShowIfZero = bShowIfZero
    FixPrice
    PropertyChanged "ShowIfZero"
End Property

Private Property Get ScrollBarValue() As Double
    ScrollBarValue = sbPriceEditor.Value
End Property
Private Property Let ScrollBarValue(ByVal dNewValue As Double)
    m.bScrollbarValueSetInCode = True
    sbPriceEditor.Value = dNewValue
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Init
'' Description: Initialize the control
'' Inputs:      Bars, Price, Min, Max, Show if Zero, Min Move
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Init(Bars As cGdBars, Optional ByVal dPrice# = 0, Optional ByVal dMin# = 0, Optional ByVal dMax# = 999999, Optional ByVal bShowIfZero As Boolean = False, Optional ByVal dMinMove As Double = 0#)

    Set m.Bars = Bars
    m.bShowIfZero = bShowIfZero
    
    If (Bars Is Nothing) And (dMinMove <> 0#) And (dMinMove <> 1#) Then
        Set m.Bars = New cGdBars
        m.Bars.Prop(eBARS_TickMove) = dMinMove
        m.Bars.Prop(eBARS_MinMoveInTicks) = 1
    End If
    
    If Not m.Bars Is Nothing Then
        m.dMinMove = m.Bars.MinMove(Date)
    ElseIf dMinMove = 0# Then
        m.dMinMove = 1
    Else
        m.dMinMove = dMinMove
    End If
    If m.dMinMove = 0 Then m.dMinMove = 0.01
    
    With sbPriceEditor
        .TabStop = False
        
        .Min = -Round(dMax / m.dMinMove)
        .Max = -Round(dMin / m.dMinMove)
        
        If .Min >= .Max Then
            txtPriceEditor.Locked = True
            .Visible = False
        Else
            txtPriceEditor.Locked = False
            .Visible = True
        End If
    End With
    
    FixPrice dPrice

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    sbPriceEditor_Change
'' Description: Handle a change in the scroll bar
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub sbPriceEditor_Change()
                    
    Dim dDiff As Double                 ' Difference in the scroll bar value
                    
    If m.bScrollbarValueSetInCode Then
        FixPrice -ScrollBarValue * m.dMinMove
    Else
        If m.bChanged Then
            dDiff = -(ScrollBarValue - m.dPrevScrollBarValue)
            FixPrice txtPriceEditor.Text, dDiff
        Else
            FixPrice -ScrollBarValue * m.dMinMove
        End If
        
        MoveFocus txtPriceEditor
    End If
    
    m.dPrevScrollBarValue = ScrollBarValue
    m.bScrollbarValueSetInCode = False

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPriceEditor_Change
'' Description: Handle a change in the text box
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPriceEditor_Change()

    If m.Bars Is Nothing Then
        FixPrice txtPriceEditor.Text
    Else
        m.bChanged = True
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPriceEditor_GotFocus
'' Description: Handle the text box getting the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPriceEditor_GotFocus()

    m.bChanged = False
    SelectAll txtPriceEditor

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPriceEditor_KeyDown
'' Description: Handle the user pressing a key in the text box
'' Inputs:      Key Code, Shift/Ctrl/Alt status
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPriceEditor_KeyDown(KeyCode As Integer, Shift As Integer)

    If Shift = 0 Then
        If KeyCode = vbKeyUp Then
            KeyCode = 0
            
            If m.bChanged Then
                FixPrice txtPriceEditor.Text
                m.bChanged = False
            End If
            
            If (ScrollBarValue > sbPriceEditor.Min) And (txtPriceEditor.Locked = False) Then
                ScrollBarValue = ScrollBarValue - 1
            End If
        ElseIf KeyCode = vbKeyDown Then
            KeyCode = 0
            
            If m.bChanged Then
                FixPrice txtPriceEditor.Text
                m.bChanged = False
            End If
            
            If (ScrollBarValue < sbPriceEditor.Max) And (txtPriceEditor.Locked = False) Then
                ScrollBarValue = ScrollBarValue + 1
            End If
        End If
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPriceEditor_KeyPress
'' Description: Handle the user pressing a key in the text box
'' Inputs:      Ascii Key Pressed
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPriceEditor_KeyPress(KeyAscii As Integer, Shift As Integer)

    If KeyAscii = vbKeyReturn Then
        If m.bChanged Then
            FixPrice txtPriceEditor.Text
            m.bChanged = False
        End If
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    txtPriceEditor_LostFocus
'' Description: Handle the text box losing the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtPriceEditor_LostFocus()

    If m.bChanged Then
        FixPrice txtPriceEditor.Text
        m.bChanged = False
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    UserControl_LostFocus
'' Description: Handle the control losing the focus
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UserControl_LostFocus()

    If m.bChanged Then
        FixPrice txtPriceEditor.Text
        m.bChanged = False
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FixPrice
'' Description: Fix the new price
'' Inputs:      New price, Adjustment
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FixPrice(Optional vNewPrice As Variant, Optional ByVal dAdjustment As Double = 0#)

    Dim dValue As Double                ' Value to display
    Dim strText As String               ' Text to display
    Dim dPrevPrice As Double            ' Previous price
    Static bInProgress As Boolean       ' Are we already doing something?

    If m.dMinMove = 0 Then
        txtPriceEditor.Text = ""
    ElseIf bInProgress = False Then
        bInProgress = True
        dPrevPrice = m.dPrice
    
        ' get new price if passed
        If Not IsMissing(vNewPrice) Then
            If (VarType(vNewPrice) = vbString) And (Not m.Bars Is Nothing) Then
                m.dPrice = m.Bars.PriceFromString(vNewPrice)
            Else
                m.dPrice = Val(vNewPrice)
            End If
        End If
        
        ' set value of scrollbar, which automatically takes
        ' Min and Max into account as well as rounding to min move
        dValue = -(Round(m.dPrice / m.dMinMove) + dAdjustment)
        If ScrollBarValue <> dValue Then
            ScrollBarValue = dValue
        End If
        m.dPrice = -ScrollBarValue * m.dMinMove
        
        ' display price
        If (m.dPrice = -999999) Or ((m.dPrice = 0) And (sbPriceEditor.Max = 0) And (m.bShowIfZero = False)) Then
            txtPriceEditor.Text = ""
        ElseIf m.Bars Is Nothing Then
            If m.dMinMove = 0.25 Then
                strText = Format(m.dPrice, "0.00")
            ElseIf m.dMinMove = 0.5 Then
                strText = Format(m.dPrice, "0.0")
            Else
                strText = Str(m.dPrice)
            End If
        Else
            strText = m.Bars.PriceDisplay(m.dPrice, m.bAsTradingUnits)
        End If
        If strText <> txtPriceEditor.Text Then
            txtPriceEditor.Text = strText
            txtPriceEditor.SelStart = Len(txtPriceEditor.Text)
        End If
        
        If m.dPrice <> dPrevPrice Then
            RaiseEvent Changed
        End If
        
        bInProgress = False
    End If

End Sub

