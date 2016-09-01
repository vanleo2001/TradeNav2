VERSION 5.00
Object = "{0A09F193-58DB-11D4-B9AB-005004C2D746}#1.17#0"; "gdOCX.ocx"
Object = "{94728B7C-47F7-43C9-9036-7A36A9487035}#1.2#0"; "HexUniControls42.ocx"
Begin VB.Form frmAssignFillToLot 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin HexUniControls.ctlUniFrameWL fraButtons 
      Height          =   495
      Left            =   2235
      TabIndex        =   5
      Top             =   960
      Width           =   2535
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
      Caption         =   "frmAssignFillToLot.frx":0000
      Enabled         =   -1  'True
      ForeColor       =   -2147483642
      BackColor       =   -2147483633
      Tip             =   "frmAssignFillToLot.frx":002C
      VistaStyle      =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAssignFillToLot.frx":004C
      RightToLeft     =   0   'False
      Begin HexUniControls.ctlUniButtonImageXP cmdCancel 
         Cancel          =   -1  'True
         Height          =   495
         Left            =   1320
         TabIndex        =   0
         Top             =   0
         Width           =   1215
         _ExtentX        =   0
         _ExtentY        =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmAssignFillToLot.frx":0068
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAssignFillToLot.frx":0096
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAssignFillToLot.frx":00B6
         RightToLeft     =   0   'False
      End
      Begin HexUniControls.ctlUniButtonImageXP cmdOK 
         Default         =   -1  'True
         Height          =   495
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1215
         _ExtentX        =   0
         _ExtentY        =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "frmAssignFillToLot.frx":00D2
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         ShowFocus       =   -1  'True
         Tristate        =   0   'False
         Pressed         =   0   'False
         Tip             =   "frmAssignFillToLot.frx":00F8
         Style           =   -1
         RoundedBorders  =   -1  'True
         xTranspColor    =   0
         yTranspColor    =   0
         MousePointer    =   0
         MouseIcon       =   "frmAssignFillToLot.frx":0118
         RightToLeft     =   0   'False
      End
   End
   Begin HexUniControls.ctlUniComboImageXP cboLots 
      Height          =   315
      Left            =   4140
      TabIndex        =   4
      Top             =   420
      Width           =   2655
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ButtonBackColor =   -2147483633
      ButtonForeColor =   -2147483630
      ButtonStyle     =   -1
      SelectorStyle   =   -1
      SelBackColor    =   -2147483635
      SelForeColor    =   -2147483634
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tip             =   "frmAssignFillToLot.frx":0134
      Sorted          =   0   'False
      HScroll         =   0   'False
      RoundedBorders  =   -1  'True
      IconDim         =   16
      MousePointer    =   0
      MouseIcon       =   "frmAssignFillToLot.frx":0154
      DropDownOnTextClick=   -1  'True
      DropDownWidth   =   -1
      RightToLeft     =   0   'False
   End
   Begin HexUniControls.ctlUniTextBoxXP txtQuantity 
      Height          =   315
      Left            =   1620
      TabIndex        =   1
      Top             =   420
      Width           =   780
      _ExtentX        =   0
      _ExtentY        =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      Locked          =   0   'False
      Text            =   "frmAssignFillToLot.frx":0170
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
      Tip             =   "frmAssignFillToLot.frx":019A
      HideSelection   =   -1  'True
      RightToLeft     =   0   'False
      ManualStart     =   0   'False
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAssignFillToLot.frx":01BA
   End
   Begin gdOCX.gdScrollBar sbQuantity 
      Height          =   360
      Left            =   2400
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   397
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   635
   End
   Begin HexUniControls.ctlUniLabelXP lblLot 
      Height          =   255
      Left            =   2880
      Top             =   450
      Width           =   1215
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
      Caption         =   "frmAssignFillToLot.frx":01D6
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmAssignFillToLot.frx":021A
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAssignFillToLot.frx":023A
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
   Begin HexUniControls.ctlUniLabelXP lblQuantity 
      Height          =   255
      Left            =   180
      Top             =   450
      Width           =   1395
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
      Caption         =   "frmAssignFillToLot.frx":0256
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      Alignment       =   0
      VAlignment      =   0
      BackStyle       =   1
      BorderStyle     =   0
      AutoSize        =   0   'False
      Tip             =   "frmAssignFillToLot.frx":029E
      Style           =   0
      Enabled         =   -1  'True
      Margin          =   0
      RoundedBorders  =   0   'False
      MousePointer    =   0
      MouseIcon       =   "frmAssignFillToLot.frx":02BE
      RightToLeft     =   0   'False
      WordWrap        =   0   'False
   End
End
Attribute VB_Name = "frmAssignFillToLot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        frmAssignFillToLot.frm
'' Description: Form for allowing user to assign some portion of a fill to a lot
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 09/14/2012   DAJ         Added resize code to override minimum size after PlaceForm
'' 10/22/2012   DAJ         Rename Turnkey to HedgeLinc
'' 11/15/2013   DAJ         Changed the way to get Turnkey icon for the form
'' 02/18/2014   DAJ         Error if combo box of lots is empty
'' 03/07/2014   DAJ         Moved into NavCattle.DLL
'' 05/22/2014   DAJ         Renamed frmTurnkeyFillAssign to frmAssignFillToLot
'' 05/22/2014   DAJ         Renamed g.Turnkey to g.Cattle
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    bOK As Boolean                      ' Did the user click OK?

    Quantity As cPriceEditor            ' Quantity control
    lAssignedQuantity As Long           ' Quantity user chose to assign
    strFeedYardLotID As String          ' Lot user chose to assign to
End Type
Private m As mPrivate

Public Property Get AssignedQuantity() As Long
    AssignedQuantity = m.lAssignedQuantity
End Property

Public Property Get FeedYardLotID() As String
    FeedYardLotID = m.strFeedYardLotID
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMe
'' Description: Setup and show the form
'' Inputs:      Fill, Associated Fill, Associated Fills, Feed Yard Lot ID
'' Returns:     True if OK, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ShowMe(ByVal Fill As cBrokerMessage, ByVal AssociatedFill As cBrokerMessage, ByVal AssociatedFills As cGdTree, Optional ByVal strFeedYardLotID As String = "") As Boolean
On Error GoTo ErrSection:

    Dim lFillQuantity As Long           ' Quantity of the fill
    Dim lAssociatedQuantity As Long     ' Quantity of the fill that is already associated
    Dim lMaxAvailableQuantity As Long   ' Maximum available quantity to assign
    Dim lQuantity As Long               ' Quantity to set the control
    
    lFillQuantity = CLng(Val(Fill("Quantity")))
    lAssociatedQuantity = AssociatedQuantity(AssociatedFills)
    LoadLotsCombo AssociatedFill, AssociatedFills

    If cboLots.ListCount > 0 Then
        If AssociatedFill Is Nothing Then
            If Len(strFeedYardLotID) > 0 Then
                SelectComboByItemData cboLots, CLng(Val(strFeedYardLotID))
            Else
                cboLots.ListIndex = 0
            End If
            lMaxAvailableQuantity = lFillQuantity - lAssociatedQuantity
            lQuantity = 1&
        Else
            SelectComboByItemData cboLots, CLng(Val(AssociatedFill("FeedYardLotID")))
            lQuantity = CLng(Val(AssociatedFill("AssociatedQuantity")))
            lMaxAvailableQuantity = lFillQuantity - lAssociatedQuantity + lQuantity
        End If
        
        Set m.Quantity = New cPriceEditor
        m.Quantity.Init sbQuantity, txtQuantity, Nothing, lQuantity, 1, lMaxAvailableQuantity, , , 1
    
        ShowForm Me, eForm_Modal, g.frmMain
        
        If m.bOK Then
            m.lAssignedQuantity = m.Quantity.Price
            m.strFeedYardLotID = Str(cboLots.ItemData(cboLots.ListIndex))
        End If
    Else
        m.bOK = False
        InfBox "This fill is already associated with all of the lots.  If you wish to assign more of the quantity to a particular lot, modify the association.", "!", , "Error"
    End If
    
    ShowMe = m.bOK

ErrExit:
    Unload Me
    Exit Function
    
ErrSection:
    Unload Me
    RaiseError "frmAssignFillToLot.ShowMe"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdCancel_Click
'' Description: Handle the user clicking on the Cancel button
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
    RaiseError "frmAssignFillToLot.cmdCancel_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    cmdOK_Click
'' Description: Handle the user clicking on the OK button
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
    RaiseError "frmAssignFillToLot.cmdOK_Click"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Load
'' Description: Setup the form when it is loaded
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
On Error GoTo ErrSection:

    Icon = g.AppBridge.Picture16(g.Cattle.IconName)
    
    g.Styler.StyleForm Me
    
    Caption = "Fill Association"
    PlaceForm Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAssignFillToLot.Form_Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_QueryUnload
'' Description: Handle the user clicking on the X
'' Inputs:      Cancel the Unload?, Mode of the Unload
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
    RaiseError "frmAssignFillToLot.Form_QueryUnload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Resize
'' Description: Move and resize controls as the form is resized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Resize()
On Error Resume Next

    Dim lMinScaleWidth As Long          ' Minimum scale width
    Dim lMinScaleHeight As Long         ' Minimum scale height
    
    lMinScaleWidth = 7005
    lMinScaleHeight = 1725
    
    If Not LimitFormSize(Me, lMinScaleWidth, lMinScaleHeight) Then
        With fraButtons
            .Move (ScaleWidth - .Width) / 2
        End With
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Form_Unload
'' Description: Clean up when the form is unloaded
'' Inputs:      Cancel the Unload?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrSection:

    SaveFormPlacement Me

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAssignFillToLot.Form_Unload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadLotsCombo
'' Description: Load the Lots combo
'' Inputs:      Associated fill, Associated fills
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadLotsCombo(ByVal AssociatedFill As cBrokerMessage, ByVal AssociatedFills As cGdTree)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim Lot As cBrokerMessage           ' Lot object
    Dim Fill As cBrokerMessage          ' Fill object
    Dim astrLotsUsed As cGdArray        ' Array of lots already associated to the fill
    Dim lPos As Long                    ' Position in the array
    Dim strFeedYardLotID As String      ' Feed Yard lot ID for the associated fill
    
    If AssociatedFill Is Nothing Then
        strFeedYardLotID = ""
    Else
        strFeedYardLotID = AssociatedFill("FeedYardLotID")
    End If
    
    Set astrLotsUsed = New cGdArray
    If Not AssociatedFills Is Nothing Then
        For lIndex = 1 To AssociatedFills.Count
            Set Fill = AssociatedFills(lIndex)
            
            If Fill("FeedYardLotID") <> strFeedYardLotID Then
                If astrLotsUsed.BinarySearch(Fill("FeedYardLotID"), lPos) = False Then
                    astrLotsUsed.Add Fill("FeedYardLotID"), lPos
                End If
            End If
        Next lIndex
    End If
    
    cboLots.Clear
    For lIndex = 1 To g.Cattle.Lots.Count
        Set Lot = g.Cattle.Lots(lIndex)
        
        If astrLotsUsed.BinarySearch(Lot("FeedYardLotID")) = False Then
            cboLots.AddItem g.Cattle.LotDisplay(Lot)
            cboLots.ItemData(cboLots.NewIndex) = CLng(Val(Lot("FeedYardLotID")))
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "frmAssignFillToLot.LoadLotsCombo"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AssociatedQuantity
'' Description: Calculate the total quantity of the given associated fills
'' Inputs:      Associated Fills
'' Returns:     Associated Quantity
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function AssociatedQuantity(ByVal AssociatedFills As cGdTree) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    Dim AssociatedFill As cBrokerMessage ' Associated fill
    
    lReturn = 0&
    If Not AssociatedFills Is Nothing Then
        For lIndex = 1 To AssociatedFills.Count
            Set AssociatedFill = AssociatedFills(lIndex)
            lReturn = lReturn + CLng(Val(AssociatedFill("AssociatedQuantity")))
        Next lIndex
    End If
    
    AssociatedQuantity = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "frmAssignFillToLot.AssociatedQuantity"
    
End Function

