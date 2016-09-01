VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTradeSenseOrderGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cTradeSenseOrderGroup.cls
'' Description: Class that handles a Trade Sense order group
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History
'' Date         Author      Description
'' 07/20/2010   DAJ         Added capabilities for inputs
'' 07/21/2010   DAJ         Changed over to the Provided.INI for Daniel Code stuff
'' 08/05/2010   DAJ         Chnaged the Daniel Code load
'' 08/12/2010   DAJ         Added the ContainsOrder function
'' 08/23/2010   DAJ         Added required module flag for TradeSense orders/groups
'' 09/29/2010   DAJ         Added reload function, clear orders when load
'' 09/30/2010   DAJ         Added check for only exit orders
'' 11/16/2010   DAJ         Added allow manual submission flag
'' 12/06/2010   DAJ         Use default value for non-DC inputs
'' 03/17/2011   DAJ         Added flatten opposite property
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Const kTsoGroupVersion = 1
Private Const kExtension = "TSG"

Private Type mPrivate
    lVersion As Long                    ' File Version
    strID As String                     ' Unique identifier used as the filename
    strName As String                   ' Name of the order group
    strDescription As String            ' Description
    Orders As cGdTree                   ' Collection of Trade Sense orders
    bCustom As Boolean                  ' Is this user generated?
    bLinkInputs As Boolean              ' Link inputs with the same name?
    Inputs As cTradeSenseOrderInputs    ' Collection of expression inputs
    strRequiredMod As String            ' Required module
    bAllowManualSubmission As Boolean   ' Allow manual submission of the order group?
    bFlattenOpposite As Boolean         ' When this group enters a position, flatten others in opposite direction?
End Type
Private m As mPrivate

Public Property Get ID() As String
    ID = m.strID
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

Public Property Get Orders() As cGdTree
    Set Orders = m.Orders
End Property

Public Property Get Custom() As Boolean
    Custom = m.bCustom
End Property

Public Property Get Order(ByVal vIndexOrKey As Variant) As cTradeSenseOrder
On Error GoTo ErrSection:

    Dim tsOrder As cTradeSenseOrder     ' Trade Sense order group object
    
    If OrderExists(vIndexOrKey) Then
        Set tsOrder = m.Orders.Item(vIndexOrKey)
    Else
        Set tsOrder = Nothing
    End If
    
    Set Order = tsOrder

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cTradeSenseOrderGroup.Order.Get"
    
End Property
Public Property Let Order(ByVal vIndexOrKey As Variant, tsOrder As cTradeSenseOrder)
On Error GoTo ErrSection:
    
    If OrderExists(vIndexOrKey) Then
        Set m.Orders(vIndexOrKey) = tsOrder
    Else
        AddOrder tsOrder
    End If

ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cTradeSenseOrderGroup.Order.Let"
    
End Property

Public Property Get LinkInputs() As Boolean
    LinkInputs = m.bLinkInputs
End Property
Public Property Let LinkInputs(ByVal bLinkInputs As Boolean)
    m.bLinkInputs = bLinkInputs
End Property

Public Property Get Inputs() As cTradeSenseOrderInputs
    Set Inputs = m.Inputs
End Property
Public Property Let Inputs(ByVal tsInputs As cTradeSenseOrderInputs)
    Set m.Inputs = tsInputs
End Property

Public Property Get RequiredMod() As String
    RequiredMod = m.strRequiredMod
End Property
Public Property Let RequiredMod(ByVal strRequiredMod As String)
    m.strRequiredMod = strRequiredMod
End Property

Public Property Get AllowManualSubmission() As Boolean
    AllowManualSubmission = m.bAllowManualSubmission
End Property
Public Property Let AllowManualSubmission(ByVal bAllowManualSubmission As Boolean)
    m.bAllowManualSubmission = bAllowManualSubmission
End Property

Public Property Get FlattenOpposite() As Boolean
    FlattenOpposite = m.bFlattenOpposite
End Property
Public Property Let FlattenOpposite(ByVal bFlattenOpposite As Boolean)
    m.bFlattenOpposite = bFlattenOpposite
End Property

Public Function OrderCount() As Long
    OrderCount = m.Orders.Count
End Function

Public Function OrderExists(ByVal vIndexOrKey As Variant)
    OrderExists = m.Orders.Exists(vIndexOrKey)
End Function

Public Function AddOrder(ByVal tsOrder As cTradeSenseOrder)
    AddOrder = m.Orders.Add(tsOrder.MakeCopy, Str(tsOrder.OrderNumber))
End Function

Public Function RemoveOrder(ByVal vIndexOrKey As Variant) As Boolean
    RemoveOrder = m.Orders.Remove(vIndexOrKey)
End Function

Public Sub ClearOrders()
    m.Orders.Clear
End Sub

Public Sub ClearID()
    m.strID = ""
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    FromFile
'' Description: Load up the object from the given file
'' Inputs:      Filename, Custom?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub FromFile(ByVal strFileName As String, ByVal bCustom As Boolean)
On Error GoTo ErrSection:

    Dim astrFile As New cGdArray        ' Array to read the file into
    Dim lIndex As Long                  ' Index into a for loop
    Dim lPos As Long                    ' Position of the equal sign in a line
    Dim strKey As String                ' Key for the value
    Dim strValue As String              ' Value for the key
    Dim tsOrder As cTradeSenseOrder     ' Trade sense order object
    
    If astrFile.FromFile(strFileName) Then
        m.strID = FileBase(strFileName)
        m.bCustom = bCustom
        m.Orders.Clear
        
        For lIndex = 0 To astrFile.Size - 1
            lPos = InStr(astrFile(lIndex), "=")
            If lPos > 0 Then
                strKey = Left(astrFile(lIndex), lPos - 1)
                strValue = Mid(astrFile(lIndex), lPos + 1)
                
                Select Case UCase(strKey)
                    Case "VERSION"
                        m.lVersion = CLng(Val(strValue))
                    Case "NAME"
                        m.strName = strValue
                    Case "DESCRIPTION"
                        m.strDescription = strValue
                    Case "ORDER"
                        Set tsOrder = New cTradeSenseOrder
                        tsOrder.FromGroupString strValue
                        AddOrder tsOrder
                    Case "LINKINPUTS"
                        m.bLinkInputs = (Len(strValue) > 0) And (strValue <> "0")
                    Case "INPUTS"
                        m.Inputs.FromString strValue
                    Case "REQUIRED"
                        m.strRequiredMod = strValue
                    Case "ALLOWMANUAL"
                        m.bAllowManualSubmission = (Len(strValue) = 0) Or (strValue <> "0")
                    Case "FLATTENOPPOSITE"
                        m.bFlattenOpposite = (Len(strValue) = 0) Or (strValue <> "0")
                End Select
            End If
        Next lIndex
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeSenseOrderGroup.FromFile"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Reload
'' Description: Reload the order group from the file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Reload()
On Error GoTo ErrSection:

    If Len(m.strID) > 0 Then
        FromFile FileName, m.bCustom
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeSenseOrderGroup.Reload"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ToFile
'' Description: Save the object to a file
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ToFile()
On Error GoTo ErrSection:

    Dim astrFile As New cGdArray        ' Array of information to save to file
    Dim lIndex As Long                  ' Index into a for loop
    Dim tsOrder As cTradeSenseOrder     ' Trade Sense order object
    
    If Len(m.strID) = 0 Then
        m.strID = FileBase(GetUniqueCustomFilename(kExtension))
    End If
    
    astrFile.Add "Version=" & Str(m.lVersion)
    astrFile.Add "Name=" & m.strName
    astrFile.Add "Description=" & m.strDescription
    
    For lIndex = 1 To m.Orders.Count
        Set tsOrder = Order(lIndex)
        If Not tsOrder Is Nothing Then
            astrFile.Add "Order=" & tsOrder.ToGroupString
        End If
    Next lIndex
    
    astrFile.Add "LinkInputs=" & Str(m.bLinkInputs)
    astrFile.Add "Inputs=" & m.Inputs.ToString
    astrFile.Add "Required=" & m.strRequiredMod
    astrFile.Add "AllowManual=" & Str(m.bAllowManualSubmission)
    astrFile.Add "FlattenOpposite=" & Str(m.bFlattenOpposite)
    
    astrFile.ToFile FileName

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeSenseOrderGroup.ToFile"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Filename
'' Description: Determine the filename and path given the name and custom flag
'' Inputs:      None
'' Returns:     Filename and Path
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function FileName() As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function

    If m.bCustom Then
        strReturn = AddSlash(App.Path) & "Custom\" & m.strID & "." & kExtension
    Else
        strReturn = AddSlash(App.Path) & "Provided\" & m.strID & "." & kExtension
    End If
    
    FileName = strReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeSenseOrderGroup.Filename"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetInputValues
'' Description: Set input values from the given inputs collection
'' Inputs:      Inputs
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetInputValues(ByVal Inputs As cTradeSenseOrderInputs)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim strKey As String                ' Key into the collection
    
    If Not Inputs Is Nothing Then
        For lIndex = 1 To Inputs.Count
            strKey = Inputs(lIndex).Key(True)
            SetInputValue strKey, Inputs(lIndex).Value
        Next lIndex
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeSenseOrderGroup.SetInputValues"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadDanielCodeGroup
'' Description: Load up the appropriate Daniel Code group
'' Inputs:      Group Information, Entry Price, Stop Loss, Profit Target
'' Returns:     True if loaded, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LoadDanielCodeGroup(ByVal strGroupInfo As String, ByVal strEntryPrice As String, ByVal strStopLoss As String, ByVal strProfitTarget As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim astrGroupInfo As New cGdArray   ' Group information
    Dim lIndex As Long                  ' Index into a for loop
    
    If Len(strGroupInfo) > 0 Then
        astrGroupInfo.SplitFields strGroupInfo, vbTab
        
        If astrGroupInfo.Size > 0 Then
            FromFile AddSlash(App.Path) & astrGroupInfo(0), False
            
            For lIndex = 1 To m.Inputs.Count
                If UCase(m.Inputs(lIndex).Name) = UCase(astrGroupInfo(1)) Then
                    SetInputValue lIndex, strEntryPrice
                ElseIf UCase(m.Inputs(lIndex).Name) = UCase(astrGroupInfo(2)) Then
                    SetInputValue lIndex, strStopLoss
                ElseIf UCase(m.Inputs(lIndex).Name) = UCase(astrGroupInfo(3)) Then
                    SetInputValue lIndex, strProfitTarget
                Else
                    SetInputValue lIndex, m.Inputs(lIndex).DefaultValue
                End If
            Next lIndex
            
            bReturn = True
        End If
    End If
    
    LoadDanielCodeGroup = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeSenseOrderGroup.LoadDanielCodeGroup"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ContainsOrder
'' Description: Does this order group contain the given order?
'' Inputs:      Order ID
'' Returns:     True if group contains order, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ContainsOrder(ByVal strOrderID As String) As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    bReturn = False
    For lIndex = 1 To OrderCount
        If Order(lIndex).ID = strOrderID Then
            bReturn = True
            Exit For
        End If
    Next lIndex
    
    ContainsOrder = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeSenseOrderGroup.ContainsOrder"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    OnlyContainsExits
'' Description: Does this group only contain exit orders?
'' Inputs:      None
'' Returns:     True if only exits, False otherwise
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OnlyContainsExits() As Boolean
On Error GoTo ErrSection:

    Dim bReturn As Boolean              ' Return value for the function
    Dim lIndex As Long                  ' Index into a for loop
    
    If m.Orders.Count = 0 Then
        bReturn = False
    Else
        bReturn = True
        For lIndex = 1 To m.Orders.Count
            If Order(lIndex).Entry = True Then
                bReturn = False
                Exit For
            End If
        Next lIndex
    End If
    
    OnlyContainsExits = bReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cTradeSenseOrderGroup.OnlyContainsExits"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SetInputValue
'' Description: Set the value for the given input name
'' Inputs:      Name, Value
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub SetInputValue(ByVal vIndexOrKey As Variant, ByVal strValue As String)
On Error GoTo ErrSection:

    If m.Inputs.Exists(vIndexOrKey) Then
        m.Inputs(vIndexOrKey).Value = strValue
        
        If OrderExists(Str(m.Inputs(vIndexOrKey).OrderNumber)) Then
            Order(Str(m.Inputs(vIndexOrKey).OrderNumber)).SetInputValue m.Inputs(vIndexOrKey)
        End If
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeSenseOrderGroup.SetInputValue"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize class members when the class is intialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.Orders = New cGdTree
    m.bCustom = True
    m.lVersion = kTsoGroupVersion
    
    Set m.Inputs = New cTradeSenseOrderInputs
    m.Inputs.ForGroups = True
    m.strRequiredMod = ""
    m.bAllowManualSubmission = True
    m.bFlattenOpposite = True

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeSenseOrderGroup.Class_Initialize"

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up class members when the class is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.Orders = Nothing
    Set m.Inputs = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTradeSenseOrderGroup.Class_Terminate"

End Sub

