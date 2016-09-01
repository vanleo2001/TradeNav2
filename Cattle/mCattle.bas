Attribute VB_Name = "mCattle"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        mCattle.bas
'' Description: Main module for the Cattle DLL
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 03/07/2014   DAJ         Created
'' 03/14/2014   DAJ         Added formatted DLL version function
'' 04/08/2014   DAJ         Copied the Grid Scroll fix into DLL from NavSuite project to fix error
'' 05/22/2014   DAJ         Renamed cTurnkey to cCattle; Renamed g.Turnkey to g.Cattle
'' 05/30/2014   DAJ         Utilized new accounts object
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Public Const kTurnkeyCompanyName As String = "HedgeLinc"

'RH
Public Enum eStyleColorTypes
    'form
    eForm_Background = 0
    
    'frame
    eFrame_Background = 1
    eFrame_Border = 2
    
    'button
    eButton_Background = 3
    eButton_Border = 4
    eButton_Text = 5
    
    'checkbox
    eCheck_Border = 6
    eCheck_Background = 7
    eCheck_Forecolor = 8
    
    'flexgrid
    eGrid_Background = 9
    
End Enum


Type gGlobal
    AppBridge As cCattleTn              ' Application bridge from Cattle DLL to Trade Navigator
    strAppPath As String                ' Main application path
    strIniFile As String                ' Main application INI file
    lDataServiceID As Long              ' Customers data service ID
    strMachineID As String              ' Customers Machine ID
    strPassword As String               ' Customers Password
    frmMain As Form                     ' Main form
    lStreamInterval As Long             ' Interval for the stream timer
    bStreamActive As Boolean            ' Is the Trade Navigator stream active?
    Help As cHelp                       ' Help object
    
    Cattle As cCattle                   ' Global Cattle object
    BrokerEnums As cBrokerEnums         ' Broker enumerations object
    CattleEnums As cCattleEnums         ' Cattle enumerations object
    CattleKeyValue As cCattleKeyValue   ' Object to go back and forth with key value objects
    
    'RH
    Styler As New cStyler
    
End Type
Global g As gGlobal

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    PlaceForm
'' Description: Place the given form appropriately
'' Inputs:      Form to Place
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PlaceForm(FormToPlace As Form)
On Error GoTo ErrSection:

    mGenesis.PlaceTheForm FormToPlace, g.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mCattle.PlaceForm"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    SaveFormPlacement
'' Description: Save the placement of the given form
'' Inputs:      Form to Save
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SaveFormPlacement(FormToSave As Form)
On Error GoTo ErrSection:

    mGenesis.SaveTheFormPlacement FormToSave, g.strIniFile

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "mCattle.SaveFormPlacement"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Picture16
'' Description: Grab the appropriate image from the appropriate image list
'' Inputs:      Picture name, Image List
'' Returns:     Image
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Picture16(ByVal strPicture$, Optional ByVal iImageList As Integer = 0) As Object
    Set Picture16 = g.AppBridge.Picture16(strPicture, iImageList)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ShowMsg
'' Description: Show an error message
'' Inputs:      Form to Save
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub ShowMsg(Optional ByVal lErrNum& = 0, Optional ByVal strSource$ = "", Optional ByVal strDesc$ = "")
    
    Dim RetVal As Variant
    
    If lErrNum = 0 Then
        lErrNum = Err.Number
        strSource = Err.Source
        strDesc = Err.Description
    End If
    
    RetVal = LockWindowUpdate(0)
    Screen.MousePointer = vbDefault
    
    If lErrNum < 0 Then
        Replace strDesc, vbCrLf, "|"
        InfBox strDesc, , , "Error", , , , , , , , eGDAlign_Left
    Else
        Replace strDesc, vbCrLf, "|"
        InfBox "An unexpected error occurred.||Please report the following: " & _
            "|Source:  " & strSource & _
            "|Message: " & strDesc, , , "Error", , , , , , , , eGDAlign_Left
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddBrokerAccountNumberToOrders
'' Description: Walk through given collection of orders and add broker account number
'' Inputs:      Orders
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddBrokerAccountNumberToOrders(Orders As cGdTree)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To Orders.Count
        AddBrokerAccountNumberToOrder Orders(lIndex)
    Next lIndex

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mCattle.AddBrokerAccountNumberToOrders"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddBrokerAccountNumberToOrder
'' Description: Add broker account number to the given order
'' Inputs:      Order
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddBrokerAccountNumberToOrder(Order As cBrokerMessage)
On Error GoTo ErrSection:

    Order.Add "BrokerAccountID", g.Cattle.Accounts.AccountIdForBrokerNumber(Order("BrokerAccountNumber"), CLng(Val(Order("Broker"))))

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mCattle.AddBrokerAccountNumberToOrder"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddBrokerAccountNumberToFills
'' Description: Walk through given collection of fills and add broker account number
'' Inputs:      Fills
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddBrokerAccountNumberToFills(Fills As cGdTree)
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To Fills.Count
        AddBrokerAccountNumberToFill Fills(lIndex)
    Next lIndex

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mCattle.AddBrokerAccountNumberToFills"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddBrokerAccountNumberToFill
'' Description: Add broker account number to the given fill
'' Inputs:      Fill
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub AddBrokerAccountNumberToFill(Fill As cBrokerMessage)
On Error GoTo ErrSection:

    Fill.Add "BrokerAccountID", g.Cattle.Accounts.AccountIdForBrokerNumber(Fill("BrokerAccountNumber"), CLng(Val(Fill("Broker"))))

ErrExit:
    Exit Sub

ErrSection:
    RaiseError "mCattle.AddBrokerAccountNumberToFill"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    CopyTreeFromObject
'' Description: Build a cGdTree copy from an object passed from the app
'' Inputs:      Tree to Copy
'' Returns:     Copied tree
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CopyTreeFromObject(TreeToCopy As Object) As cGdTree
On Error GoTo ErrSection:

    Dim ReturnTree As cGdTree           ' Tree to return from the function
    Dim lIndex As Long                  ' Index into a for loop
    
    If Not TreeToCopy Is Nothing Then
        Set ReturnTree = New cGdTree
        For lIndex = 1 To TreeToCopy.Count
            ReturnTree.Add TreeToCopy(lIndex), TreeToCopy.Key(lIndex)
        Next lIndex
    End If
    
    Set CopyTreeFromObject = ReturnTree

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mCattle.CopyTreeFromObject"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DllVersion
'' Description: Format the version of the DLL
'' Inputs:      Include Revision?, Include File Date?
'' Returns:     Formatted version
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DllVersion(Optional ByVal bIncludeRevision As Boolean = False, Optional ByVal bIncludeFileDate As Boolean = False) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim strFileName As String           ' Filename for the DLL

    strReturn = Format(App.Major, "#0") & "." & Format(App.Minor, "#0")
    
    If bIncludeRevision Then
        strReturn = strReturn & "." & Str(App.Revision)
    End If
    
    If bIncludeFileDate Then
        strFileName = AddSlash(g.strAppPath) & "..\SharedSelfReg\" & App.EXEName & ".DLL"
        strReturn = strReturn & " " & DateFormat(FileDate(strFileName), MM_DD_YYYY, HH_MM, AMPM_UPPER)
    End If
    
    DllVersion = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mCattle.DllVersion"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GridScrollCheck
'' Description: Help fix the inadvertant scrolling issues with the FlexGrids while
''              streaming is on ( call from the grid's BeforeScroll event )
'' Inputs:      Grid, Old Top Row, Old Left Column, New Top Row, New Left Column, Cancel?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GridScrollCheck(fg As VSFlexGrid, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)
On Error Resume Next
    
    Dim pt As POINTAPI
    Static iVertHorizMode As Integer ' mode: 0=None, 1=Vert, 2=Horiz
       
    ' If mouse is not pressed, then just clear the mode
    If fg Is Nothing Or Not MouseIsPressed Then
        iVertHorizMode = 0
    ElseIf iVertHorizMode = 0 Then
        ' But if mouse is pressed and the mode has not yet been set, then user must be
        ' pressing a grid scrollbar -- so need to find out which one (vert or horiz).
        ' First get the X,Y coords of the mouse:
        If GetCursorPos(pt) <> 0 Then
            ' then convert to coords based on the grid itself
            If ScreenToClient(fg.hWnd, pt) <> 0 Then
                ' then see if the mouse is either below the grid (horiz scrollbar)
                ' or is to the right of the grid (vert scrollbar)
                If pt.X >= fg.ClientWidth / Screen.TwipsPerPixelX Then
                    iVertHorizMode = 1 ' Vert scrollbar is right of grid
                ElseIf pt.Y >= fg.ClientHeight / Screen.TwipsPerPixelY Then
                    iVertHorizMode = 2 ' Horiz scrollbar is below grid
                End If
            End If
        End If
        ' and start the timer which will clear this stuff as soon as the mouse is no longer pressed
        If iVertHorizMode <> 0 Then
            frmCattleAM.tmrGridScrollPressed.Enabled = True
        End If
    End If
        
    If iVertHorizMode = 1 Then
        ' while the vertical scrollbar is pressed, we will NOT be scrolling columns
        If OldLeftCol <> NewLeftCol Then
            Cancel = True
        End If
    ElseIf iVertHorizMode = 2 Then
        ' while the horiz scrollbar is pressed, we will NOT be scrolling rows
        If OldTopRow <> NewTopRow Then
            Cancel = True
        End If
    End If

End Sub
