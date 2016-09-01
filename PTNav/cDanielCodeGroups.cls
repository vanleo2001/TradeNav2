VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDanielCodeGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        cDanielCodeGroups.cls
'' Description: Class that handles Daniel Code groups
''
'' Author:      Genesis Financial Technologies
''              4775 Centennial Blvd Ste 150
''              Colorado Springs, CO  80919
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author      Description
'' 10/27/2010   DAJ         Implemented default group for type
'' 11/15/2010   DAJ         Don't include groups if user doesn't have required module
'' 11/16/2010   DAJ         Added group map to look up defaults
'' 08/21/2012   DAJ         Fix some of the RaiseError strings
'' 08/31/2012   DAJ         Load different INI file properties for GmajPro
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type mPrivate
    bGmajPro As Boolean                 ' Gmaj Pro version?
    strIniFileSection As String         ' Ini File Section
    
    Groups As cGdTree                   ' Collection of arrays of group information
    GroupMap As cGdTree                 ' Map of indexes to groups
    Defaults As cGdTree                 ' Collection of default indexes
End Type
Private m As mPrivate

Private Property Get IniFile() As String
    IniFile = AddSlash(App.Path) & "Provided\Provided.INI"
End Property

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Load
'' Description: Load the Daniel Code group information
'' Inputs:      GmajPro?
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Load(ByVal bGmajPro As Boolean)
On Error GoTo ErrSection:

    m.bGmajPro = bGmajPro
    If m.bGmajPro Then
        m.strIniFileSection = "GmajPro"
    Else
        m.strIniFileSection = "DanielCode"
    End If

    LoadType "LongFut"
    LoadType "ShortFut"
    LoadType "LongFx"
    LoadType "ShortFx"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cDanielCodeGroups.Load"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetGroups
'' Description: Get the groups for a given type
'' Inputs:      Direction, Symbol
'' Returns:     Array of Groups
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetGroups(ByVal strDirection As String, ByVal strSymbol As String) As cGdArray
On Error GoTo ErrSection:

    Dim strType As String               ' Type for the given direction and symbol
    Dim astrGroups As cGdArray          ' Array to return
    
    strType = GetType(strDirection, strSymbol)
    If Len(strType) = 0 Then
        Set astrGroups = New cGdArray
        astrGroups.Create eGDARRAY_Strings
    Else
        Set astrGroups = m.Groups(strType)
    End If
    
    Set GetGroups = astrGroups

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cDanielCodeGroups.GetGroups"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetDefault
'' Description: Get the default index for a given type
'' Inputs:      Direction, Symbol, Type
'' Returns:     Index into the Array
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetDefault(ByVal strDirection As String, ByVal strSymbol As String, ByVal strType As String) As Long
On Error GoTo ErrSection:

    Dim lReturn As Long                 ' Return value for the function
    Dim strGroupType As String          ' Type for the given direction and symbol
    Dim strGroup As String              ' Group key
    
    lReturn = 0
    If (Len(strDirection) > 0) And (Len(strSymbol) > 0) And (Len(strType) > 0) Then
        strGroupType = GetType(strDirection, strSymbol)
        If Len(strGroupType) > 0 Then
            If m.Defaults.Exists(strGroupType & strType) Then
                strGroup = strGroupType & Str(m.Defaults(strGroupType & strType))
                If m.GroupMap.Exists(strGroup) Then
                    lReturn = m.GroupMap(strGroup)
                End If
            End If
        End If
    End If
    
    GetDefault = lReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cDanielCodeGroups.GetDefault"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetGroupInfo
'' Description: Get the group information for a given type, group name
'' Inputs:      Direction, Symbol, Group Name
'' Returns:     Group Information
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetGroupInfo(ByVal strDirection As String, ByVal strSymbol As String, ByVal strGroupName As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim astrGroups As cGdArray          ' Groups for the given direction and symbol
    Dim lIndex As Long                  ' Index into a for loop
    Dim astrGroupInfo As New cGdArray   ' Group information
    
    strReturn = ""
    Set astrGroups = GetGroups(strDirection, strSymbol)
    If astrGroups.Size > 0 Then
        For lIndex = 0 To astrGroups.Size - 1
            If Parse(astrGroups(lIndex), vbTab, 1) = strGroupName Then
                astrGroupInfo.SplitFields astrGroups(lIndex), vbTab
                astrGroupInfo.Remove 0
                
                strReturn = astrGroupInfo.JoinFields(vbTab)
                Exit For
            End If
        Next lIndex
    End If
    
    GetGroupInfo = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cDanielCodeGroups.GetGroupInfo"

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadType
'' Description: Load groups for the given type
'' Inputs:      Type
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadType(ByVal strType As String)
On Error GoTo ErrSection:

    Dim lNumGroups As Long              ' Number of groups
    Dim strGroupInfo As String          ' Group information
    Dim lIndex As Long                  ' Index into a for loop
    Dim astrGroups As cGdArray          ' Array of group information
    
    Set astrGroups = New cGdArray
    astrGroups.Create eGDARRAY_Strings
    
    lNumGroups = GetIniFileProperty("Num" & strType, 0&, m.strIniFileSection, IniFile)
    If lNumGroups = 0 Then
        strGroupInfo = LoadGroup(strType)
        If Len(strGroupInfo) > 0 Then
            astrGroups.Add strGroupInfo
        End If
    Else
        For lIndex = 1 To lNumGroups
            strGroupInfo = LoadGroup(strType & Str(lIndex))
            If Len(strGroupInfo) > 0 Then
                astrGroups.Add strGroupInfo
                m.GroupMap.Add astrGroups.Size - 1, strType & Str(lIndex)
            End If
        Next lIndex
    End If

    m.Groups.Add astrGroups, strType
    
    LoadDefaultsForType strType

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cDanielCodeGroups.LoadType"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadDefaultsForType
'' Description: Load defaults for the given type
'' Inputs:      Type
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub LoadDefaultsForType(ByVal strType As String)
On Error GoTo ErrSection:

    Dim astrSignalTypes As cGdArray     ' Signal types
    Dim lDefault As Long                ' Default value
    Dim lIndex As Long                  ' Index into a for loop
    Dim strPropName As String           ' Property name
    
    Set astrSignalTypes = New cGdArray
    astrSignalTypes.SplitFields GetIniFileProperty("SignalTypes", "", m.strIniFileSection, IniFile), ","
    
    For lIndex = 0 To astrSignalTypes.Size - 1
        strPropName = strType & "Def" & astrSignalTypes(lIndex)
        
        lDefault = GetIniFileProperty(strPropName, kNullData, m.strIniFileSection, IniFile)
        If lDefault <> kNullData Then
            m.Defaults.Add lDefault, strType & astrSignalTypes(lIndex)
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cDanielCodeGroups.LoadDefaultsForType"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    LoadGroup
'' Description: Load group information for the given property name
'' Inputs:      Property Name
'' Returns:     Group Information
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LoadGroup(ByVal strPropName As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim tsoGrp As cTradeSenseOrderGroup ' TradeSense order group object
    Dim strFileName As String           ' Filename for the order group
    Dim bCustom As Boolean              ' Is the group a custom group?
    
    strReturn = GetIniFileProperty(strPropName, "", m.strIniFileSection, IniFile)
    If Len(strReturn) > 0 Then
        strFileName = AddSlash(App.Path) & Parse(strReturn, vbTab, 1)
        bCustom = (UCase(Parse(strReturn, "\", 1)) = "CUSTOM")
    
        Set tsoGrp = New cTradeSenseOrderGroup
        tsoGrp.FromFile strFileName, bCustom
        
        If HasModule(tsoGrp.RequiredMod) Then
            strReturn = tsoGrp.Name & vbTab & strReturn
        Else
            strReturn = ""
        End If
    End If
    
    LoadGroup = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cDanielCodeGroups.LoadGroup"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    GetType
'' Description: Get the type given the direction and the symbol
'' Inputs:      Direction, Symbol
'' Returns:     Type
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetType(ByVal strDirection As String, ByVal strSymbol As String) As String
On Error GoTo ErrSection:

    Dim strReturn As String             ' Return value for the function
    Dim bBuy As Boolean                 ' Is this a Buy?
    
    bBuy = (UCase(strDirection) = "BUY")
    If IsForex(strSymbol) Then
        If bBuy Then
            strReturn = "LongFx"
        Else
            strReturn = "ShortFx"
        End If
    Else
        If bBuy Then
            strReturn = "LongFut"
        Else
            strReturn = "ShortFut"
        End If
    End If
    
    GetType = strReturn

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cDanielCodeGroups.GetType"
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Initialize
'' Description: Initialize class members when object is initialized
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
On Error GoTo ErrSection:

    Set m.Groups = New cGdTree
    Set m.GroupMap = New cGdTree
    Set m.Defaults = New cGdTree

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cDanielCodeGroups.Class_Initialize"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    Class_Terminate
'' Description: Clean up class members when object is terminated
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Terminate()
On Error GoTo ErrSection:

    Set m.Groups = Nothing
    Set m.GroupMap = Nothing
    Set m.Defaults = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cDanielCodeGroups.Class_Terminate"
    
End Sub

