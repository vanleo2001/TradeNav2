VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSettingsFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit
'==================================================================
'Summary
'-------
'This class manages data entry type grids.  Each row in the grid
'represents a data item.  Column 1 contains labels describing the data
'(i.e. System Name, System Type), Column 2 contains the data which can be
'changed by the user.  A preview box is shown below the data entry grid.
'As you move through data items in the grid, the detailed description
'of the current data item is shown here.
'
'Settings File
'-------------
'The data used in the grids come from a serialized gdTable called
'gdSettings.dat.  This file is built off of "tblSettings".  If a developer
'changes settings to a report or maintenance form, run the global
'module: "BuildSettingsFile" in order to update "gdSettings.dat".  This
'file must be distributed to users.

Private Type mPrivate
    SettingsFile As cGdTable
    lFirstRow As Long
    lLastRow As Long
    strPath As String
    strFileName As String
    bIsStrategy As Boolean
End Type
Private m As mPrivate

'Settings gdtable fields
Private Const mS_AppID = 0
Private Const mS_SettingID = 1
Private Const mS_SettingName = 2
Private Const mS_LabelName = 3
Private Const mS_Desc = 4
Private Const mS_Type = 5
Private Const mS_DecPos = 6
Private Const mS_Formatting = 7
Private Const mS_ValFrom = 8
Private Const mS_ValTo = 9
Private Const mS_Length = 10
Private Const mS_CanEdit = 11
Private Const mS_Required = 12
Private Const mS_ShowEdit = 13
Private Const mS_ShowAdd = 14
Private Const mS_Default = 15
Private Const mS_Group = 16
Private Const mS_Global = 17
Private Const mS_Save = 18
Private Const mS_DefaultKey = 19
Private Const mS_Value = 20

Public Enum ensgd_Settings
     ensgd_AppID = 0
     ensgd_SettingID = 1
     ensgd_SettingName = 2
     ensgd_LabelName = 3
     ensgd_Desc = 4
     ensgd_Type = 5
     ensgd_DecPos = 6
     ensgd_Formatting = 7
     ensgd_ValFrom = 8
     ensgd_ValTo = 9
     ensgd_Length = 10
     ensgd_CanEdit = 11
     ensgd_Required = 12
     ensgd_ShowEdit = 13
     ensgd_ShowAdd = 14
     ensgd_Default = 15
     ensgd_Group = 16
     ensgd_Global = 17
     ensgd_Save = 18
     ensgd_DefaultKey = 19
     ensgd_Value = 20
End Enum

Property Get TableHandle() As Long
    TableHandle = m.SettingsFile.TableHandle
End Property
Property Get NumRecords() As Long
    NumRecords = m.SettingsFile.NumRecords
End Property

Property Get Item(ByVal nName As ensgd_Settings, ByVal lRecord As Long) As Variant
    Item = m.SettingsFile.Item(nName, lRecord)
End Property
Property Let Item(ByVal nName As ensgd_Settings, ByVal lRecord As Long, ByVal vValue As Variant)
    m.SettingsFile.Item(nName, lRecord) = vValue
End Property

Property Get Num(ByVal nName As ensgd_Settings, ByVal lRecord As Long) As Variant
    Num = m.SettingsFile.Num(nName, lRecord)
End Property
Property Let Num(ByVal nName As ensgd_Settings, ByVal lRecord As Long, ByVal vValue As Variant)
    m.SettingsFile.Num(nName, lRecord) = vValue
End Property

Public Property Get Path() As String
    Path = m.strPath
End Property
Public Property Let Path(ByVal strPath As String)
    m.strPath = strPath
End Property

Public Property Get FileName() As String
    FileName = m.strFileName
End Property
Public Property Let FileName(ByVal Data As String)
    m.strFileName = Data
End Property

Public Property Get IsStrategy() As Boolean
    IsStrategy = m.bIsStrategy
End Property
Public Property Let IsStrategy(ByVal bIsStrategy As Boolean)
    m.bIsStrategy = bIsStrategy
End Property

Public Function ColNbr(ByVal pName As ensgd_Settings) As Long
On Error GoTo ErrSection:

    ColNbr = pName

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSettingsFile.ColNbr", eGDRaiseError_Raise, g.strAppPath

End Function

'Determine and return the From/To record position of the application
Public Sub GetFromTo(ByVal lAppID As Long, lTotalRows As Long, lFromRec As Long, lToRec As Long)
On Error GoTo ErrSection:
    
    Dim lIndex As Long
    
    lTotalRows = 0
    lFromRec = 0
    lToRec = 0
    
    For lIndex = 1 To m.SettingsFile.NumRecords - 1
        If lAppID = Item(ensgd_AppID, lIndex) Then
            If lFromRec = 0 Then lFromRec = lIndex
            lToRec = lIndex
        End If
    Next lIndex
    lTotalRows = lToRec - lFromRec + 1
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSettingsFile.GetFromTo", eGDRaiseError_Raise, g.strAppPath

End Sub

'This sets the value of all data items for the application.  If application
'type is "AllReports" then the same DataItemName is updated for all reports.
Public Sub SetItem(ByVal lAppID As Long, ByVal strDataItemName As String, _
    ByVal vValue As Variant, Optional ByVal Col As ensgd_Settings = ensgd_Default)
On Error GoTo ErrSection:
    
    Dim lIndex As Long                  ' Index into a for loop
    
    For lIndex = 1 To m.SettingsFile.NumRecords - 1
        If (lAppID = Item(ensgd_AppID, lIndex)) Or (lAppID = 0&) Then
            If Item(ensgd_SettingName, lIndex) = strDataItemName Then
                Select Case m.SettingsFile.Item(mS_Type, lIndex)
                    Case "Number", "Pct"
                        m.SettingsFile.Item(Col, lIndex) = Str(ValOfText(vValue))
                    Case "DateTime"
                        m.SettingsFile.Item(Col, lIndex) = Str(DateOf(vValue))
                    Case Else
                        m.SettingsFile.Item(Col, lIndex) = Str(vValue)
                End Select
                If lAppID <> 0& Then Exit For
            End If
        End If
    Next lIndex
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSettingsFile.SetItem", eGDRaiseError_Raise, g.strAppPath

End Sub

'This saves the value of a field from the grid back to gdSettings
Public Sub Save(ByVal strFieldName As String, ByVal vValue As Variant, ByVal vKeyValue As Variant, _
    ByVal lFirstRow As Long, ByVal lLastRow As Long)
On Error GoTo ErrSection:
    
    Dim lIndex As Integer               ' Index into a for loop
    
    For lIndex = lFirstRow To lLastRow
        If strFieldName = m.SettingsFile.Item(mS_SettingName, lIndex) Then
            Select Case m.SettingsFile.Item(mS_Type, lIndex)
                Case "Number", "Pct"
                    m.SettingsFile.Item(mS_Value, lIndex) = Str(ValOfText(vValue))
                    
                Case "DateTime"
                    If Len(vValue) > 0 Then
                        m.SettingsFile.Item(mS_Value, lIndex) = Str(CDate(vValue))
                    End If
                    
                Case "TableLookup"
                    m.SettingsFile.Item(mS_DefaultKey, lIndex) = vKeyValue
                    m.SettingsFile.Item(mS_Default, lIndex) = vValue
                                   
                Case Else
                    m.SettingsFile.Item(mS_Value, lIndex) = Str(vValue)
            End Select
            Exit For
        End If
    Next lIndex
                
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSettingsFile.Save", eGDRaiseError_Raise, g.strAppPath

End Sub

'Loads settings from gdSettings.dat.  If a new version exists then file
'gdSettings.gzp will exist in the root.  Backup the old and copy the new one
'in it's place.  Overlay the users current default setting values.  This
'should only be called one time at program startup since settings are changed
'by the user throughout the session.
#If 0 Then
Public Sub Load()
On Error GoTo ErrSection:
    
    Dim fs As New FileSystemObject
    Dim bNewFound As Boolean
    Dim gdsOld As cGdTable
    Dim lIndex As Long
    Dim lIndex2 As Long
    
    ' Look for a new version of setting files (gdSettings.gzp).
    If Dir(AddSlash(m.strPath) & FileBase(m.strFileName) & ".gzp", vbNormal) <> "" Then
        If Dir(AddSlash(m.strPath) & m.strFileName, vbNormal) <> "" Then
            fs.CopyFile AddSlash(m.strPath) & m.strFileName, AddSlash(m.strPath) & "gdSettings.old", True
        End If
        ZipExecute "U", AddSlash(m.strPath) & FileBase(m.strFileName) & ".gzp", m.strPath
        bNewFound = True
    End If
    
    ' Load Tables from disk
    If Dir(AddSlash(m.strPath) & m.strFileName, vbNormal) = "" Then
        Err.Raise gUserErr, , m.strFileName & " not found"
    End If
    m.SettingsFile.Serialize AddSlash(m.strPath) & m.strFileName, False
    
    ' If new settings found then Load settings file and copy over the default
    ' values of the old into the new.
    If bNewFound Then
        gdsOld.Serialize AddSlash(m.strPath) & FileBase(m.strFileName) & ".old", False
        For lIndex = 1 To gdsOld.NumRecords - 1
            For lIndex2 = 1 To m.SettingsFile.NumRecords - 1
                If gdsOld.Item(mS_SettingName, lIndex) = _
                   m.SettingsFile.Item(mS_SettingName, lIndex2) Then
                    m.SettingsFile.Item(mS_Default, lIndex2) = gdsOld.Item(mS_Default, lIndex)
                    m.SettingsFile.Item(mS_DefaultKey, lIndex2) = gdsOld.Item(mS_DefaultKey, lIndex)
                    Exit For
                End If
            Next lIndex2
        Next lIndex
        KillFile AddSlash(m.strPath) & m.strFileName
        KillFile AddSlash(m.strPath) & FileBase(m.strFileName) & ".gzp"
    End If
    
ErrExit:
    Set gdsOld = Nothing
    Set fs = Nothing
    Exit Sub
    
ErrSection:
    Set gdsOld = Nothing
    Set fs = Nothing
    RaiseError "cSettingsFile.Load", eGDRaiseError_Raise, g.strAppPath

End Sub
#Else
Public Sub Load()
On Error GoTo ErrSection:

    Dim lIndex As Long                  ' Index into a for loop
    Dim strIniFile As String            ' Ini file name and path
    Dim strValue As String              ' Value from the ini file

    ' Set the Ini File name and path...
    strIniFile = AddSlash(g.strAppPath) & "Reports.INI"

    ' Load Table from disk...
    If Dir(AddSlash(m.strPath) & m.strFileName, vbNormal) = "" Then
        Err.Raise vbObjectError + 1000, , m.strFileName & " not found"
    End If
    m.SettingsFile.Serialize AddSlash(m.strPath) & m.strFileName, False
    
    ' Add Value column to table...
    m.SettingsFile.CreateField eGDARRAY_Strings, mS_Value, "Value"
   
    ' Fill in Value column...
    For lIndex = 1 To m.SettingsFile.NumRecords - 1
        If m.SettingsFile.Num(mS_Global, lIndex) = 0 Then
            strValue = GetIniFileProperty(m.SettingsFile.Item(mS_SettingName, lIndex), "", Str(m.SettingsFile.Num(mS_AppID, lIndex)), strIniFile)
        Else
            strValue = GetIniFileProperty(m.SettingsFile.Item(mS_SettingName, lIndex), "", "Global", strIniFile)
        End If
        
        If m.SettingsFile.Item(mS_SettingName, lIndex) = "ChartBy" Then
            If m.bIsStrategy Then
                Item(ensgd_Default, lIndex) = "Date"
            Else
                Item(ensgd_Default, lIndex) = "Trade"
            End If
        End If
        
        If Len(strValue) = 0 Then
            m.SettingsFile.Item(mS_Value, lIndex) = m.SettingsFile.Item(mS_Default, lIndex)
        Else
            If m.SettingsFile.Item(mS_SettingName, lIndex) = "StatToChart" And strValue = "Net Profit" Then
                If InStr(m.SettingsFile.Item(mS_Formatting, lIndex), "Net Profit") = 0 Then
                    strValue = "Profit"
                End If
            End If
            m.SettingsFile.Item(mS_Value, lIndex) = strValue
        End If
    Next lIndex

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSettingsFile.Load", eGDRaiseError_Raise, g.strAppPath
    
End Sub
#End If

Private Sub Class_Initialize()
On Error GoTo ErrSection:
    
    Dim vRetVal As Variant
    
    Set m.SettingsFile = New cGdTable
    vRetVal = m.SettingsFile.CreateField(eGDARRAY_Longs, mS_AppID, "AppID")
    vRetVal = m.SettingsFile.CreateField(eGDARRAY_Longs, mS_SettingID, "SettingID")
    vRetVal = m.SettingsFile.CreateField(eGDARRAY_Strings, mS_SettingName, "SettingName")
    vRetVal = m.SettingsFile.CreateField(eGDARRAY_Strings, mS_LabelName, "LabelName")
    vRetVal = m.SettingsFile.CreateField(eGDARRAY_Strings, mS_Desc, "Desc")
    vRetVal = m.SettingsFile.CreateField(eGDARRAY_Strings, mS_Type, "Type")
    vRetVal = m.SettingsFile.CreateField(eGDARRAY_Longs, mS_DecPos, "DecPos")
    vRetVal = m.SettingsFile.CreateField(eGDARRAY_Strings, mS_Formatting, "Formatting")
    vRetVal = m.SettingsFile.CreateField(eGDARRAY_Doubles, mS_ValFrom, "ValFrom")
    vRetVal = m.SettingsFile.CreateField(eGDARRAY_Doubles, mS_ValTo, "ValTo")
    vRetVal = m.SettingsFile.CreateField(eGDARRAY_Longs, mS_Length, "Length")
    vRetVal = m.SettingsFile.CreateField(eGDARRAY_TinyInts, mS_CanEdit, "CanEdit")
    vRetVal = m.SettingsFile.CreateField(eGDARRAY_TinyInts, mS_Required, "Required")
    vRetVal = m.SettingsFile.CreateField(eGDARRAY_TinyInts, mS_ShowEdit, "ShowEdit")
    vRetVal = m.SettingsFile.CreateField(eGDARRAY_TinyInts, mS_ShowAdd, "ShowAdd")
    vRetVal = m.SettingsFile.CreateField(eGDARRAY_Strings, mS_Default, "Default")
    vRetVal = m.SettingsFile.CreateField(eGDARRAY_Strings, mS_DefaultKey, "DefaultKey")
    vRetVal = m.SettingsFile.CreateField(eGDARRAY_Strings, mS_Group, "Group")
    vRetVal = m.SettingsFile.CreateField(eGDARRAY_TinyInts, mS_Global, "Global")
    vRetVal = m.SettingsFile.CreateField(eGDARRAY_TinyInts, mS_Save, "Save")
    
    m.bIsStrategy = True
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSettingsFile.Initialize", eGDRaiseError_Raise, g.strAppPath

End Sub

Public Sub Serialize()
On Error GoTo ErrSection:
    
#If 0 Then
    Dim fs As New FileSystemObject
    
    ' Save current settings file settings
    fs.CopyFile AddSlash(m.strPath) & m.strFileName, AddSlash(m.strPath) & FileBase(m.strFileName) & ".bak"
    
    KillFile AddSlash(m.strPath) & m.strFileName
    m.SettingsFile.Serialize AddSlash(m.strPath) & m.strFileName, True

#Else

    Dim lIndex As Long                  ' Index into a for loop
    Dim strIniFile As String            ' Ini file name and path
    Dim strValue As String              ' Value from the ini file
    Dim bSkip As Boolean

    ' Set the Ini File name and path...
    strIniFile = AddSlash(g.strAppPath) & "Reports.INI"
    
    ' Save any overrides to the ini file...
    For lIndex = 1 To m.SettingsFile.NumRecords - 1
        With m.SettingsFile
            If .Num(mS_Save, lIndex) <> 0 Then
                bSkip = False
                If .Item(mS_Value, lIndex) = .Item(mS_Default, lIndex) Then
                    strValue = ""
                Else
                    strValue = .Item(mS_Value, lIndex)
                End If
                
                If .Item(mS_SettingName, lIndex) = "RiskBasedOn" Then
                    If ValOfText(strValue) <> 0 Then
                        bSkip = True ' don't save if RiskBasedOn is a number
                    End If
                End If
                
                If Not bSkip Then
                    If .Num(mS_Global, lIndex) = 0 Then
                        SetIniProp .Item(mS_SettingName, lIndex), strValue, Str(.Num(mS_AppID, lIndex)), strIniFile
                    Else
                        SetIniProp .Item(mS_SettingName, lIndex), strValue, "Global", strIniFile
                    End If
                End If
            End If
        End With
    Next lIndex
#End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cSettingsFile.Serialize", eGDRaiseError_Raise, g.strAppPath

End Sub

Private Sub Class_Terminate()
    Set m.SettingsFile = Nothing
End Sub

Private Function SetIniProp(ByVal strPropName$, ByVal vPropValue As Variant, ByVal strSection$, ByVal strIniFile$) As Variant
On Error GoTo ErrSection:

    Dim strValue As String              ' Value in the ini file
    
    strValue = GetIniFileProperty(strPropName, "DoesNotExist", strSection, strIniFile)
    
    If strValue <> "DoesNotExist" Or Len(Str(vPropValue)) > 0 Then
        SetIniFileProperty strPropName, vPropValue, strSection, strIniFile
    End If

ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cSettingsFile.SetIniProp", eGDRaiseError_Raise, g.strAppPath
    
End Function
