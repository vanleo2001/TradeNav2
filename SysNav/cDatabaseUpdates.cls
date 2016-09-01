VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDatabaseUpdates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
'This class is responsible for applying structural changes to the Libraries.MDB
'database.  The most recent version of the MDB is stored on the registry.  if the
'value in the registry is not current, one or more groups of changes are made

Private Const CURRENT_VERSION = 56

Private Type mPrivate
    DB     As Database
End Type
Private m As mPrivate

Private Sub Class_Terminate()
    Set m.DB = Nothing
End Sub

Private Property Get Version() As Long
On Error GoTo ErrSection:

    Dim mHeaderRS As Recordset
    
    'get database "Header" recordset (one record table)
    If ItemExists(m.DB.TableDefs, "tblDatabase") Then
        Set mHeaderRS = m.DB.OpenRecordset("tblDatabase", dbOpenDynaset)
        Version = mHeaderRS!Version
        
        If Version >= 25 Then
            
            ' just for upgrading purposes, we need to use whatever CustID
            ' was here in the first place (if this was a stolen database,
            ' we will upgrade, but still not allow it to be used)
            g.lLCD = mHeaderRS!CID
            
'TEMPORARY: to fix checksum
'If Version = 25 Then
'    mHeaderRS.Edit
'    mHeaderRS!CheckSum = BuildCheckSum(mHeaderRS, "tblDatabase")
'    mHeaderRS.Update
'End If
            
            ValidateCheckSums mHeaderRS, "tblDatabase"
            If mHeaderRS!CheckSum = 0.5 Then
                Version = -1
            End If
        End If
    Else
        'create one record table
        Dim t As TableDef
        Set t = New TableDef
        With t
            .Name = "tblDatabase"
            .Fields.Append .CreateField("Version", dbLong)
        End With
        m.DB.TableDefs.Append t
        Set mHeaderRS = m.DB.OpenRecordset("tblDatabase", dbOpenDynaset)
        With mHeaderRS
            .AddNew
            !Version = 0
            .Update
            .MoveFirst
        End With
        'also delete obsolete table
        If ItemExists(m.DB.TableDefs, "tblVersion") Then
            m.DB.TableDefs.Delete "tblVersion"
        End If
        Set t = Nothing
        Version = 0
    End If
    
    Set mHeaderRS = Nothing
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cDatabaseUpdates.Version.Get", eGDRaiseError_Raise
    
End Property

Private Property Let Version(ByVal nNewVersion As Long)
On Error GoTo ErrSection:

    Dim mHeaderRS As Recordset
    Set mHeaderRS = m.DB.OpenRecordset("tblDatabase", dbOpenDynaset)
    With mHeaderRS
        .Edit
        !Version = nNewVersion
        If nNewVersion >= 25 Then
            If !CheckSum <> 0.5 Then
                !CheckSum = BuildCheckSum(mHeaderRS, "tblDatabase")
            End If
        End If
        .Update
    End With
    Set mHeaderRS = Nothing
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cDatabaseUpdates.Version.Let", eGDRaiseError_Raise
    
End Property

Public Property Get DB() As Database
    Set DB = m.DB
End Property

Property Let DB(pData As Database)
On Error GoTo ErrSection:

    Dim iVersion&
    
    Set m.DB = pData
    
    'check version now just to make sure version table exists
    ''iVersion = Version
    
ErrExit:
    Exit Property
    
ErrSection:
    RaiseError "cDatabaseUpdates.DB.Let", eGDRaiseError_Raise
    
End Property

Public Function Upgrade() As Boolean
On Error GoTo ErrSection:
    
    Dim i&, v As Long, nStartVersion As Long
    Dim strBackup$, strUpgrading$
    Dim rs As Recordset

'Version = 1
    
    'See if need to upgrade
    g.lLCD = 0 '(just for upgrading purposes)
    nStartVersion = Version
    If nStartVersion < CURRENT_VERSION And nStartVersion > 0 Then
        'log message regarding upgrade
        MakeDir App.Path & "\Backup\", True
        FileFromString App.Path & "\Backup\Upgrade.log", _
            "Your database will now be upgraded (from v." _
            & CStr(nStartVersion) & " to v." & CStr(CURRENT_VERSION) & ").", _
            True, True
        frmSplash.Message -1, "Applying update to Libraries.MDB ..."
        
        'make a backup before any upgrading (and kill the really old backups)
        For i = 1 To nStartVersion - 2
            strBackup = App.Path & "\Backup\" & FileBase(m.DB.Name) & "." & Format(i, "000")
            If FileDate(strBackup) < Date - 90 Then
                KillFile strBackup
            End If
        Next
        strUpgrading = ReplaceFileExt(m.DB.Name, "_!_")
        KillFile strUpgrading
        mGenesis.CopyFiles m.DB.Name, strUpgrading
               
        ' first verify all checksums (so will set to 0.5 for anything that's been monkeyed with)
        If nStartVersion > 25 Then
            Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblDatabase];", dbOpenDynaset)
            ValidateCheckSums rs, "tblDatabase", g.lLCD
            Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblLibrarys];", dbOpenDynaset)
            ValidateCheckSums rs, "tblLibrarys", g.lLCD
            Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctions];", dbOpenDynaset)
            ValidateCheckSums rs, "tblFunctions", g.lLCD
            Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRules];", dbOpenDynaset)
            ValidateCheckSums rs, "tblRules", g.lLCD
            Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblSystems];", dbOpenDynaset)
            ValidateCheckSums rs, "tblSystems", g.lLCD
            Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblSystemRules];", dbOpenDynaset)
            ValidateCheckSums rs, "tblSystemRules", g.lLCD
        End If
        Set rs = Nothing
               
        'Apply each version of MDB changes...
        For v = nStartVersion + 1 To CURRENT_VERSION
            ApplyVersion v
            Version = v
        Next
    
        ' Fix all the checksums
        FixTableChecksums "tblDatabase"
        FixTableChecksums "tblLibrarys"
        FixTableChecksums "tblFunctions"
        FixTableChecksums "tblRules"
        FixTableChecksums "tblSystems"
        FixTableChecksums "tblSystemRules"
        FixTableChecksums "tblStrategyBaskets"
        FixTableChecksums "tblStrategyBasketItems"
           
        'rename upgrading to backup after upgrade is complete
        strBackup = App.Path & "\Backup\" & FileBase(m.DB.Name) & "." & Format(nStartVersion, "000")
        KillFile strBackup
        mGenesis.MoveFiles strUpgrading, strBackup
        g.bDirtyLibrariesMDB = True
        Upgrade = True
    Else
        Upgrade = False
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "cDatabaseUpdates.Upgrade", eGDRaiseError_Raise

End Function

'Apply MDB structural changes for version 1
Private Sub Apply1()
On Error GoTo ErrSection:
    
    Dim t           As TableDef
    Dim i           As Index
    Dim r           As Relation
    Dim q           As QueryDef

  Exit Sub 'don't do this anymore

    If ItemExists(m.DB.TableDefs, "tblMMSignalInputs") Then
        Exit Sub 'already done!
    End If

    '========================== TABLES ================================
    'Delete fields from "tblMMInputs"
    Set t = m.DB.TableDefs.Item("tblMMInputs")
    t.Indexes.Delete "OptListID"
    t.Fields.Delete "IfOptimize"
    t.Fields.Delete "OptFromValue"
    t.Fields.Delete "OptToValue"
    t.Fields.Delete "OptStepValue"
    t.Fields.Delete "OptListID"
    
    'Delete fields from "tblMMSystemInputs"
    Set t = m.DB.TableDefs.Item("tblMMSystemInputs")
    t.Indexes.Delete "OptListID"
    t.Fields.Delete "OptListID"
    
    'Create new table "tblMMSignalInputs"
    Set t = New TableDef
    With t
        .Name = "tblMMSignalInputs"
        .Fields.Append .CreateField("SystemNumber", dbLong)
        .Fields.Append .CreateField("SignalID", dbLong)
        .Fields.Append .CreateField("ParmID", dbLong)
        .Fields.Append .CreateField("Value", dbText)
        .Fields.Append .CreateField("IfOptimize", dbBoolean)
        .Fields.Append .CreateField("OptFromValue", dbDouble)
        .Fields.Append .CreateField("OptToValue", dbDouble)
        .Fields.Append .CreateField("OptStepValue", dbDouble)
    End With
    m.DB.TableDefs.Append t
    
    'Add Primary index for tblMMSignalInputs
    Set i = t.CreateIndex("Prim")
    With i
        .Fields.Append .CreateField("SystemNumber")
        .Fields.Append .CreateField("SignalID")
        .Fields.Append .CreateField("ParmID")
        .Primary = True
    End With
    t.Indexes.Append i
    
    '========================== RELATIONS ================================
    'Add relationship for: Signals and Signal inputs...
    'Set R = m.DB.CreateRelation("ExtSystemSignalsMMInputs", "tblMMExtSystemSignals", "tblMMSignalInputs", _
    '    dbRelationDeleteCascade)
    'With R
    '    .Fields.Append .CreateField("SystemNumber")
    '    .Fields.Append .CreateField("ID")
    '    .Fields.Item("SystemNumber").ForeignName = "SystemNumber"
    '    .Fields.Item("ID").ForeignName = "SignalID"
    'End With
    'm.DB.Relations.Append R
    
    'Add relationships for System MM Inputs
    Set r = m.DB.CreateRelation("SystemMMInputs", "tblSystems", "tblMMSystemInputs", _
        dbRelationDeleteCascade)
    With r
        .Fields.Append .CreateField("SystemNumber")
        .Fields.Item("SystemNumber").ForeignName = "SystemNumber"
    End With
    m.DB.Relations.Append r
    
    'Add relationships for System Rules and System Rule MM Inputs
    'Set R = m.DB.CreateRelation("SystemSignalsMMInputs", "tblSystemRules", "tblMMSignalInputs", _
    '    dbRelationDeleteCascade)
    'With R
    '    .Fields.Append .CreateField("SystemNumber")
    '    .Fields.Append .CreateField("RuleID")
    '    .Fields.Item("SystemNumber").ForeignName = "SystemNumber"
    '    .Fields.Item("RuleID").ForeignName = "SignalID"
    'End With
    'm.DB.Relations.Append R
    
    'Add relationships for Rule MM Inputs
    Set r = m.DB.CreateRelation("RuleMMInputs", "tblRuleParms", "tblMMSignalInputs", _
        dbRelationDeleteCascade)
    With r
        .Fields.Append .CreateField("ParmID")
        .Fields.Item("ParmID").ForeignName = "ParmID"
    End With
    m.DB.Relations.Append r
    
    '========================== QUERIES ================================
    'Remove query: qryMMGetMMInputs
    m.DB.QueryDefs.Delete "qryMMGetMMInputs"
    
    'Insert query: qryMMInputs
    Set q = m.DB.CreateQueryDef("qryMMInputs")
    q.SQL = "SELECT tblMM.MMID, tblRules.Name AS RuleName, tblMMInputs.Value, tblRuleParms.ParmID, tblRuleParms.ParmSeq, tblRuleParms.ParmName, tblRules.RuleID, tblRuleParms.ParmDesc, tblRuleParms.ParmTypeID, tblRuleParms.DefaultValue, tblRuleParms.FromValue, tblRuleParms.ToValue, tblRuleParms.ListID " & _
            "FROM tblRules INNER JOIN (tblRuleParms INNER JOIN (tblMM INNER JOIN tblMMInputs ON tblMM.MMID = tblMMInputs.MMid) ON tblRuleParms.ParmID = tblMMInputs.ParmID) ON tblRules.RuleID = tblRuleParms.RuleID " & _
            "WHERE (((tblMM.MMID) = [SelMMid])) ORDER BY tblMM.MMID;"
    
    'Insert new query:  qryMMSignalInputs
    Set q = m.DB.CreateQueryDef("qryMMSignalInputs")
    q.SQL = "SELECT tblMMSignalInputs.SystemNumber, tblMMSignalInputs.SignalID, tblRules.Name AS SignalName, tblMMSignalInputs.Value, tblMMSignalInputs.ParmID, tblRuleParms.RuleID, tblRuleParms.ParmSeq, tblRuleParms.ParmName, tblRuleParms.ParmDesc, tblRuleParms.ParmTypeID, tblRuleParms.DefaultValue, tblRuleParms.FromValue, tblRuleParms.ToValue, tblRuleParms.ListID, tblMMSignalInputs.IfOptimize, tblMMSignalInputs.OptFromValue, tblMMSignalInputs.OptToValue, tblMMSignalInputs.OptStepValue " & _
            "FROM tblRules INNER JOIN (tblRuleParms INNER JOIN tblMMSignalInputs ON tblRuleParms.ParmID = tblMMSignalInputs.ParmID) ON tblRules.RuleID = tblRuleParms.RuleID " & _
            "WHERE (((tblMMSignalInputs.SystemNumber) = [SelSystemNumber]) And ((tblMMSignalInputs.SignalID) = [SelSignalID])) " & _
            "ORDER BY tblMMSignalInputs.SystemNumber, tblMMSignalInputs.SignalID;"
    
    'Insert new query: qryDep12
    Set q = m.DB.CreateQueryDef("qryDep12")
    q.SQL = "SELECT tblLibrarys.LibraryName, First(tblLibrarys.LastModified) AS LastMod, First(tblLibrarys.Version) AS Vers " & _
            "FROM tblLibrarys INNER JOIN (tblMM INNER JOIN tblSystems ON tblMM.MMID = tblSystems.MMID) ON tblLibrarys.LibraryID = tblMM.LibraryID " & _
            "GROUP BY tblLibrarys.LibraryName, tblMM.LibraryID, tblSystems.LibraryID " & _
            "Having (((tblMM.LibraryID) <> [SelLibID]) And ((tblSystems.LibraryID) = [SelLibID])) " & _
            "ORDER BY tblSystems.LibraryID;"
    
    'Insert new query: qryDep13
    Set q = m.DB.CreateQueryDef("qryDep13")
    q.SQL = "SELECT tblLibrarys.LibraryName, First(tblLibrarys.LastModified) AS LastMod, First(tblLibrarys.Version) AS Vers " & _
            "FROM (tblLibrarys INNER JOIN tblRules ON tblLibrarys.LibraryID = tblRules.LibraryID) INNER JOIN ((tblSystems INNER JOIN tblMM ON tblSystems.MMID = tblMM.MMID) INNER JOIN tblMMRules ON tblMM.MMID = tblMMRules.MMid) ON tblRules.RuleID = tblMMRules.RuleID " & _
            "GROUP BY tblLibrarys.LibraryName, tblRules.LibraryID, tblSystems.LibraryID " & _
            "Having (((tblRules.LibraryID) <> [SelLibID]) And ((tblSystems.LibraryID) = [SelLibID])) " & _
            "ORDER BY tblSystems.LibraryID;"
    
    'Insert new query: qryDep14
    Set q = m.DB.CreateQueryDef("qryDep14")
    q.SQL = "SELECT tblLibrarys.LibraryName, First(tblLibrarys.LastModified) AS LastMod, First(tblLibrarys.Version) AS Vers " & _
            "FROM (tblRules INNER JOIN (tblLibrarys INNER JOIN (tblFunctions INNER JOIN tblFunctionRules ON tblFunctions.FunctionID = tblFunctionRules.FunctionIDRef) ON tblLibrarys.LibraryID = tblFunctions.LibraryID) ON tblRules.RuleID = tblFunctionRules.RuleID) INNER JOIN ((tblSystems INNER JOIN tblMM ON tblSystems.MMID = tblMM.MMID) INNER JOIN tblMMRules ON tblMM.MMID = tblMMRules.MMid) ON tblRules.RuleID = tblMMRules.RuleID " & _
            "GROUP BY tblLibrarys.LibraryName, tblFunctions.LibraryID, tblSystems.LibraryID " & _
            "Having (((tblFunctions.LibraryID) <> [SelLibID]) And ((tblSystems.LibraryID) = [SelLibID])) " & _
            "ORDER BY tblSystems.LibraryID;"
    
    'Insert new query: qryDep15
    Set q = m.DB.CreateQueryDef("qryDep15")
    q.SQL = "SELECT tblLibrarys.LibraryName, First(tblLibrarys.LastModified) AS LastMod, First(tblLibrarys.Version) AS Vers " & _
            "FROM (tblRules INNER JOIN ((tblFunctions INNER JOIN ((tblFunctionRefs INNER JOIN tblFunctions AS tblFunctions_1 ON tblFunctionRefs.FunctionIDRef = tblFunctions_1.FunctionID) INNER JOIN tblLibrarys ON tblFunctions_1.LibraryID = tblLibrarys.LibraryID) ON tblFunctions.FunctionID = tblFunctionRefs.FunctionID) INNER JOIN tblFunctionRules ON tblFunctions.FunctionID = tblFunctionRules.FunctionIDRef) ON tblRules.RuleID = tblFunctionRules.RuleID) INNER JOIN ((tblSystems INNER JOIN tblMM ON tblSystems.MMID = tblMM.MMID) INNER JOIN tblMMRules ON tblMM.MMID = tblMMRules.MMid) ON tblRules.RuleID = tblMMRules.RuleID " & _
            "GROUP BY tblLibrarys.LibraryName, tblFunctions_1.LibraryID, tblSystems.LibraryID " & _
            "Having (((tblFunctions_1.LibraryID) <> [SelLibID]) And ((tblSystems.LibraryID) = [SelLibID])) " & _
            "ORDER BY tblSystems.LibraryID;"
    
    'Insert new query: qryDep16
    Set q = m.DB.CreateQueryDef("qryDep16")
    q.SQL = "SELECT tblLibrarys.LibraryName, First(tblLibrarys.LastModified) AS LastMod, First(tblLibrarys.Version) AS Vers " & _
            "FROM (tblSystems INNER JOIN tblSystemRules ON tblSystems.SystemNumber = tblSystemRules.SystemNumber) INNER JOIN (tblLibrarys INNER JOIN tblMM ON tblLibrarys.LibraryID = tblMM.LibraryID) ON tblSystemRules.UnitsID = tblMM.MMID " & _
            "GROUP BY tblLibrarys.LibraryName, tblMM.LibraryID, tblSystems.LibraryID " & _
            "Having (((tblMM.LibraryID) <> [SelLibID]) And ((tblSystems.LibraryID) = [SelLibID])) " & _
            "ORDER BY tblSystems.LibraryID;"
    
    'Insert new query: qryDep17
    Set q = m.DB.CreateQueryDef("qryDep17")
    q.SQL = "SELECT tblLibrarys.LibraryName, First(tblLibrarys.LastModified) AS LastMod, First(tblLibrarys.Version) AS Vers " & _
            "FROM (tblLibrarys INNER JOIN tblRules ON tblLibrarys.LibraryID = tblRules.LibraryID) INNER JOIN ((tblSystems INNER JOIN (tblSystemRules INNER JOIN tblMM ON tblSystemRules.UnitsID = tblMM.MMID) ON tblSystems.SystemNumber = tblSystemRules.SystemNumber) INNER JOIN tblMMRules ON tblMM.MMID = tblMMRules.MMid) ON tblRules.RuleID = tblMMRules.RuleID " & _
            "GROUP BY tblLibrarys.LibraryName, tblRules.LibraryID, tblSystems.LibraryID " & _
            "Having (((tblRules.LibraryID) <> [SelLibID]) And ((tblSystems.LibraryID) = [SelLibID])) " & _
            "ORDER BY tblSystems.LibraryID;"
    
    'Insert new query: qryDep18
    Set q = m.DB.CreateQueryDef("qryDep18")
    q.SQL = "SELECT tblLibrarys.LibraryName, First(tblLibrarys.LastModified) AS LastMod, First(tblLibrarys.Version) AS Vers " & _
            "FROM tblLibrarys INNER JOIN (tblFunctions INNER JOIN ((tblSystems INNER JOIN (tblRules INNER JOIN ((tblSystemRules INNER JOIN tblMM ON tblSystemRules.UnitsID = tblMM.MMID) INNER JOIN tblMMRules ON tblMM.MMID = tblMMRules.MMid) ON tblRules.RuleID = tblMMRules.RuleID) ON tblSystems.SystemNumber = tblSystemRules.SystemNumber) INNER JOIN tblFunctionRules ON tblRules.RuleID = tblFunctionRules.RuleID) ON tblFunctions.FunctionID = tblFunctionRules.FunctionIDRef) ON tblLibrarys.LibraryID = tblFunctions.LibraryID " & _
            "GROUP BY tblLibrarys.LibraryName, tblFunctions.LibraryID, tblSystems.LibraryID " & _
            "Having (((tblFunctions.LibraryID) <> [SelLibID]) And ((tblSystems.LibraryID) = [SelLibID])) " & _
            "ORDER BY tblSystems.LibraryID;"
    
    'Insert new query: qryDep19
    Set q = m.DB.CreateQueryDef("qryDep19")
    q.SQL = "SELECT tblLibrarys.LibraryName, First(tblLibrarys.LastModified) AS LastMod, First(tblLibrarys.Version) AS Vers " & _
            "FROM (((tblSystems INNER JOIN ((tblRules INNER JOIN (tblFunctions INNER JOIN tblFunctionRules ON tblFunctions.FunctionID = tblFunctionRules.FunctionIDRef) ON tblRules.RuleID = tblFunctionRules.RuleID) INNER JOIN ((tblSystemRules INNER JOIN tblMM ON tblSystemRules.UnitsID = tblMM.MMID) INNER JOIN tblMMRules ON tblMM.MMID = tblMMRules.MMid) ON tblRules.RuleID = tblMMRules.RuleID) ON tblSystems.SystemNumber = tblSystemRules.SystemNumber) INNER JOIN tblFunctionRefs ON tblFunctions.FunctionID = tblFunctionRefs.FunctionID) INNER JOIN tblFunctions AS tblFunctions_1 ON tblFunctionRefs.FunctionIDRef = tblFunctions_1.FunctionID) INNER JOIN tblLibrarys ON tblFunctions_1.LibraryID = tblLibrarys.LibraryID " & _
            "GROUP BY tblLibrarys.LibraryName, tblFunctions_1.LibraryID, tblSystems.LibraryID " & _
            "Having (((tblFunctions_1.LibraryID) <> [SelLibID]) And ((tblSystems.LibraryID) = [SelLibID])) " & _
            "ORDER BY tblSystems.LibraryID;"
    
    'Insert new query: qryPackageSystemMMInputs
    Set q = m.DB.CreateQueryDef("qryPackageSystemMMInputs")
    q.SQL = "SELECT tblSystems.SystemName, tblRules.Name AS RuleName, tblRuleParms.ParmName, tblMMSystemInputs.Value, tblMMSystemInputs.IfOptimize, tblMMSystemInputs.OptFromValue, tblMMSystemInputs.OptToValue, tblMMSystemInputs.OptStepValue " & _
            "FROM tblSystems INNER JOIN ((tblRules INNER JOIN tblRuleParms ON tblRules.RuleID = tblRuleParms.RuleID) INNER JOIN tblMMSystemInputs ON tblRuleParms.ParmID = tblMMSystemInputs.ParmID) ON tblSystems.SystemNumber = tblMMSystemInputs.SystemNumber " & _
            "WHERE (((tblSystems.LibraryID)=[SelLibraryID]));"
    
    'Insert new query: qryPackageSystemRulesMMInputs
    Set q = m.DB.CreateQueryDef("qryPackageSystemRulesMMInputs")
    q.SQL = "SELECT tblSystems.SystemName, tblRules.Name AS RuleName, tblRuleParms.ParmName, tblMMSignalInputs.Value, tblMMSignalInputs.IfOptimize, tblMMSignalInputs.OptFromValue, tblMMSignalInputs.OptToValue, tblMMSignalInputs.OptStepValue " & _
            "FROM tblSystems INNER JOIN (tblSystemRules INNER JOIN (tblRules INNER JOIN (tblRuleParms INNER JOIN tblMMSignalInputs ON tblRuleParms.ParmID = tblMMSignalInputs.ParmID) ON tblRules.RuleID = tblRuleParms.RuleID) ON (tblSystemRules.RuleID = tblMMSignalInputs.SignalID) AND (tblSystemRules.SystemNumber = tblMMSignalInputs.SystemNumber)) ON tblSystems.SystemNumber = tblSystemRules.SystemNumber " & _
            "Where (((tblSystems.LibraryID) = [SelLibraryID])) " & _
            "ORDER BY tblSystems.SystemName, tblRules.Name, tblRuleParms.ParmName;"
            
    'Update SQL in: qryMMSystemInputs
    Set q = m.DB.QueryDefs("qryMMSystemInputs")
    q.SQL = "SELECT tblMMSystemInputs.SystemNumber, tblRules.Name AS Signal, tblMMSystemInputs.Value, tblRuleParms.RuleID, tblRuleParms.ParmID, tblRuleParms.ParmSeq, tblRuleParms.ParmName, tblRuleParms.ParmDesc, tblRuleParms.ParmTypeID, tblRuleParms.DefaultValue, tblRuleParms.FromValue, tblRuleParms.ToValue, tblRuleParms.ListID, tblMMSystemInputs.IfOptimize, tblMMSystemInputs.OptFromValue, tblMMSystemInputs.OptToValue, tblMMSystemInputs.OptStepValue " & _
            "FROM tblRules INNER JOIN (tblRuleParms INNER JOIN tblMMSystemInputs ON tblRuleParms.ParmID = tblMMSystemInputs.ParmID) ON tblRules.RuleID = tblRuleParms.RuleID " & _
            "WHERE (((tblMMSystemInputs.SystemNumber) = [SelSystemNumber])) " & _
            "ORDER BY tblMMSystemInputs.SystemNumber, tblRules.Name;"
            
    'Update SQL in: qryPackageMMRules
    Set q = m.DB.QueryDefs("qryPackageMMRules")
    q.SQL = "SELECT tblMM.Name, tblRules.Name AS RuleName, tblMMRules.Seq, tblMMRules.LastModKnown " & _
            "FROM tblRules INNER JOIN (tblMM INNER JOIN tblMMRules ON tblMM.MMID = tblMMRules.MMid) ON tblRules.RuleID = tblMMRules.RuleID " & _
            "WHERE (((tblMM.LibraryID)=[SelLibraryID]));"
            
    'Update SQL in: qryMMNavSignals
    Set q = m.DB.QueryDefs("qryMMNavSignals")
    q.SQL = "SELECT tblSystemRules.SystemNumber, tblRules.RuleID AS SignalID, tblRules.Name AS Signal, tblSystemRules.RuleUse AS SignalType, tblRules.BuySell, tblSystemRules.RuleUse, tblSystemRules.UnitsID, tblMM.Name AS MMName, tblMM.Notes AS MMPreview, tblRules.PreviewRTF AS SignalPreview, tblRules.SecurityLevel, tblRules.CannotDelete, tblRules.Password " & _
            "FROM tblRules INNER JOIN (tblMM RIGHT JOIN tblSystemRules ON tblMM.MMID = tblSystemRules.UnitsID) ON tblRules.RuleID = tblSystemRules.RuleID " & _
            "WHERE (((tblSystemRules.SystemNumber) = [SelSystemNumber])) " & _
            "ORDER BY tblSystemRules.SystemNumber, tblRules.RuleID;"
            
    'Update SQL in: qryPackageMMInputs
    Set q = m.DB.QueryDefs("qryPackageMMInputs")
    q.SQL = "SELECT tblMM.Name AS MMName, tblRules.Name AS RuleName, tblRuleParms.ParmName, tblMMInputs.Value " & _
            "FROM tblRules INNER JOIN (tblRuleParms INNER JOIN (tblMM INNER JOIN tblMMInputs ON tblMM.MMID = tblMMInputs.MMid) ON tblRuleParms.ParmID = tblMMInputs.ParmID) ON tblRules.RuleID = tblRuleParms.RuleID " & _
            "WHERE (((tblMM.LibraryID)=[SelLibraryID]));"

    Set q = Nothing
    Set r = Nothing
    Set i = Nothing
    Set t = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cDatabaseUpdates.Apply1", eGDRaiseError_Raise

End Sub

'Apply MDB structural changes for version XX
Private Sub ApplyVersion(ByVal nVersion As Long)
On Error GoTo ErrSection:
    
    Dim t           As TableDef
    Dim i           As Index
    Dim r           As Relation
    Dim q           As QueryDef
    Dim rs          As Recordset
    Dim rs2         As Recordset
    Dim strName     As String
    Dim bFlag       As Boolean
    Dim lIndex      As Long
    Dim lItems As Long
    Dim strText As String

    Select Case nVersion
    Case 1:
        Apply1
    
    Case 2:
        'create new pyramiding fields
        Set t = m.DB.TableDefs("tblSystems")
        With t
            If Not ItemExists(.Fields, "Pyramid") Then
                .Fields.Append .CreateField("Pyramid", dbBoolean)
                .Fields.Append .CreateField("TradeDepth", dbLong)
                .Fields("Pyramid").DefaultValue = 0
                .Fields("TradeDepth").DefaultValue = 50
                'fill new fields with defaults
                Set q = m.DB.CreateQueryDef("") 'temp update query
                q.SQL = "UPDATE tblSystems SET tblSystems.TradeDepth = 50, tblSystems.Pyramid = No;"
                q.Execute
            End If
        End With
        'Replace query: qryEngineSystem
        strName = "qryEngineSystem"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblSystems.SystemNumber, tblSystems.SystemName, tblSystems.BarTimeFrame, tblSystems.FromDate, tblSystems.ToDate, tblSystems.ToEndOfData, tblSystems.TradesPath, tblSystems.Expenses, tblSystems.LastModified, tblSystems.LibraryID, IIf(IsNull([TradeDepth]),1,(IIf(1>[TradeDepth],1,[TradeDepth]))) AS MaxTradeDepth, tblSystems.Pyramid" & _
                " FROM tblSystems WHERE (((tblSystems.SystemName)=[Enter System Name]));"
                
        'create new UsesHLCNextBar field
        Set t = m.DB.TableDefs("tblFunctions")
        With t
            If Not ItemExists(.Fields, "UsesHLCNextBar") Then
                .Fields.Append .CreateField("UsesHLCNextBar", dbBoolean)
                .Fields("UsesHLCNextBar").DefaultValue = 0
                'fill new fields with defaults
                Set q = m.DB.CreateQueryDef("") 'temp update query
                q.SQL = "UPDATE tblFunctions SET tblFunctions.UsesHLCNextBar = No;"
                q.Execute
            End If
        End With
        'Replace query: qryEngineFunction
        strName = "qryEngineFunction"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblFunctions.FunctionID, tblFunctions.FunctionName, tblFunctions.ImplementationTypeID, tblFunctions.ReturnTypeID, tblFunctions.CodedName, tblFunctions.CodedText, tblFunctions.Usage, tblFunctions.LateCalculating, tblFunctions.UsesOpenNextBar, tblFunctions.UsesMMData, tblFunctions.LastModified, tblFunctions.LibraryID, tblFunctions.UsesHLCNextBar FROM tblFunctions WHERE (((tblFunctions.LibraryID)=[Enter Library ID])) ORDER BY tblFunctions.FunctionName;"
    
    Case 3:
        'Replace query: qryFunctionsAndParms
        strName = "qryFunctionsAndParms"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblFunctions.FunctionID, tblFunctions.FunctionName, tblFunctions.Description, tblFunctions.FunctionCategoryID, tblFunctions.ImplementationTypeID, tblFunctions.ReturnTypeID, tblFunctions.CodedName, tblFunctions.Formatted, tblFunctions.FormattedWithFillWords, tblFunctions.CodedText, tblFunctions.Usage, tblFunctions.LateCalculating, tblFunctions.UsesOpenNextBar, tblFunctions.UsesMMData, tblFunctions.SecurityLevel, tblFunctions.CannotDelete, tblFunctions.Password, tblFunctions.LastModified, tblFunctions.LibraryID, tblFunctions.TradeSenseUsage, tblFunctionParms.ParmID, tblFunctionParms.ParmNbr, tblFunctionParms.ParmText, tblFunctionParms.ParmDesc, tblFunctionParms.ParmTypeID, tblFunctionParms.DefaultValue, tblFunctionParms.FromValue, tblFunctionParms.ToValue, tblFunctionParms.Required, tblFunctionParms.ListID, tblFunctionParms.FillPre, tblFunctionParms.FillPost, tblFunctions.Reverify, tblFunctions.Abbrev, tblFunctions.UsesHLCNextBar" _
                & " FROM tblFunctions LEFT JOIN tblFunctionParms ON tblFunctions.FunctionID = tblFunctionParms.FunctionID" _
                & " ORDER BY tblFunctions.FunctionName, tblFunctionParms.ParmNbr;"
        'Replace query: qryPackageFunctions
        strName = "qryPackageFunctions"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblFunctions.FunctionName, tblFunctions.Description, tblFunctions.FunctionCategoryID, tblFunctions.ImplementationTypeID, tblFunctions.ReturnTypeID, tblFunctions.TradeSenseUsage, tblFunctions.Formatted, tblFunctions.FormattedWithFillWords, tblFunctions.CodedName, tblFunctions.CodedText, tblFunctions.Usage, tblFunctions.LateCalculating, tblFunctions.UsesOpenNextBar, tblFunctions.UsesMMData, tblFunctions.SecurityLevel, tblFunctions.CannotDelete, tblFunctions.Password, tblFunctions.LastModified, tblFunctions.LibraryID, tblFunctions.Reverify, tblFunctions.Abbrev, tblFunctions.UsesHLCNextBar" _
                & " From tblFunctions" _
                & " Where (((tblFunctions.LibraryID) = [SelLibraryID]))" _
                & " ORDER BY tblFunctions.CodedName;"

    Case 4:
        'Delete "Abbrev" index
        Set t = m.DB.TableDefs("tblFunctions")
        If ItemExists(t.Indexes, "Abbrev") Then
            t.Indexes.Delete "Abbrev"
        End If

        'Add index for SystemName
        Set t = m.DB.TableDefs("tblSystems")
        If Not ItemExists(t.Indexes, "SystemName") Then
            Set i = t.CreateIndex("SystemName")
            With i
                .Fields.Append .CreateField("SystemName")
                .unique = True
                .Primary = False
            End With
            t.Indexes.Append i
        End If
        
        'Add "BuiltIn" field for libraries
        Set t = m.DB.TableDefs("tblLibrarys")
        With t
            If Not ItemExists(.Fields, "BuiltIn") Then
                .Fields.Append .CreateField("BuiltIn", dbBoolean)
                .Fields("BuiltIn").DefaultValue = 0
                'fill new field with defaults
                Set q = m.DB.CreateQueryDef("") 'temp update query
                q.SQL = "UPDATE tblLibrarys SET tblLibrarys.BuiltIn = Yes" _
                        & " WHERE (((tblLibrarys.LibraryName) Like 'Genesis *'));"
                q.Execute
            End If
        End With

        'Replace query: qryAllFunctions
        strName = "qryAllFunctions"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblFunctions.FunctionID, tblFunctions.Usage, tblFunctions.FunctionName, tblLibrarys.LibraryName, tblFunctionCategories.FunctionCategory, tblFunctions.LastModified, tblFunctions.Description, tblFunctions.SecurityLevel, tblFunctions.CannotDelete, tblFunctions.Password, tblFunctions.TradeSenseUsage, tblFunctions.ImplementationTypeID, tblFunctions.Reverify, tblFunctions.CodedName, tblFunctions.ReturnTypeID" _
                & " FROM tblLibrarys INNER JOIN (tblFunctionCategories INNER JOIN tblFunctions ON tblFunctionCategories.FunctionCategoryID = tblFunctions.FunctionCategoryID) ON tblLibrarys.LibraryID = tblFunctions.LibraryID" _
                & " WHERE (((tblFunctions.SecurityLevel)<>3));"

        'Add function categories
        Set rs = m.DB.OpenRecordset("tblFunctionCategories", dbOpenDynaset)
        With rs
            .AddNew
            !FunctionCategory = "BarsAgo"
            !CannotDelete = True
            .Update
        End With
    
    Case 5:
        'Replace query: qryFunctionsAndParms
        strName = "qryFunctionsAndParms"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblFunctions.FunctionID, tblFunctions.FunctionName, tblFunctions.Description, tblFunctions.FunctionCategoryID, tblFunctions.ImplementationTypeID, tblFunctions.ReturnTypeID, tblFunctions.CodedName, tblFunctions.Formatted, tblFunctions.FormattedWithFillWords, tblFunctions.CodedText, tblFunctions.Usage, tblFunctions.LateCalculating, tblFunctions.UsesOpenNextBar, tblFunctions.UsesMMData, tblFunctions.SecurityLevel, tblFunctions.CannotDelete, tblFunctions.Password, tblFunctions.LastModified, tblFunctions.LibraryID, tblFunctions.TradeSenseUsage, tblFunctionParms.ParmID, tblFunctionParms.ParmNbr, tblFunctionParms.ParmText, tblFunctionParms.ParmDesc, tblFunctionParms.ParmTypeID, tblFunctionParms.DefaultValue, tblFunctionParms.FromValue, tblFunctionParms.ToValue, tblFunctionParms.Required, tblFunctionParms.ListID, tblFunctionParms.FillPre, tblFunctionParms.FillPost, tblFunctions.Reverify, tblFunctions.Abbrev, tblFunctions.UsesHLCNextBar, tblFunctionCategories.FunctionCategory" _
                & " FROM tblFunctionCategories INNER JOIN (tblFunctions LEFT JOIN tblFunctionParms ON tblFunctions.FunctionID = tblFunctionParms.FunctionID) ON tblFunctionCategories.FunctionCategoryID = tblFunctions.FunctionCategoryID" _
                & " ORDER BY tblFunctions.FunctionName, tblFunctionParms.ParmNbr;"

    Case 6:
        'Add function categories (if not there yet)
        Set rs = m.DB.OpenRecordset("tblFunctionCategories", dbOpenDynaset)
        With rs
            bFlag = True
            .MoveFirst
            Do While Not .EOF
                If !FunctionCategory = "Data" Then
                    bFlag = False '(don't need to add)
                    Exit Do
                End If
                .MoveNext
            Loop
            If bFlag Then
                .AddNew
                !FunctionCategory = "Data"
                !CannotDelete = True
                .Update
                .AddNew
                !FunctionCategory = "Misc"
                !CannotDelete = True
                .Update
            End If
        End With
        
    Case 7:
        'Add new fields to tblRules, tblSystemRules, and tblSystems for pyramiding
        Set t = m.DB.TableDefs("tblSystems")
        With t
            If Not ItemExists(.Fields, "CanPyramidSameEntry") Then
                .Fields.Append .CreateField("CanPyramidSameEntry", dbBoolean)
            End If
        End With

        Set q = m.DB.CreateQueryDef("") 'temp update query
        q.SQL = "UPDATE tblSystems SET tblSystems.CanPyramidSameEntry = No;"
        q.Execute
        
        Set t = m.DB.TableDefs("tblRules")
        With t
            If Not ItemExists(.Fields, "ExitOnEntryBar") Then
                .Fields.Append .CreateField("ExitOnEntryBar", dbBoolean)
            End If
            If Not ItemExists(.Fields, "ExitBasedOnEachTrade") Then
                .Fields.Append .CreateField("ExitBasedOnEachTrade", dbBoolean)
            End If
            If Not ItemExists(.Fields, "NumberContracts") Then
                .Fields.Append .CreateField("NumberContracts", dbLong)
            End If
            If Not ItemExists(.Fields, "AsPercentOfPosition") Then
                .Fields.Append .CreateField("AsPercentOfPosition", dbBoolean)
            End If
        End With
        
        Set q = m.DB.CreateQueryDef("") 'temp update query
        q.SQL = "UPDATE tblRules SET tblRules.ExitOnEntryBar = No, " & _
                "tblRules.ExitBasedOnEachTrade = No, tblRules.NumberContracts = 1, " & _
                "tblRules.AsPercentOfPosition = No;"
        q.Execute
        
        Set q = m.DB.CreateQueryDef("")
        q.SQL = "UPDATE tblRules SET tblRules.RuleType = 1 " & _
                "WHERE tblRules.LateCondition = Yes OR tblRules.LateAction = Yes;"
        q.Execute
        
        Set t = m.DB.TableDefs("tblSystemRules")
        With t
            If Not ItemExists(.Fields, "ExitOnEntryBar") Then
                .Fields.Append .CreateField("ExitOnEntryBar", dbBoolean)
            End If
            If Not ItemExists(.Fields, "ExitBasedOnEachTrade") Then
                .Fields.Append .CreateField("ExitBasedOnEachTrade", dbBoolean)
            End If
            If Not ItemExists(.Fields, "NumberContracts") Then
                .Fields.Append .CreateField("NumberContracts", dbLong)
            End If
            If Not ItemExists(.Fields, "AsPercentOfPosition") Then
                .Fields.Append .CreateField("AsPercentOfPosition", dbBoolean)
            End If
        End With

        Set q = m.DB.CreateQueryDef("") 'temp update query
        q.SQL = "UPDATE tblSystemRules SET tblSystemRules.ExitOnEntryBar = No, " & _
                "tblSystemRules.ExitBasedOnEachTrade = No, tblSystemRules.NumberContracts = 1, " & _
                "tblSystemRules.AsPercentOfPosition = No;"
        q.Execute
        
        strName = "qryAllRegularRules"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblRules.RuleID, tblRules.Name, tblLibrarys.LibraryName, tblRules.RuleType, tblRules.BuySell, tblRules.LateCondition, tblRules.LateAction, tblRules.LastModified, tblRules.PreviewRTF, tblRules.CondRTF, tblRules.PriceRTF, tblRules.OrderPlacement, tblRules.SecurityLevel, tblRules.CannotDelete, tblRules.ExitOnEntryBar, tblRules.ExitBasedOnEachTrade, tblRules.NumberContracts, tblRules.AsPercentOfPosition, tblRules.Password, tblRules.Reverify " & _
                "FROM tblLibrarys INNER JOIN tblRules ON tblLibrarys.LibraryID = tblRules.LibraryID " & _
                "WHERE (((tblRules.RuleType)<2) AND ((tblRules.SecurityLevel)<>3));"
        
        strName = "qryEngineSystem"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblSystems.SystemNumber, tblSystems.SystemName, tblSystems.BarTimeFrame, tblSystems.FromDate, tblSystems.ToDate, tblSystems.ToEndOfData, tblSystems.TradesPath, tblSystems.Expenses, tblSystems.LastModified, tblSystems.LibraryID, IIf(IsNull([TradeDepth]),1,(IIf(1>[TradeDepth],1,[TradeDepth]))) AS MaxTradeDepth, tblSystems.Pyramid, tblSystems.CanPyramidSameEntry " & _
                "From tblSystems " & _
                "WHERE (((tblSystems.SystemName)=[Enter System Name]));"

        strName = "qryEngineSystemRule"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblSystemRules.SystemNumber AS SystemID, tblSystemRules.RuleID, tblSystemRules.Seq, tblRules.Name, tblRules.CondCoded, tblRules.PriceCoded, tblRules.Price2Coded, tblRules.BuySell, tblRules.LateCondition, tblRules.LateAction, tblRules.OrderPlacement, tblRules.RuleType, tblRules.LastModified, tblRules.LibraryID, tblSystemRules.Selected, tblSystemRules.Alternate, tblSystemRules.RuleUse, tblSystemRules.LastModifiedKnown, tblSystemRules.ExitOnEntryBar, tblSystemRules.ExitBasedOnEachTrade, tblSystemRules.NumberContracts, tblSystemRules.AsPercentOfPosition " & _
                "FROM tblRules INNER JOIN tblSystemRules ON tblRules.RuleID = tblSystemRules.RuleID " & _
                "WHERE (((tblSystemRules.SystemNumber)=[Enter System ID]));"

        strName = "qryPackageRules"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblRules.Name, tblRules.PreviewRTF, tblRules.CondRTF, tblRules.CondCoded, tblRules.PriceRTF, tblRules.PriceCoded, tblRules.Price2RTF, tblRules.Price2Coded, tblRules.OrderPlacement, tblRules.ActionCodedName, tblRules.BuySell, tblRules.ExitOnEntryBar, tblRules.ExitBasedOnEachTrade, tblRules.NumberContracts, tblRules.AsPercentOfPosition, tblRules.LateCondition, tblRules.LateAction, tblRules.RuleType, tblRules.SecurityLevel, tblRules.CannotDelete, tblRules.Password, tblRules.LastModified, tblRules.LibraryID, tblRules.Reverify " & _
                "From tblRules " & _
                "Where (((tblRules.LibraryID) = [SelLibraryID])) " & _
                "ORDER BY tblRules.Name;"

        strName = "qryPackageSystemRules"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblSystems.SystemName, tblRules.Name AS RuleName, tblSystemRules.Seq, tblSystemRules.Selected, tblSystemRules.Alternate, tblSystemRules.RuleUse, tblSystemRules.LastModifiedKnown, tblSystemRules.UnitsID, tblSystemRules.LinkedRules, tblSystemRules.ExitOnEntryBar, tblSystemRules.ExitBasedOnEachTrade, tblSystemRules.NumberContracts, tblSystemRules.AsPercentOfPosition " & _
                "FROM tblSystems INNER JOIN (tblRules INNER JOIN tblSystemRules ON tblRules.RuleID = tblSystemRules.RuleID) ON tblSystems.SystemNumber = tblSystemRules.SystemNumber " & _
                "WHERE (((tblSystems.LibraryID)=[SelLibraryID]));"

        strName = "qryPackageSystems"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblSystems.SystemName, tblSystems.BarTimeFrame, tblSystems.Expenses, tblSystems.FromDate, tblSystems.ToDate, tblSystems.ToEndOfData, tblSystems.LibraryID, tblSystems.Developer, tblSystems.TradesPath, tblSystems.Notes, tblSystems.SecurityLevel, tblSystems.CannotDelete, tblSystems.Password, tblSystems.LastModified, tblSystems.BarsLoadedBeforeTrading, tblSystems.BarsTradedBeforeOrders, tblSystems.Reverify, tblSystems.Pyramid, tblSystems.TradeDepth, tblSystems.CanPyramidSameEntry " & _
                "From tblSystems " & _
                "WHERE (((tblSystems.LibraryID)=[SelLibraryID]));"

        strName = "qrySystemRules"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblSystemRules.SystemNumber, tblSystemRules.RuleID, tblSystemRules.Seq, tblRules.RuleType, tblSystemRules.Selected, tblSystemRules.Alternate, tblSystemRules.LastModifiedKnown, tblSystemRules.UnitsID, tblRules.Name, tblRules.PreviewRTF, tblRules.CondRTF, tblRules.CondCoded, tblRules.PriceRTF, tblRules.PriceCoded, tblRules.Price2RTF, tblRules.Price2Coded, tblRules.ActionCodedName, tblRules.LastModified, tblRules.LibraryID, tblSystemRules.RuleUse, tblRules.BuySell, tblRules.LateCondition, tblRules.LateAction, tblRules.OrderPlacement, tblRules.SecurityLevel, tblRules.CannotDelete, tblRules.Password, tblSystemRules.LinkedRules, tblSystemRules.ExitOnEntryBar, tblSystemRules.ExitBasedOnEachTrade, tblSystemRules.NumberContracts, tblSystemRules.AsPercentOfPosition, tblRules.Reverify " & _
                "FROM tblRules INNER JOIN tblSystemRules ON tblRules.RuleID = tblSystemRules.RuleID " & _
                "Where (((tblSystemRules.SystemNumber) = [SelSystemNumber])) " & _
                "ORDER BY tblSystemRules.SystemNumber, tblSystemRules.Seq;"

    Case 8:
        ' Fixing these two queries that were screwed up in version 7.
        ' MAKE SURE THAT THE qryPackage??? HAS THE CORRECT FIELD ORDER, OTHERWISE
        ' THE AUTO IMPORT WILL NOT WORK CORRECTLY!!!!
        ' 5/23/2001 DAJ
        strName = "qryPackageRules"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblRules.Name, tblRules.PreviewRTF, tblRules.CondRTF, tblRules.CondCoded, tblRules.PriceRTF, tblRules.PriceCoded, tblRules.Price2RTF, tblRules.Price2Coded, tblRules.OrderPlacement, tblRules.ActionCodedName, tblRules.BuySell, tblRules.LateCondition, tblRules.LateAction, tblRules.RuleType, tblRules.SecurityLevel, tblRules.CannotDelete, tblRules.Password, tblRules.LastModified, tblRules.LibraryID, tblRules.Reverify, tblRules.ExitOnEntryBar, tblRules.ExitBasedOnEachTrade, tblRules.NumberContracts, tblRules.AsPercentOfPosition " & _
                "From tblRules " & _
                "Where (((tblRules.LibraryID) = [SelLibraryID])) " & _
                "ORDER BY tblRules.Name;"

        strName = "qryPackageSystems"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblSystems.SystemName, tblSystems.BarTimeFrame, tblSystems.Expenses, tblSystems.FromDate, tblSystems.ToDate, tblSystems.ToEndOfData, tblSystems.LibraryID, tblSystems.Developer, tblSystems.TradesPath, tblSystems.Notes, tblSystems.SecurityLevel, tblSystems.CannotDelete, tblSystems.Password, tblSystems.LastModified, tblSystems.BarsLoadedBeforeTrading, tblSystems.BarsTradedBeforeOrders, tblSystems.Reverify, tblSystems.MMID, tblSystems.Pyramid, tblSystems.TradeDepth, tblSystems.CanPyramidSameEntry " & _
                "From tblSystems " & _
                "WHERE (((tblSystems.LibraryID)=[SelLibraryID]));"
                
    Case 9:
        Set t = m.DB.TableDefs("tblDatabase")
        With t
            If Not ItemExists(.Fields, "ProductKey") Then
                .Fields.Append .CreateField("ProductKey", dbMemo)
                .Fields.Append .CreateField("MasterBuild", dbText)
                .Fields("ProductKey").AllowZeroLength = True
                .Fields("MasterBuild").AllowZeroLength = True
            End If
        End With
    
    Case 10:
        strName = "qryEngineSystemRule"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblSystemRules.SystemNumber AS SystemID, tblSystemRules.RuleID, tblSystemRules.Seq, tblRules.Name, tblRules.CondCoded, tblRules.PriceCoded, tblRules.Price2Coded, tblRules.BuySell, tblRules.LateCondition, tblRules.LateAction, tblRules.OrderPlacement, tblRules.RuleType, tblRules.LastModified, tblRules.LibraryID, tblSystemRules.Selected, tblSystemRules.Alternate, tblSystemRules.RuleUse, tblSystemRules.LastModifiedKnown, tblSystemRules.LinkedRules, tblSystemRules.ExitOnEntryBar, tblSystemRules.ExitBasedOnEachTrade, tblSystemRules.NumberContracts, tblSystemRules.AsPercentOfPosition " & _
                "FROM tblRules INNER JOIN tblSystemRules ON tblRules.RuleID = tblSystemRules.RuleID " & _
                "WHERE (((tblSystemRules.SystemNumber)=[Enter System ID]));"
                
    Case 11:
        Apply11
        
    Case 12:
        strName = "qryEngineSystemSecurity"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT DISTINCT tblSystemSecurities.SystemNumber AS SystemID, tblRuleParms.ParmName, tblSystemSecurities.SystemNumber, tblSystemSecurities.Path, tblSystemSecurities.Symbol, tblSystemSecurities.MarketSymbol, tblSystemSecurities.Periodicity, tblSystemSecurities.Format, tblSystemSecurities.SecurityType, tblSystemSecurities.SecurityName, IIf([SecurityType]=" & Chr(34) & "S" & Chr(34) & "," & Chr(34) & "!" & Chr(34) & ",IIf([SecurityType]=" & Chr(34) & "I" & Chr(34) & "," & Chr(34) & "$" & Chr(34) & ",[MarketSymbol])) AS JoinSymbol, tblSystems.BarTimeFrame " & _
                "FROM tblSystems INNER JOIN (tblRuleParms INNER JOIN tblSystemSecurities ON tblRuleParms.ParmID = tblSystemSecurities.ParmID) ON tblSystems.SystemNumber = tblSystemSecurities.SystemNumber " & _
                "WHERE (((tblSystemSecurities.SystemNumber)=[Enter System ID]));"
                
    Case 13:    ' Add DLLSecurityCode to tblLibrarys and related Queries
        Set t = m.DB.TableDefs("tblLibrarys")
        With t
            If Not ItemExists(.Fields, "DLLSecurityCode") Then
                .Fields.Append .CreateField("DLLSecurityCode", dbLong)
                .Fields("DLLSecurityCode").DefaultValue = 0
            End If
        End With
                
        Set q = m.DB.CreateQueryDef("") 'temp update query
        q.SQL = "UPDATE tblLibrarys SET tblLibrarys.DLLSecurityCode = 0;"
        q.Execute
        
        strName = "qryAllLibraries"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblLibrarys.LibraryID, tblLibrarys.LibraryName, tblLibrarys.LibraryDesc, tblLibrarys.Path, tblLibrarys.Version, tblLibrarys.LastModified, tblLibrarys.Author, tblLibrarys.SecurityLevel, tblLibrarys.CannotDelete, tblLibrarys.Password, tblLibrarys.rtfFileName, tblLibrarys.EMail, tblLibrarys.Phone, tblLibrarys.WebSite, tblLibrarys.LibraryType, tblLibrarys.BuiltIn, tblLibrarys.DLLSecurityCode " & _
                "FROM tblLibrarys;"

        strName = "qryEngineLibrary"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblLibrarys.LibraryID, tblLibrarys.LibraryName, tblLibrarys.Path, tblLibrarys.Version, tblLibrarys.LastModified, tblLibrarys.DLLSecurityCode, IIf(IsNull([tblLibrarys].[LibraryType]),0,[tblLibrarys].[LibraryType]) AS LibraryType " & _
                "FROM tblLibrarys " & _
                "ORDER BY tblLibrarys.LibraryID;"

        strName = "qryLibrarysByName"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblLibrarys.LibraryName, tblLibrarys.LibraryID, tblLibrarys.LibraryDesc, tblLibrarys.Path, tblLibrarys.Version, tblLibrarys.LastModified, tblLibrarys.Author, tblLibrarys.SecurityLevel, tblLibrarys.CannotDelete, tblLibrarys.Password, tblLibrarys.rtfFileName, tblLibrarys.EMail, tblLibrarys.Phone, tblLibrarys.WebSite, tblLibrarys.LibraryType, tblLibrarys.BuiltIn, tblLibrarys.DLLSecurityCode " & _
                "FROM tblLibrarys " & _
                "ORDER BY tblLibrarys.LibraryName;"
                
    Case 14:    ' Add tblFiles to the database
        If Not ItemExists(m.DB.TableDefs, "tblFiles") Then
            Set t = m.DB.CreateTableDef("tblFiles")
            With t
                If Not ItemExists(.Fields, "LibraryID") Then
                    .Fields.Append .CreateField("LibraryID", dbLong)
                    .Fields("LibraryID").Required = True
                End If
                If Not ItemExists(.Fields, "FileName") Then
                    .Fields.Append .CreateField("FileName", dbText)
                    .Fields("FileName").Required = True
                End If
                If Not ItemExists(.Fields, "FileSize") Then
                    .Fields.Append .CreateField("FileSize", dbLong)
                End If
                If Not ItemExists(.Fields, "FileDate") Then
                    .Fields.Append .CreateField("FileDate", dbDate)
                End If
                If Not ItemExists(.Fields, "Flags") Then
                    .Fields.Append .CreateField("Flags", dbLong)
                End If
                If Not ItemExists(.Fields, "FileInfo") Then
                    .Fields.Append .CreateField("FileInfo", dbMemo)
                    .Fields("FileInfo").AllowZeroLength = True
                End If
                
                If Not ItemExists(.Indexes, "Primary") Then
                    Set i = .CreateIndex("Primary")
                    With i
                        .Fields.Append .CreateField("LibraryID")
                        .Fields.Append .CreateField("FileName")
                        .Primary = True
                    End With
                    .Indexes.Append i
                End If
            End With
            m.DB.TableDefs.Append t
        End If
        
        If Not ItemExists(m.DB.Relations, "LibraryFiles") Then
            Set r = m.DB.CreateRelation("LibraryFiles", "tblLibrarys", "tblFiles", _
                dbRelationDeleteCascade)
            With r
                .Fields.Append .CreateField("LibraryID")
                .Fields.Item("LibraryID").ForeignName = "LibraryID"
            End With
            m.DB.Relations.Append r
        End If
        
        strName = "qryFilesByLibraryID"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblFiles.LibraryID, tblFiles.FileName, tblFiles.FileSize, tblFiles.FileDate, tblFiles.Flags, tblFiles.FileInfo " & _
                "FROM tblFiles " & _
                "WHERE LibraryID = [Enter Library ID];"
                    
    Case 15:
        ' This version was a temporary version that had some Equity Curve code
        ' in it.  This version only went to Mark S. and Glen, however we feel it
        ' safer to skip this version because it went to anybody at all.
                
    Case 16:    ' Added the "Linked Inputs" option
        Set t = m.DB.TableDefs("tblSystems")
        With t
            If Not ItemExists(.Fields, "LinkInputs") Then
                .Fields.Append .CreateField("LinkInputs", dbBoolean)
                
                Set q = m.DB.CreateQueryDef("") 'temp update query
                q.SQL = "UPDATE tblSystems SET tblSystems.LinkInputs = No;"
                q.Execute
            End If
        End With

        strName = "qryEngineSystem"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblSystems.SystemNumber, tblSystems.SystemName, tblSystems.BarTimeFrame, tblSystems.FromDate, tblSystems.ToDate, tblSystems.ToEndOfData, tblSystems.TradesPath, tblSystems.Expenses, tblSystems.LastModified, tblSystems.LibraryID, IIf(IsNull([TradeDepth]),1,(IIf(1>[TradeDepth],1,[TradeDepth]))) AS MaxTradeDepth, tblSystems.Pyramid, tblSystems.CanPyramidSameEntry, tblSystems.LinkInputs " & _
                "From tblSystems " & _
                "WHERE (((tblSystems.SystemName)=[Enter System Name]));"
    
        strName = "qryPackageSystems"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblSystems.SystemName, tblSystems.BarTimeFrame, tblSystems.Expenses, tblSystems.FromDate, tblSystems.ToDate, tblSystems.ToEndOfData, tblSystems.LibraryID, tblSystems.Developer, tblSystems.TradesPath, tblSystems.Notes, tblSystems.SecurityLevel, tblSystems.CannotDelete, tblSystems.Password, tblSystems.LastModified, tblSystems.BarsLoadedBeforeTrading, tblSystems.BarsTradedBeforeOrders, tblSystems.Reverify, tblSystems.Pyramid, tblSystems.TradeDepth, tblSystems.CanPyramidSameEntry, tblSystems.LinkInputs " & _
                "From tblSystems " & _
                "WHERE (((tblSystems.LibraryID)=[SelLibraryID]));"
    Case 17:
        strName = "qryPackageSystems"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblSystems.SystemName, tblSystems.BarTimeFrame, tblSystems.Expenses, tblSystems.FromDate, tblSystems.ToDate, tblSystems.ToEndOfData, tblSystems.LibraryID, tblSystems.Developer, tblSystems.TradesPath, tblSystems.Notes, tblSystems.SecurityLevel, tblSystems.CannotDelete, tblSystems.Password, tblSystems.LastModified, tblSystems.BarsLoadedBeforeTrading, tblSystems.BarsTradedBeforeOrders, tblSystems.Reverify, tblSystems.MMID, tblSystems.Pyramid, tblSystems.TradeDepth, tblSystems.CanPyramidSameEntry, tblSystems.LinkInputs " & _
                "From tblSystems " & _
                "WHERE (((tblSystems.LibraryID)=[SelLibraryID]));"
    
    Case 18:
        Set rs = DB.OpenRecordset("SELECT * FROM tblReports WHERE ReportID=18;", dbOpenDynaset)
        If rs.EOF Then
            Set q = m.DB.CreateQueryDef("") 'temp update query
            q.SQL = "INSERT INTO tblReports ( ReportID, Category, Seq, ReportName, ReportClass, Description ) " & _
                "SELECT 18 AS Expr1, 'System Reports' AS Expr2, 67 AS Expr3, 'By Trading Day of Year' as Expr4, 'cRptSummaryDetail' as Expr5, 'Shows a summary level snapshot of performance by Trading Day of Year' as Expr6;"
            q.Execute
        End If
        Set rs = DB.OpenRecordset("SELECT * FROM tblReports WHERE ReportID=19;", dbOpenDynaset)
        If rs.EOF Then
            Set q = m.DB.CreateQueryDef("") 'temp update query
            q.SQL = "INSERT INTO tblReports ( ReportID, Category, Seq, ReportName, ReportClass, Description ) " & _
                "SELECT 19 AS Expr1, 'System Reports' AS Expr2, 68 AS Expr3, 'By Trading Day of Month' as Expr4, 'cRptSummaryDetail' as Expr5, 'Shows a summary level snapshot of performance by Trading Day of Month' as Expr6;"
            q.Execute
        End If
        Set rs = DB.OpenRecordset("SELECT * FROM tblReports WHERE ReportID=20;", dbOpenDynaset)
        If rs.EOF Then
            Set q = m.DB.CreateQueryDef("") 'temp update query
            q.SQL = "INSERT INTO tblReports ( ReportID, Category, Seq, ReportName, ReportClass, Description ) " & _
                "SELECT 20 AS Expr1, 'System Reports' AS Expr2, 69 AS Expr3, 'By Trading Day of Week' as Expr4, 'cRptSummaryDetail' as Expr5, 'Shows a summary level snapshot of performance by Trading Day of Week' as Expr6;"
            q.Execute
        End If
        Set q = m.DB.CreateQueryDef("") 'temp update query
        q.SQL = "UPDATE tblReports SET tblReports.ReportName = 'By Year' " & _
            "WHERE (((tblReports.ReportID)=3));"
        q.Execute
        Set q = m.DB.CreateQueryDef("") 'temp update query
        q.SQL = "UPDATE tblReports SET tblReports.ReportName = 'By Month' " & _
            "WHERE (((tblReports.ReportID)=4));"
        q.Execute
        Set q = m.DB.CreateQueryDef("") 'temp update query
        q.SQL = "UPDATE tblReports SET tblReports.ReportName = 'By Week' " & _
            "WHERE (((tblReports.ReportID)=5));"
        q.Execute
        Set q = m.DB.CreateQueryDef("") 'temp update query
        q.SQL = "UPDATE tblReports SET tblReports.ReportName = 'By Day' " & _
            "WHERE (((tblReports.ReportID)=6));"
        q.Execute
        Set q = m.DB.CreateQueryDef("") 'temp update query
        q.SQL = "UPDATE tblReports SET tblReports.ReportName = 'By Weekday' " & _
            "WHERE (((tblReports.ReportID)=17));"
        q.Execute
        
        For lIndex = 18 To 20
            Set rs = DB.OpenRecordset("SELECT * FROM tblReportOptions WHERE ReportID=" & Trim(Str(lIndex)) & ";", dbOpenDynaset)
            If rs.EOF Then
                Set q = m.DB.CreateQueryDef("") 'temp update query
                q.SQL = "INSERT INTO tblReportOptions ( ReportID, Seq, OptionTypeID, DefaultValue ) " & _
                    "SELECT " & Trim(Str(lIndex)) & " AS Expr1, 1 AS Expr2, 1 AS Expr3, 'All Trades' as Expr4;"
                q.Execute
                Set q = m.DB.CreateQueryDef("") 'temp update query
                q.SQL = "INSERT INTO tblReportOptions ( ReportID, Seq, OptionTypeID, DefaultValue ) " & _
                    "SELECT " & Trim(Str(lIndex)) & " AS Expr1, 2 AS Expr2, 2 AS Expr3, '1/1/1900' as Expr4;"
                q.Execute
                Set q = m.DB.CreateQueryDef("") 'temp update query
                q.SQL = "INSERT INTO tblReportOptions ( ReportID, Seq, OptionTypeID, DefaultValue ) " & _
                    "SELECT " & Trim(Str(lIndex)) & " AS Expr1, 3 AS Expr2, 3 AS Expr3, '12/31/2400' as Expr4;"
                q.Execute
                Set q = m.DB.CreateQueryDef("") 'temp update query
                q.SQL = "INSERT INTO tblReportOptions ( ReportID, Seq, OptionTypeID, DefaultValue ) " & _
                    "SELECT " & Trim(Str(lIndex)) & " AS Expr1, 4 AS Expr2, 7 AS Expr3, '0' as Expr4;"
                q.Execute
                Set q = m.DB.CreateQueryDef("") 'temp update query
                q.SQL = "INSERT INTO tblReportOptions ( ReportID, Seq, OptionTypeID, DefaultValue ) " & _
                    "SELECT " & Trim(Str(lIndex)) & " AS Expr1, 5 AS Expr2, 4 AS Expr3, '-99999' as Expr4;"
                q.Execute
                Set q = m.DB.CreateQueryDef("") 'temp update query
                q.SQL = "INSERT INTO tblReportOptions ( ReportID, Seq, OptionTypeID, DefaultValue ) " & _
                    "SELECT " & Trim(Str(lIndex)) & " AS Expr1, 6 AS Expr2, 5 AS Expr3, '99999' as Expr4;"
                q.Execute
                Set q = m.DB.CreateQueryDef("") 'temp update query
                q.SQL = "INSERT INTO tblReportOptions ( ReportID, Seq, OptionTypeID, DefaultValue ) " & _
                    "SELECT " & Trim(Str(lIndex)) & " AS Expr1, 7 AS Expr2, 17 AS Expr3, '0' as Expr4;"
                q.Execute
                Set q = m.DB.CreateQueryDef("") 'temp update query
                q.SQL = "INSERT INTO tblReportOptions ( ReportID, Seq, OptionTypeID, DefaultValue ) " & _
                    "SELECT " & Trim(Str(lIndex)) & " AS Expr1, 8 AS Expr2, 18 AS Expr3, '0' as Expr4;"
                q.Execute
            End If
        Next lIndex

    Case 19:
        Set t = m.DB.TableDefs("tblSystems")
        With t
            If ItemExists(.Fields, "LinkInputs") Then
                .Fields("LinkInputs").DefaultValue = 0
            End If
            If ItemExists(.Fields, "CanPyramidSameEntry") Then
                .Fields("CanPyramidSameEntry").DefaultValue = 0
            End If
        End With
        Set t = m.DB.TableDefs("tblRules")
        With t
            If ItemExists(.Fields, "ExitOnEntryBar") Then
                .Fields("ExitOnEntryBar").DefaultValue = 0
            End If
            If ItemExists(.Fields, "ExitBasedOnEachTrade") Then
                .Fields("ExitBasedOnEachTrade").DefaultValue = 0
            End If
            If ItemExists(.Fields, "NumberContracts") Then
                .Fields("NumberContracts").DefaultValue = 1
            End If
            If ItemExists(.Fields, "AsPercentOfPosition") Then
                .Fields("AsPercentOfPosition").DefaultValue = 0
            End If
        End With
        Set t = m.DB.TableDefs("tblSystemRules")
        With t
            If ItemExists(.Fields, "ExitOnEntryBar") Then
                .Fields("ExitOnEntryBar").DefaultValue = 0
            End If
            If ItemExists(.Fields, "ExitBasedOnEachTrade") Then
                .Fields("ExitBasedOnEachTrade").DefaultValue = 0
            End If
            If ItemExists(.Fields, "NumberContracts") Then
                .Fields("NumberContracts").DefaultValue = 1
            End If
            If ItemExists(.Fields, "AsPercentOfPosition") Then
                .Fields("AsPercentOfPosition").DefaultValue = 0
            End If
        End With
    
    Case 20: ' Reset the report option default values to bigger values
        Set q = m.DB.CreateQueryDef("") 'temp update query
        q.SQL = "UPDATE tblReportOptions SET tblReportOptions.DefaultValue = '999999' " & _
            "WHERE (((tblReportOptions.DefaultValue)='99999'));"
        q.Execute
        Set q = m.DB.CreateQueryDef("") 'temp update query
        q.SQL = "UPDATE tblReportOptions SET tblReportOptions.DefaultValue = '-999999' " & _
            "WHERE (((tblReportOptions.DefaultValue)='-99999'));"
        q.Execute
        
    Case 21: ' Add Module, Expiration, and Ignore fields to tblLibrarys
        Set t = m.DB.TableDefs("tblLibrarys")
        With t
            If Not ItemExists(.Fields, "RequiredMod") Then
                .Fields.Append .CreateField("RequiredMod", dbText, 10)
                With .Fields("RequiredMod")
                    .AllowZeroLength = True
                    .DefaultValue = ""
                    .Required = False
                End With
            End If
            If Not ItemExists(.Fields, "Expiration") Then
                .Fields.Append .CreateField("Expiration", dbDate)
                With .Fields("Expiration")
                    .DefaultValue = 0
                    .Required = False
                End With
                Set q = m.DB.CreateQueryDef("")
                q.SQL = "UPDATE tblLibrarys SET Expiration=0"
                q.Execute
            End If
            If Not ItemExists(.Fields, "Ignore") Then
                .Fields.Append .CreateField("Ignore", dbBoolean)
                With .Fields("Ignore")
                    .DefaultValue = False
                    .Required = False
                End With
            End If
        End With
        
        strName = "qryLibrarysByName"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblLibrarys.LibraryName, tblLibrarys.LibraryID, tblLibrarys.LibraryDesc, tblLibrarys.Path, tblLibrarys.Version, tblLibrarys.LastModified, tblLibrarys.Author, tblLibrarys.SecurityLevel, tblLibrarys.CannotDelete, tblLibrarys.Password, tblLibrarys.rtfFileName, tblLibrarys.EMail, tblLibrarys.Phone, tblLibrarys.WebSite, tblLibrarys.LibraryType, tblLibrarys.BuiltIn, tblLibrarys.DLLSecurityCode, tblLibrarys.RequiredMod, tblLibrarys.Expiration, tblLibrarys.Ignore " & _
                "FROM tblLibrarys " & _
                "WHERE (((tblLibrarys.Ignore) = 0)) " & _
                "ORDER BY tblLibrarys.LibraryName;"

        strName = "qryEngineLibrary"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblLibrarys.LibraryID, tblLibrarys.LibraryName, tblLibrarys.Path, tblLibrarys.Version, tblLibrarys.LastModified, tblLibrarys.DLLSecurityCode, tblLibrarys.RequiredMod, tblLibrarys.Expiration, tblLibrarys.Ignore, IIf(IsNull([tblLibrarys].[LibraryType]),0,[tblLibrarys].[LibraryType]) AS LibraryType " & _
                "FROM tblLibrarys " & _
                "WHERE (((tblLibrarys.Ignore) = 0)) " & _
                "ORDER BY tblLibrarys.LibraryID;"
        
        strName = "qryFunctionsAndParms"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblFunctions.FunctionID, tblFunctions.FunctionName, tblFunctions.Description, tblFunctions.FunctionCategoryID, tblFunctions.ImplementationTypeID, tblFunctions.ReturnTypeID, tblFunctions.CodedName, tblFunctions.Formatted, " & _
                "tblFunctions.FormattedWithFillWords, tblFunctions.CodedText, tblFunctions.Usage, tblFunctions.LateCalculating, tblFunctions.UsesOpenNextBar, tblFunctions.UsesMMData, tblFunctions.SecurityLevel, tblFunctions.CannotDelete, tblFunctions.Password, tblFunctions.LastModified, tblFunctions.LibraryID, tblFunctions.TradeSenseUsage, tblFunctionParms.ParmID, tblFunctionParms.ParmNbr, tblFunctionParms.ParmText, tblFunctionParms.ParmDesc, tblFunctionParms.ParmTypeID, tblFunctionParms.DefaultValue, tblFunctionParms.FromValue, tblFunctionParms.ToValue, tblFunctionParms.Required, tblFunctionParms.ListID, tblFunctionParms.FillPre, tblFunctionParms.FillPost, tblFunctions.Reverify, tblFunctions.Abbrev, tblFunctions.UsesHLCNextBar, tblFunctionCategories.FunctionCategory, tblFunctions.ListID AS FListID, " & _
                "tblParmTypes.ParmUsage AS IParmUsage " & _
                "FROM tblLibrarys INNER JOIN (tblParmTypes RIGHT JOIN ((tblFunctionCategories INNER JOIN tblFunctions ON tblFunctionCategories.FunctionCategoryID = tblFunctions.FunctionCategoryID) LEFT JOIN tblFunctionParms ON tblFunctions.FunctionID = tblFunctionParms.FunctionID) ON tblParmTypes.ParmTypeID = tblFunctionParms.ParmTypeID) ON tblLibrarys.LibraryID = tblFunctions.LibraryID " & _
                "WHERE (((tblLibrarys.Ignore) = 0)) " & _
                "ORDER BY tblFunctions.FunctionName, tblFunctionParms.ParmNbr;"
                
    Case 22:    ' Added SystemNumber to tblRules and updated queries accordingly
        Set t = m.DB.TableDefs("tblRules")
        With t
            If Not ItemExists(.Fields, "SystemNumber") Then
                .Fields.Append .CreateField("SystemNumber", dbLong)
                With .Fields("SystemNumber")
                    .DefaultValue = 0
                    .Required = False
                End With
        
                Set q = m.DB.CreateQueryDef("") 'temp update query
                q.SQL = "UPDATE tblRules SET tblRules.SystemNumber = 0;"
                q.Execute
            End If
        End With
                        
        strName = "qryAllRegularRules"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblRules.*, tblLibrarys.LibraryName " & _
                "FROM tblLibrarys INNER JOIN tblRules ON tblLibrarys.LibraryID = tblRules.LibraryID " & _
                "WHERE (((tblRules.RuleType)<2) AND ((tblRules.SecurityLevel)<>3));"

        strName = "qryPackageRules"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblRules.Name, tblRules.PreviewRTF, tblRules.CondRTF, tblRules.CondCoded, tblRules.PriceRTF, tblRules.PriceCoded, tblRules.Price2RTF, tblRules.Price2Coded, tblRules.OrderPlacement, tblRules.ActionCodedName, tblRules.BuySell, tblRules.LateCondition, tblRules.LateAction, tblRules.RuleType, tblRules.SecurityLevel, tblRules.CannotDelete, tblRules.Password, tblRules.LastModified, tblRules.LibraryID, tblRules.Reverify, tblRules.ExitOnEntryBar, tblRules.ExitBasedOnEachTrade, tblRules.NumberContracts, tblRules.AsPercentOfPosition, tblRules.SystemNumber " & _
                "FROM tblRules " & _
                "WHERE (((tblRules.LibraryID) = [SelLibraryID])) " & _
                "ORDER BY tblRules.Name;"

        strName = "qrySystemRules"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblSystemRules.SystemNumber, tblSystemRules.RuleID, tblSystemRules.Seq, tblRules.RuleType, tblSystemRules.Selected, tblSystemRules.Alternate, tblSystemRules.LastModifiedKnown, tblSystemRules.UnitsID, tblRules.Name, tblRules.PreviewRTF, tblRules.CondRTF, tblRules.CondCoded, tblRules.PriceRTF, tblRules.PriceCoded, tblRules.Price2RTF, tblRules.Price2Coded, tblRules.ActionCodedName, tblRules.LastModified, tblRules.LibraryID, tblSystemRules.RuleUse, tblRules.BuySell, tblRules.LateCondition, tblRules.LateAction, tblRules.OrderPlacement, tblRules.SecurityLevel, tblRules.CannotDelete, tblRules.Password, tblSystemRules.LinkedRules, tblSystemRules.ExitOnEntryBar, tblSystemRules.ExitBasedOnEachTrade, tblSystemRules.NumberContracts, tblSystemRules.AsPercentOfPosition, tblRules.Reverify, tblRules.SystemNumber AS SystemNbr " & _
                "FROM tblRules INNER JOIN tblSystemRules ON tblRules.RuleID = tblSystemRules.RuleID " & _
                "WHERE (((tblSystemRules.SystemNumber) = [SelSystemNumber])) " & _
                "ORDER BY tblSystemRules.SystemNumber, tblSystemRules.Seq;"
        
        ConvertRules
        
    Case 23:
        Set t = m.DB.TableDefs("tblRules")
        With t
            .Indexes.Delete "Name"
            Set i = t.CreateIndex("Name")
            With i
                .Fields.Append .CreateField("Name")
                .unique = False
                .Primary = False
            End With
            t.Indexes.Append i
        End With
        
    Case 24:
        FixRulePreview
        strName = "qryPackageFiles"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblFiles.LibraryID, tblFiles.FileName, tblFiles.FileSize, tblFiles.FileDate, tblFiles.Flags, tblFiles.FileInfo " & _
                "FROM tblFiles " & _
                "WHERE (((tblFiles.LibraryID)=[SelLibraryID]));"
                
    Case 25:
        Apply25
        
    Case 26:
        Set t = m.DB.TableDefs("tblSystems")
        With t
            If Not ItemExists(.Fields, "NumShares") Then
                .Fields.Append .CreateField("NumShares", dbLong)
                With .Fields("NumShares")
                    .DefaultValue = 100&
                    .Required = False
                End With
            End If
            
            If Not ItemExists(.Fields, "StockExpenses") Then
                .Fields.Append .CreateField("StockExpenses", dbCurrency)
                With .Fields("StockExpenses")
                    .DefaultValue = 15#
                    .Required = False
                End With
            End If
        End With
        
        Set rs = m.DB.OpenRecordset("SELECT * FROM [tblSystems];", dbOpenDynaset)
        Do While Not rs.EOF
            rs.Edit
            rs!NumShares = 100&
            rs!StockExpenses = 15#
            'rs!Checksum = BuildCheckSum(rs, "tblSystems")
            rs.Update
            
            rs.MoveNext
        Loop
        
        strName = "qryPackageSystems"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblSystems.SystemName, tblSystems.BarTimeFrame, tblSystems.Expenses, tblSystems.FromDate, tblSystems.ToDate, tblSystems.ToEndOfData, tblSystems.LibraryID, tblSystems.Developer, tblSystems.TradesPath, tblSystems.Notes, tblSystems.SecurityLevel, tblSystems.CannotDelete, tblSystems.Password, tblSystems.LastModified, tblSystems.BarsLoadedBeforeTrading, tblSystems.BarsTradedBeforeOrders, tblSystems.Reverify, tblSystems.MMID, tblSystems.Pyramid, tblSystems.TradeDepth, tblSystems.CanPyramidSameEntry, tblSystems.LinkInputs, tblSystems.CheckSum, tblSystems.NumShares, tblSystems.StockExpenses " & _
                "FROM tblSystems " & _
                "WHERE (((tblSystems.LibraryID)=[SelLibraryID]));"
    
    Case 27:
        Set rs = m.DB.OpenRecordset("SELECT * FROM [tblLibrarys];", dbOpenDynaset)
        Do While Not rs.EOF
            rs.Edit
            rs!Version = "1"
            'rs!Checksum = BuildCheckSum(rs, "tblLibrarys")
            rs.Update
            
            rs.MoveNext
        Loop
        rs.Close
        
        Set t = m.DB.TableDefs("tblLibrarys")
        With t
            With .Fields("Version")
                .DefaultValue = "1"
            End With
        End With
        
    Case 28:    ' new tblMarketInfo
        If Not ItemExists(m.DB.TableDefs, "tblMarketInfo") Then
            Set t = m.DB.CreateTableDef("tblMarketInfo")
            With t
                If Not ItemExists(.Fields, "Symbol") Then
                    .Fields.Append .CreateField("Symbol", dbText, 14)
                    .Fields("Symbol").Required = True
                End If
                If Not ItemExists(.Fields, "SymbolID") Then
                    .Fields.Append .CreateField("SymbolID", dbLong)
                    .Fields("SymbolID").Required = True
                End If
                If Not ItemExists(.Fields, "DataType") Then
                    .Fields.Append .CreateField("DataType", dbByte)
                    .Fields("DataType").Required = True
                End If
                If Not ItemExists(.Fields, "Value") Then
                    .Fields.Append .CreateField("Value", dbText, 30)
                    .Fields("Value").AllowZeroLength = True
                End If
                
                If Not ItemExists(.Indexes, "Primary") Then
                    Set i = .CreateIndex("Primary")
                    With i
                        .Fields.Append .CreateField("Symbol")
                        .Fields.Append .CreateField("SymbolID")
                        .Fields.Append .CreateField("DataType")
                        .Primary = True
                    End With
                    .Indexes.Append i
                End If
            End With
            m.DB.TableDefs.Append t
        End If
        
    Case 29:
        lItems = 0
        Set rs = m.DB.OpenRecordset("SELECT * FROM [tblRules];")
        rs.MoveLast
        lItems = rs.RecordCount * 2
        Set rs = m.DB.OpenRecordset("SELECT * FROM [tblSystemRules];")
        rs.MoveLast
        lItems = lItems + rs.RecordCount * 2
        Set rs = m.DB.OpenRecordset("SELECT * FROM [tblSystems];")
        rs.MoveLast
        lItems = lItems + rs.RecordCount
    
        InfBox "i=t ; w=NOWAIT ; Trade Navigator is updating your database.  Please allow several minutes for this to happen.", , , , , , , , , , , , , lItems + 1
    
        ' Add a RequiredData Field to the Functions Table for including Data
        ' Strings that are required in the customer's available string in order for
        ' the customer to see the function
        Set t = m.DB.TableDefs("tblFunctions")
        With t
            If Not ItemExists(.Fields, "RequiredData") Then
                .Fields.Append .CreateField("RequiredData", dbText, 10)
                .Fields("RequiredData").Required = False
                .Fields("RequiredData").AllowZeroLength = True
    
                Set rs = m.DB.OpenRecordset("SELECT * FROM [tblFunctions];", dbOpenDynaset)
                Do While Not rs.EOF
                    rs.Edit
                    rs!RequiredData = ""
                    'rs!Checksum = BuildCheckSum(rs, "tblFunctions")
                    rs.Update
    
                    rs.MoveNext
                Loop
            End If
        End With
        
        ' Make sure to export the System Name for the given System Number since the
        ' System Number will not be guaranteed to be the same upon import
        strName = "qryPackageRules"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblRules.Name, tblRules.PreviewRTF, tblRules.CondRTF, tblRules.CondCoded, tblRules.PriceRTF, tblRules.PriceCoded, tblRules.Price2RTF, tblRules.Price2Coded, tblRules.OrderPlacement, tblRules.ActionCodedName, tblRules.BuySell, tblRules.LateCondition, tblRules.LateAction, tblRules.RuleType, tblRules.SecurityLevel, tblRules.CannotDelete, tblRules.Password, tblRules.LastModified, tblRules.LibraryID, tblRules.Reverify, tblRules.ExitOnEntryBar, tblRules.ExitBasedOnEachTrade, tblRules.NumberContracts, tblRules.AsPercentOfPosition, IIf(IsNull([tblSystems].[SystemName]),0,[tblSystems].[SystemName]) AS SystemName, tblRules.CheckSum " & _
                "FROM tblRules LEFT JOIN tblSystems ON tblRules.SystemNumber = tblSystems.SystemNumber " & _
                "WHERE (((tblRules.LibraryID) = [SelLibraryID])) " & _
                "ORDER BY tblRules.Name;"
        
        ' Need to fix this query to include the Name of the System that contains the
        ' rule since the Rule Name is no longer necessarily unique (However, the Rule
        ' Name/System Name combo should be unique)
        strName = "qryPackageRuleParms"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblRules.Name, IIf(IsNull(tblSystems.SystemName),0,tblSystems.SystemName) AS RuleSystemName, tblRuleParms.ParmSeq, tblRuleParms.ParmName, tblRuleParms.ParmText, tblRuleParms.ParmDesc, tblRuleParms.ParmTypeID, tblRuleParms.DefaultValue, tblRuleParms.FromValue, tblRuleParms.ToValue, tblRuleParms.ListID " & _
                "FROM (tblRules INNER JOIN tblRuleParms ON tblRules.RuleID = tblRuleParms.RuleID) LEFT JOIN tblSystems ON tblRules.SystemNumber = tblSystems.SystemNumber " & _
                "WHERE (((tblRules.LibraryID) = [SelLibraryID])) " & _
                "ORDER BY tblRules.Name, tblSystems.SystemName, tblRuleParms.ParmSeq;"
                
        ' Make a local copy of all shared rules and retain Built-In Shared rules and
        ' any rules that the user had not used in a system as "favorites"
        MoveAllRulesLocal lItems
        
        ' Convert all of the Market information that may not already be ChartNav
        ' Data to ChartNav data
        ConvertAllMarketsToCN
        
        InfBox ""
        
    Case 30:
        Set t = m.DB.TableDefs("tblLibrarys")
        With t
            If Not ItemExists(.Fields, "LastExportedVersion") Then
                .Fields.Append .CreateField("LastExportedVersion", dbText, 10)
                .Fields("LastExportedVersion").Required = False
                .Fields("LastExportedVersion").AllowZeroLength = True
    
                Set rs = m.DB.OpenRecordset("SELECT * FROM [tblLibrarys];", dbOpenDynaset)
                Do While Not rs.EOF
                    rs.Edit
                    rs!LastExportedVersion = rs!Version
                    'rs!Checksum = BuildCheckSum(rs, "tblLibrarys")
                    rs.Update
    
                    rs.MoveNext
                Loop
            End If
        End With
           
        ' Upgrade 28 copied security stuff into rules in the User library, so we will
        ' make sure here that all of the stuff in the User Library has no security...
        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblSystems] WHERE [LibraryID]=8;", dbOpenDynaset)
        Do While Not rs.EOF
            rs.Edit
            rs!SecurityLevel = 0
            rs!CannotDelete = False
            rs!Password = ""
            'rs!Checksum = BuildCheckSum(rs, "tblSystems")
            rs.Update
            
            rs.MoveNext
        Loop
    
        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctions] WHERE [LibraryID]=8;", dbOpenDynaset)
        Do While Not rs.EOF
            rs.Edit
            rs!SecurityLevel = 0
            rs!CannotDelete = False
            rs!Password = ""
            'rs!Checksum = BuildCheckSum(rs, "tblFunctions")
            rs.Update
            
            rs.MoveNext
        Loop
    
        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRules] WHERE [LibraryID]=8;", dbOpenDynaset)
        Do While Not rs.EOF
            rs.Edit
            rs!SecurityLevel = 0
            rs!CannotDelete = False
            rs!Password = ""
            'rs!Checksum = BuildCheckSum(rs, "tblRules")
            rs.Update
            
            rs.MoveNext
        Loop
    
    Case 31:
        ChangeDataType "tblSystemRules", "LinkedRules", dbMemo
        
    Case 32:
        ' Needed to swap the ExitOnEntryBar and LinkedRules columns...
        strName = "qryPackageSystemRules"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblSystems.SystemName, tblRules.Name AS RuleName, tblSystemRules.Seq, tblSystemRules.Selected, tblSystemRules.Alternate, tblSystemRules.RuleUse, tblSystemRules.LastModifiedKnown, tblSystemRules.UnitsID, tblSystemRules.ExitOnEntryBar, tblSystemRules.LinkedRules, tblSystemRules.ExitBasedOnEachTrade, tblSystemRules.NumberContracts, tblSystemRules.AsPercentOfPosition, tblSystemRules.CheckSum " & _
                "FROM tblSystems INNER JOIN (tblRules INNER JOIN tblSystemRules ON tblRules.RuleID = tblSystemRules.RuleID) ON tblSystems.SystemNumber = tblSystemRules.SystemNumber " & _
                "WHERE (((tblSystems.LibraryID)=[SelLibraryID]));"
        ' and fix any DB's that may have had the columns in the wrong order
        ' (we're not sure if this has happened, but just in case)
        Set t = m.DB.TableDefs("tblSystemRules")
        If t.Fields("ExitOnEntryBar").OrdinalPosition = 10 Then
            t.Fields("ExitOnEntryBar").OrdinalPosition = 9
            t.Fields("LinkedRules").OrdinalPosition = 10
        End If
            
    Case 33:
        ' must change data type for all fields with binary data
        ' (so binary data can be read correctly by DBCS and Unicode)
        ChangeDataType "tblRules", "CondRTF"
        ChangeDataType "tblRules", "PreviewRTF"
        ChangeDataType "tblRules", "CondCoded"
        ChangeDataType "tblRules", "PriceRTF"
        ChangeDataType "tblRules", "PriceCoded"
        ChangeDataType "tblRules", "Price2RTF"
        ChangeDataType "tblRules", "Price2Coded"
        ChangeDataType "tblRules", "Password"
        
        ChangeDataType "tblFunctions", "Formatted"
        ChangeDataType "tblFunctions", "FormattedWithFillWords"
        ChangeDataType "tblFunctions", "CodedText"
        ChangeDataType "tblFunctions", "Password"
        
        ChangeDataType "tblLibrarys", "Password"
        ChangeDataType "tblSystems", "Password"
        
        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblLibrarys];", dbOpenDynaset)
        Do While Not rs.EOF
            strText = NullChk(rs!RequiredMod)
            If Len(strText) > 0 Then
                rs.Edit
                rs!RequiredMod = DecryptField(rs!RequiredMod)
                rs.Update
            End If
            rs.MoveNext
        Loop
        
    Case 34:
        ' fix any DB's that may have had the columns in the wrong order
        ' (this happened from original "ChangeDataType" -- but fixed now)
        Set t = m.DB.TableDefs("tblSystemRules")
        t.Fields("ExitOnEntryBar").OrdinalPosition = 9
        t.Fields("LinkedRules").OrdinalPosition = 10
        
        Set t = m.DB.TableDefs("tblFunctions")
        t.Fields("Formatted").OrdinalPosition = 8
        t.Fields("FormattedWithFillWords").OrdinalPosition = 9
        t.Fields("CodedName").OrdinalPosition = 10
        t.Fields("CodedText").OrdinalPosition = 11
        t.Fields("Usage").OrdinalPosition = 12
        t.Fields("Password").OrdinalPosition = 18
        t.Fields("LastModified").OrdinalPosition = 19
           
        Set t = m.DB.TableDefs("tblLibrarys")
        t.Fields("Password").OrdinalPosition = 10
        t.Fields("rtfFileName").OrdinalPosition = 11
    
        Set t = m.DB.TableDefs("tblSystems")
        t.Fields("Password").OrdinalPosition = 14
        t.Fields("LastModified").OrdinalPosition = 15
    
        Set t = m.DB.TableDefs("tblRules")
        t.Fields("PreviewRTF").OrdinalPosition = 3
        t.Fields("CondRTF").OrdinalPosition = 4
        t.Fields("CondCoded").OrdinalPosition = 5
        t.Fields("PriceRTF").OrdinalPosition = 6
        t.Fields("PriceCoded").OrdinalPosition = 7
        t.Fields("Price2RTF").OrdinalPosition = 8
        t.Fields("Price2Coded").OrdinalPosition = 9
        t.Fields("OrderPlacement").OrdinalPosition = 10
        t.Fields("Password").OrdinalPosition = 18
        t.Fields("LastModified").OrdinalPosition = 19
    
    Case 35:
        ' add new function categories
        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctionCategories];", dbOpenDynaset)
        rs.MoveLast
        lIndex = rs!FunctionCategoryID
        If lIndex < 23 Then
            rs.AddNew
            rs!FunctionCategory = "Planet Combo"
            rs!CannotDelete = True
            rs.Update
        End If
        If lIndex < 24 Then
            rs.AddNew
            rs!FunctionCategory = "Candlesticks"
            rs!CannotDelete = True
            rs.Update
        End If
        
        Set t = m.DB.TableDefs("tblFunctions")
        With t
            If ItemExists(.Fields, "RequiredData") Then
                .Fields.Delete "RequiredData"
            End If
            If Not ItemExists(.Fields, "RequiredMod") Then
                .Fields.Append .CreateField("RequiredMod", dbMemo)
                .Fields("RequiredMod").Required = False
                .Fields("RequiredMod").AllowZeroLength = True
            End If
        End With
        
        strName = "qryPackageFunctions"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblFunctions.FunctionName, tblFunctions.Description, tblFunctions.FunctionCategoryID, tblFunctions.ImplementationTypeID, tblFunctions.ReturnTypeID, tblFunctions.TradeSenseUsage, tblFunctions.Formatted, tblFunctions.FormattedWithFillWords, tblFunctions.CodedName, tblFunctions.CodedText, tblFunctions.Usage, tblFunctions.LateCalculating, tblFunctions.UsesOpenNextBar, tblFunctions.UsesMMData, tblFunctions.SecurityLevel, tblFunctions.CannotDelete, tblFunctions.Password, tblFunctions.LastModified, tblFunctions.LibraryID, tblFunctions.Reverify, tblFunctions.Abbrev, tblFunctions.UsesHLCNextBar, tblFunctions.ListID, tblFunctions.CheckSum, tblFunctions.RequiredMod " & _
                "FROM tblFunctions " & _
                "WHERE (((tblFunctions.LibraryID) = [SelLibraryID])) " & _
                "ORDER BY tblFunctions.CodedName;"
    
    Case 36:
        ' add new function categories
        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctionCategories];", dbOpenDynaset)
        rs.MoveLast
        lIndex = rs!FunctionCategoryID
        If lIndex < 25 Then
            rs.AddNew
            rs!FunctionCategory = "Scoring"
            rs!CannotDelete = True
            rs.Update
        End If
        If lIndex < 26 Then
            rs.AddNew
            rs!FunctionCategory = "Fundamentals"
            rs!CannotDelete = True
            rs.Update
        End If
        
    Case 37:
        Set t = m.DB.TableDefs("tblSystems")
        With t
            If Not ItemExists(.Fields, "AllowReverse") Then
                .Fields.Append .CreateField("AllowReverse", dbBoolean)
                .Fields("AllowReverse").DefaultValue = True
            
                Set q = m.DB.CreateQueryDef("") 'temp update query
                q.SQL = "UPDATE tblSystems SET tblSystems.AllowReverse = Yes;"
                q.Execute
            End If
        End With
        
        strName = "qryPackageSystems"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblSystems.SystemName, tblSystems.BarTimeFrame, tblSystems.Expenses, tblSystems.FromDate, tblSystems.ToDate, tblSystems.ToEndOfData, tblSystems.LibraryID, tblSystems.Developer, tblSystems.TradesPath, tblSystems.Notes, tblSystems.SecurityLevel, tblSystems.CannotDelete, tblSystems.Password, tblSystems.LastModified, tblSystems.BarsLoadedBeforeTrading, tblSystems.BarsTradedBeforeOrders, tblSystems.Reverify, tblSystems.MMID, tblSystems.Pyramid, tblSystems.TradeDepth, tblSystems.CanPyramidSameEntry, tblSystems.LinkInputs, tblSystems.CheckSum, tblSystems.NumShares, tblSystems.StockExpenses, tblSystems.AllowReverse " & _
                "FROM tblSystems " & _
                "WHERE (((tblSystems.LibraryID)=[SelLibraryID]));"
                
    Case 38:
        Set t = m.DB.TableDefs("tblSystemSecurities")
        With t
            If Not ItemExists(.Fields, "GroupID") Then
                .Fields.Append .CreateField("GroupID", dbText, 20)
                .Fields("GroupID").AllowZeroLength = True
                .Fields("GroupID").Required = False
                .Fields("GroupID").DefaultValue = ""
            
                Set q = m.DB.CreateQueryDef("") 'temp update query
                q.SQL = "UPDATE tblSystemSecurities SET tblSystemSecurities.GroupID = '';"
                q.Execute
            End If
        End With
    
        strName = "qryPackageSystemSecurities"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblSystems.SystemName, tblRules.Name AS RuleName, tblRuleParms.ParmName, tblSystemSecurities.Path, tblSystemSecurities.Symbol, tblSystemSecurities.MarketSymbol, tblSystemSecurities.Periodicity, tblSystemSecurities.Format, tblSystemSecurities.SecurityType, tblSystemSecurities.SecurityName,tblSystemSecurities.GroupID " & _
                "FROM tblSystems INNER JOIN (tblRules INNER JOIN (tblRuleParms INNER JOIN tblSystemSecurities ON tblRuleParms.ParmID = tblSystemSecurities.ParmID) ON tblRules.RuleID = tblRuleParms.RuleID) ON tblSystems.SystemNumber = tblSystemSecurities.SystemNumber " & _
                "WHERE (((tblSystems.LibraryID)=[SelLibraryID]));"

        strName = "qrySystemSecurities"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblSystemSecurities.SystemNumber, tblSystemSecurities.ParmID, tblRuleParms.RuleID, tblRuleParms.ParmName, tblSystemSecurities.Path, tblSystemSecurities.Symbol, tblSystemSecurities.MarketSymbol, tblSystemSecurities.Periodicity, tblSystemSecurities.Format, tblSystemSecurities.SecurityType, tblSystemSecurities.SecurityName,tblSystemSecurities.GroupID " & _
                "FROM tblRuleParms INNER JOIN tblSystemSecurities ON tblRuleParms.ParmID = tblSystemSecurities.ParmID " & _
                "WHERE (((tblSystemSecurities.SystemNumber) = [SelSystemNumber])) " & _
                "ORDER BY tblSystemSecurities.SystemNumber;"
    
        strName = "qrySystemSecurityByParmID"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblSystemSecurities.SystemNumber, tblSystemSecurities.ParmID, tblSystemSecurities.Path, tblSystemSecurities.Symbol, tblSystemSecurities.MarketSymbol, tblSystemSecurities.Periodicity, tblSystemSecurities.Format, tblSystemSecurities.SecurityType, tblSystemSecurities.SecurityName, tblSystemSecurities.GroupID, tblRuleParms.ParmName, tblRuleParms.RuleID " & _
                "FROM tblRuleParms INNER JOIN tblSystemSecurities ON tblRuleParms.ParmID = tblSystemSecurities.ParmID " & _
                "WHERE (((tblSystemSecurities.SystemNumber) = [SelSystemNumber]) And ((tblSystemSecurities.ParmID) = [SelParmID])) " & _
                "ORDER BY tblSystemSecurities.SystemNumber, tblSystemSecurities.ParmID;"

    Case 39:
        ChangeDataType "tblLibrarys", "RequiredMod", dbMemo
        
    Case 40:
        AddRuleCategories
        
    Case 41:
        Set t = m.DB.TableDefs("tblSystemSecurities")
        With t
            If Not ItemExists(.Fields, "SymbolID") Then
                .Fields.Append .CreateField("SymbolID", dbLong)
                .Fields("SymbolID").DefaultValue = 0
            
                Set q = m.DB.CreateQueryDef("") 'temp update query
                q.SQL = "UPDATE tblSystemSecurities SET tblSystemSecurities.SymbolID = 0;"
                q.Execute
            End If
        End With

        strName = "qryPackageSystemSecurities"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblSystems.SystemName, tblRules.Name AS RuleName, tblRuleParms.ParmName, tblSystemSecurities.Path, tblSystemSecurities.Symbol, tblSystemSecurities.MarketSymbol, tblSystemSecurities.Periodicity, tblSystemSecurities.Format, tblSystemSecurities.SecurityType, tblSystemSecurities.SecurityName, tblSystemSecurities.GroupID, tblSystemSecurities.SymbolID " & _
                "FROM tblSystems INNER JOIN (tblRules INNER JOIN (tblRuleParms INNER JOIN tblSystemSecurities ON tblRuleParms.ParmID = tblSystemSecurities.ParmID) ON tblRules.RuleID = tblRuleParms.RuleID) ON tblSystems.SystemNumber = tblSystemSecurities.SystemNumber " & _
                "WHERE (((tblSystems.LibraryID)=[SelLibraryID]));"

        strName = "qrySystemSecurities"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblSystemSecurities.SystemNumber, tblSystemSecurities.ParmID, tblRuleParms.RuleID, tblRuleParms.ParmName, tblSystemSecurities.Path, tblSystemSecurities.Symbol, tblSystemSecurities.MarketSymbol, tblSystemSecurities.Periodicity, tblSystemSecurities.Format, tblSystemSecurities.SecurityType, tblSystemSecurities.SecurityName,tblSystemSecurities.GroupID, tblSystemSecurities.SymbolID " & _
                "FROM tblRuleParms INNER JOIN tblSystemSecurities ON tblRuleParms.ParmID = tblSystemSecurities.ParmID " & _
                "WHERE (((tblSystemSecurities.SystemNumber) = [SelSystemNumber])) " & _
                "ORDER BY tblSystemSecurities.SystemNumber;"
    
        strName = "qrySystemSecurityByParmID"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblSystemSecurities.SystemNumber, tblSystemSecurities.ParmID, tblSystemSecurities.Path, tblSystemSecurities.Symbol, tblSystemSecurities.MarketSymbol, tblSystemSecurities.Periodicity, tblSystemSecurities.Format, tblSystemSecurities.SecurityType, tblSystemSecurities.SecurityName, tblSystemSecurities.GroupID, tblSystemSecurities.SymbolID, tblRuleParms.ParmName, tblRuleParms.RuleID " & _
                "FROM tblRuleParms INNER JOIN tblSystemSecurities ON tblRuleParms.ParmID = tblSystemSecurities.ParmID " & _
                "WHERE (((tblSystemSecurities.SystemNumber) = [SelSystemNumber]) And ((tblSystemSecurities.ParmID) = [SelParmID])) " & _
                "ORDER BY tblSystemSecurities.SystemNumber, tblSystemSecurities.ParmID;"
    
    Case 42:
        AddTradingItemTable
        
    Case 43:
        Set t = m.DB.TableDefs("tblAutoTradingItem")
        With t
            If Not ItemExists(.Fields, "Name") Then
                .Fields.Append .CreateField("Name", dbText, 50)
                .Fields("Name").AllowZeroLength = True
                .Fields("Name").Required = False
                .Fields("Name").DefaultValue = ""
            End If
            
            If Not ItemExists(.Fields, "ConfirmTimeout") Then
                .Fields.Append .CreateField("ConfirmTimeout", dbLong)
                .Fields("ConfirmTimeout").DefaultValue = 15&
            End If
            
            Set rs = m.DB.OpenRecordset("SELECT * FROM [tblAutoTradingItem];", dbOpenDynaset)
            lIndex = 1
            Do While Not rs.EOF
                rs.Edit
                If IsNull(rs!Name) Then
                    rs!Name = "Trading Item #" & Format(lIndex, "000")
                    lIndex = lIndex + 1
                End If
                
                If IsNull(rs!ConfirmTimeout) Then rs!ConfirmTimeout = 15&
                rs.Update
                
                rs.MoveNext
            Loop
        End With
        
    Case 44:
        ' Add the Spread Function Category...
        Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblFunctionCategories];", dbOpenDynaset)
        rs.MoveLast
        lIndex = rs!FunctionCategoryID
        If lIndex < 27 Then
            rs.AddNew
            rs!FunctionCategory = "Spread"
            rs!CannotDelete = True
            rs.Update
        End If
        
        ' Add the "Force Price to go through Limit Price" option to strategies...
        Set t = m.DB.TableDefs("tblSystems")
        With t
            If Not ItemExists(.Fields, "ForceLimitThrough") Then
                .Fields.Append .CreateField("ForceLimitThrough", dbBoolean)
                .Fields("ForceLimitThrough").DefaultValue = True
            
                Set q = m.DB.CreateQueryDef("") 'temp update query
                q.SQL = "UPDATE tblSystems SET tblSystems.ForceLimitThrough = Yes;"
                q.Execute
            End If
        End With
        
        strName = "qryPackageSystems"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblSystems.SystemName, tblSystems.BarTimeFrame, tblSystems.Expenses, tblSystems.FromDate, tblSystems.ToDate, tblSystems.ToEndOfData, tblSystems.LibraryID, tblSystems.Developer, tblSystems.TradesPath, tblSystems.Notes, tblSystems.SecurityLevel, tblSystems.CannotDelete, tblSystems.Password, tblSystems.LastModified, tblSystems.BarsLoadedBeforeTrading, tblSystems.BarsTradedBeforeOrders, tblSystems.Reverify, tblSystems.MMID, tblSystems.Pyramid, tblSystems.TradeDepth, tblSystems.CanPyramidSameEntry, tblSystems.LinkInputs, tblSystems.CheckSum, tblSystems.NumShares, tblSystems.StockExpenses, tblSystems.AllowReverse, tblSystems.ForceLimitThrough " & _
                "FROM tblSystems " & _
                "WHERE (((tblSystems.LibraryID)=[SelLibraryID]));"
                
    Case 45:
        Set rs = m.DB.OpenRecordset("SELECT * FROM [tblMarketInfo] " & _
                "WHERE [Symbol]='$GBP-USD' AND [DataType]=3 AND [Value]='17.65';", dbOpenDynaset)
        Do While Not rs.EOF
            rs.Delete
            rs.MoveNext
        Loop
                        
    Case 46:
        ' Add the "Dollars per Trade" option to strategies...
        Set t = m.DB.TableDefs("tblSystems")
        With t
            If Not ItemExists(.Fields, "DollarsPerTrade") Then
                .Fields.Append .CreateField("DollarsPerTrade", dbCurrency)
                .Fields("DollarsPerTrade").DefaultValue = 10000#
            
                Set q = m.DB.CreateQueryDef("") 'temp update query
                q.SQL = "UPDATE tblSystems SET tblSystems.DollarsPerTrade = 10000;"
                q.Execute
            End If
            If Not ItemExists(.Fields, "UseSharesPerTrade") Then
                .Fields.Append .CreateField("UseSharesPerTrade", dbBoolean)
                .Fields("UseSharesPerTrade").DefaultValue = False
            
                'Set q = m.DB.CreateQueryDef("") 'temp update query
                'q.SQL = "UPDATE tblSystems SET tblSystems.UseSharesPerTrade = No;"
                'q.Execute
            End If
        End With
        
        strName = "qryPackageSystems"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblSystems.SystemName, tblSystems.BarTimeFrame, tblSystems.Expenses, tblSystems.FromDate, tblSystems.ToDate, tblSystems.ToEndOfData, tblSystems.LibraryID, tblSystems.Developer, tblSystems.TradesPath, tblSystems.Notes, tblSystems.SecurityLevel, tblSystems.CannotDelete, tblSystems.Password, tblSystems.LastModified, tblSystems.BarsLoadedBeforeTrading, tblSystems.BarsTradedBeforeOrders, tblSystems.Reverify, tblSystems.MMID, tblSystems.Pyramid, tblSystems.TradeDepth, tblSystems.CanPyramidSameEntry, tblSystems.LinkInputs, tblSystems.CheckSum, tblSystems.NumShares, tblSystems.StockExpenses, tblSystems.AllowReverse, tblSystems.ForceLimitThrough, tblSystems.DollarsPerTrade, tblSystems.UseSharesPerTrade " & _
                "FROM tblSystems " & _
                "WHERE (((tblSystems.LibraryID)=[SelLibraryID]));"
                
    Case 47:
        ' Add the "Forex Expenses" to strategies...
        Set t = m.DB.TableDefs("tblSystems")
        With t
            If Not ItemExists(.Fields, "ForexExpenses") Then
                .Fields.Append .CreateField("ForexExpenses", dbCurrency)
                .Fields("ForexExpenses").DefaultValue = 5#
            
                Set q = m.DB.CreateQueryDef("") 'temp update query
                q.SQL = "UPDATE tblSystems SET tblSystems.ForexExpenses = 5;"
                q.Execute
            End If
        End With
        
        strName = "qryPackageSystems"
        If ItemExists(m.DB.QueryDefs, strName) Then
            m.DB.QueryDefs.Delete strName
        End If
        Set q = m.DB.CreateQueryDef(strName)
        q.SQL = "SELECT tblSystems.SystemName, tblSystems.BarTimeFrame, tblSystems.Expenses, tblSystems.FromDate, tblSystems.ToDate, tblSystems.ToEndOfData, tblSystems.LibraryID, tblSystems.Developer, tblSystems.TradesPath, tblSystems.Notes, tblSystems.SecurityLevel, tblSystems.CannotDelete, tblSystems.Password, tblSystems.LastModified, tblSystems.BarsLoadedBeforeTrading, tblSystems.BarsTradedBeforeOrders, tblSystems.Reverify, tblSystems.MMID, tblSystems.Pyramid, tblSystems.TradeDepth, tblSystems.CanPyramidSameEntry, tblSystems.LinkInputs, tblSystems.CheckSum, tblSystems.NumShares, tblSystems.StockExpenses, tblSystems.AllowReverse, tblSystems.ForceLimitThrough, tblSystems.DollarsPerTrade, tblSystems.UseSharesPerTrade, tblSystems.ForexExpenses " & _
                "FROM tblSystems " & _
                "WHERE (((tblSystems.LibraryID)=[SelLibraryID]));"
                
    Case 48:    ' Add AccountPositionID field to tblAutoTradeItem
        ApplyVersion48
        
    Case 49:    ' Add Deleted field to tblAutoTradeItem
        ApplyVersion49
        
    Case 50:    ' Add OnCloseTimeExch field to tblAutoTradeItem
        ApplyVersion50
        
    Case 51:    ' Bring strategy baskets into the database
        ApplyVersion51
    
    Case 52:    ' Extra fields in tblAutoTradeItem for basket information
        ApplyVersion52
        
    Case 53:    ' Add "Guru" security flag to strategies and baskets
        ApplyVersion53
        
    Case 54:    ' Fix the qryPackageSystems that got screwed up in version 53
        ApplyVersion54
        
    Case 55:    ' Add "Override" and "Expose" flags to strategy basket item parameters
        ApplyVersion55
        
    Case 56:    ' Add "ExitAtEndOfDay" flag to automated trading items
        ApplyVersion56
        
    End Select
    
    Set rs2 = Nothing
    Set rs = Nothing
    Set q = Nothing
    Set i = Nothing
    Set r = Nothing
    Set t = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cDatabaseUpdates.ApplyVersion", eGDRaiseError_Raise

End Sub

'Changes required for New TradeSense merge
Private Sub Apply11()
On Error GoTo ErrSection:
    
    Dim t           As TableDef
    Dim i           As Index
    Dim r           As Relation
    Dim q           As QueryDef

    '========================== tblFunctions ==============================
    'Add LISTID field to tblFunctions
    Set t = m.DB.TableDefs.Item("tblFunctions")
    With t
        .Fields.Append .CreateField("ListID", dbLong)
    End With
    
    'Set ListID for all functions to zero
    Set q = m.DB.CreateQueryDef("") 'temp update query
    q.SQL = "UPDATE tblFunctions SET tblFunctions.ListID = 0;"
    q.Execute

    'Update SQL in: qryFunctionsAndParms
    Set q = m.DB.QueryDefs("qryFunctionsAndParms")
    q.SQL = "SELECT tblFunctions.FunctionID, tblFunctions.FunctionName, tblFunctions.Description, tblFunctions.FunctionCategoryID," & _
            "tblFunctions.ImplementationTypeID, tblFunctions.ReturnTypeID, tblFunctions.CodedName, tblFunctions.Formatted," & _
            "tblFunctions.FormattedWithFillWords, tblFunctions.CodedText, tblFunctions.Usage, tblFunctions.LateCalculating, " & _
            "tblFunctions.UsesOpenNextBar, tblFunctions.UsesMMData, tblFunctions.SecurityLevel, tblFunctions.CannotDelete," & _
            "tblFunctions.Password, tblFunctions.LastModified, tblFunctions.LibraryID, tblFunctions.TradeSenseUsage, " & _
            "tblFunctionParms.ParmID, tblFunctionParms.ParmNbr, tblFunctionParms.ParmText, tblFunctionParms.ParmDesc, " & _
            "tblFunctionParms.ParmTypeID, tblFunctionParms.DefaultValue, tblFunctionParms.FromValue, tblFunctionParms.ToValue," & _
            "tblFunctionParms.Required, tblFunctionParms.ListID, tblFunctionParms.FillPre, tblFunctionParms.FillPost, tblFunctions.Reverify," & _
            "tblFunctions.Abbrev, tblFunctions.UsesHLCNextBar, tblFunctionCategories.FunctionCategory, tblFunctions.ListID AS FListID," & _
            "tblParmTypes.ParmUsage AS IParmUsage " & _
        "FROM tblParmTypes RIGHT JOIN ((tblFunctionCategories INNER JOIN tblFunctions ON tblFunctionCategories.FunctionCategoryID " & _
            "= tblFunctions.FunctionCategoryID) LEFT JOIN tblFunctionParms ON tblFunctions.FunctionID = tblFunctionParms.FunctionID) ON " & _
            "tblParmTypes.ParmTypeID = tblFunctionParms.ParmTypeID " & _
        "ORDER BY tblFunctions.FunctionName, tblFunctionParms.ParmNbr;"
            
    'Update SQL in: qryEngineFunction
    Set q = m.DB.QueryDefs("qryEngineFunction")
    q.SQL = "SELECT tblFunctions.FunctionID, tblFunctions.FunctionName, tblFunctions.ImplementationTypeID, tblFunctions.ReturnTypeID, " & _
            "tblFunctions.CodedName, tblFunctions.CodedText, tblFunctions.Usage, tblFunctions.LateCalculating, tblFunctions.UsesOpenNextBar, " & _
            "tblFunctions.UsesMMData, tblFunctions.LastModified, tblFunctions.LibraryID, tblFunctions.UsesHLCNextBar, tblFunctions.ListID " & _
        "From tblFunctions " & _
        "Where (((tblFunctions.LibraryID) = [Enter Library ID])) " & _
        "ORDER BY tblFunctions.FunctionName;"
            
    'Update SQL in: qryPackageFunctions
    Set q = m.DB.QueryDefs("qryPackageFunctions")
    q.SQL = "SELECT tblFunctions.FunctionName, tblFunctions.Description, tblFunctions.FunctionCategoryID, tblFunctions.ImplementationTypeID, " & _
                "tblFunctions.ReturnTypeID, tblFunctions.TradeSenseUsage, tblFunctions.Formatted, tblFunctions.FormattedWithFillWords, tblFunctions.CodedName, " & _
                "tblFunctions.CodedText, tblFunctions.Usage, tblFunctions.LateCalculating, tblFunctions.UsesOpenNextBar, tblFunctions.UsesMMData, " & _
                "tblFunctions.SecurityLevel, tblFunctions.CannotDelete, tblFunctions.Password, tblFunctions.LastModified, tblFunctions.LibraryID, " & _
                "tblFunctions.Reverify, tblFunctions.Abbrev, tblFunctions.UsesHLCNextBar, tblFunctions.ListID " & _
            "From tblFunctions " & _
            "Where (((tblFunctions.LibraryID) = [SelLibraryID])) " & _
            "ORDER BY tblFunctions.CodedName;"

    
    '========================== tblLists ============================
    'Change 'Markets' to 'Symbol Names'.  TradeSense keys on Name here...
    Set q = m.DB.CreateQueryDef("") 'temp update query
    q.SQL = "UPDATE tblLists SET tblLists.ListName = 'Symbol Names' " & _
        "WHERE (((tblLists.ListID)=9));"
    q.Execute

    'Delete existing Order types...
    Set q = m.DB.CreateQueryDef("") 'temp update query
    q.SQL = "DELETE tblListItems.ListID From tblListItems " & _
        "WHERE (((tblListItems.ListID)=3));"

    'Insert new Order types...
    '1=Market, 2=Stop, 3=Limit, 4=Stop with Limit, 5=Market on Close,
    '6=Stop close only, 7=Limit close only, 8=Stop with Limit close only
    Set q = m.DB.CreateQueryDef("") 'temp update query
    q.SQL = "INSERT INTO tblListItems ( ListID, Seq, ListItem ) " & _
        "SELECT 3 AS Expr1, 1 AS Expr2, 'Market' AS Expr3;"
    q.Execute
    q.SQL = "INSERT INTO tblListItems ( ListID, Seq, ListItem ) " & _
        "SELECT 3 AS Expr1, 2 AS Expr2, 'Stop' AS Expr3;"
    q.Execute
    q.SQL = "INSERT INTO tblListItems ( ListID, Seq, ListItem ) " & _
        "SELECT 3 AS Expr1, 3 AS Expr2, 'Limit' AS Expr3;"
    q.Execute
    q.SQL = "INSERT INTO tblListItems ( ListID, Seq, ListItem ) " & _
        "SELECT 3 AS Expr1, 4 AS Expr2, 'Stop with Limit' AS Expr3;"
    q.Execute
    q.SQL = "INSERT INTO tblListItems ( ListID, Seq, ListItem ) " & _
        "SELECT 3 AS Expr1, 5 AS Expr2, 'Market on Close' AS Expr3;"
    q.Execute
    q.SQL = "INSERT INTO tblListItems ( ListID, Seq, ListItem ) " & _
        "SELECT 3 AS Expr1, 6 AS Expr2, 'Stop close only' AS Expr3;"
    q.Execute
    q.SQL = "INSERT INTO tblListItems ( ListID, Seq, ListItem ) " & _
        "SELECT 3 AS Expr1, 7 AS Expr2, 'Limit close only' AS Expr3;"
    q.Execute
    q.SQL = "INSERT INTO tblListItems ( ListID, Seq, ListItem ) " & _
        "SELECT 3 AS Expr1, 8 AS Expr2, 'Stop with Limit close only' AS Expr3;"
    q.Execute
    
    '========================== tblParmTypes ============================
    'Insert new Parm type for Series of text values...
    q.SQL = "INSERT INTO tblParmTypes ( ParmTypeID, ParmType, ParmTypeDesc, ParmUsage ) " & _
        "SELECT 8 AS Expr1, 'Series of Text values' AS Expr2, 'Series of Text values' AS Expr3, 2 AS Expr4;"
    q.Execute

    'Update the Parm usage column for all other parm types.
    q.SQL = "UPDATE tblParmTypes SET tblParmTypes.ParmUsage = 2 " & _
        "WHERE (((tblParmTypes.ParmTypeID)=1));"
    q.Execute
    q.SQL = "UPDATE tblParmTypes SET tblParmTypes.ParmUsage = 2 " & _
        "WHERE (((tblParmTypes.ParmTypeID)=2));"
    q.Execute
    q.SQL = "UPDATE tblParmTypes SET tblParmTypes.ParmUsage = 2 " & _
        "WHERE (((tblParmTypes.ParmTypeID)=3));"
    q.Execute
    q.SQL = "UPDATE tblParmTypes SET tblParmTypes.ParmUsage = 2 " & _
        "WHERE (((tblParmTypes.ParmTypeID)=4));"
    q.Execute
    q.SQL = "UPDATE tblParmTypes SET tblParmTypes.ParmUsage = 3 " & _
        "WHERE (((tblParmTypes.ParmTypeID)=5));"
    q.Execute
    q.SQL = "UPDATE tblParmTypes SET tblParmTypes.ParmUsage = 2 " & _
        "WHERE (((tblParmTypes.ParmTypeID)=6));"
    q.Execute
    q.SQL = "UPDATE tblParmTypes SET tblParmTypes.ParmUsage = 3 " & _
        "WHERE (((tblParmTypes.ParmTypeID)=7));"
    q.Execute
    q.SQL = "UPDATE tblParmTypes SET tblParmTypes.ParmUsage = 2 " & _
        "WHERE (((tblParmTypes.ParmTypeID)=8));"
    q.Execute
    q.SQL = "UPDATE tblParmTypes SET tblParmTypes.ParmUsage = 3 " & _
        "WHERE (((tblParmTypes.ParmTypeID)=14));"
    q.Execute
    
    ' Make sure that anything late calculating is an exit rule
    q.SQL = "UPDATE tblRules SET tblRules.RuleType = 1 " & _
        "WHERE tblRules.LateCondition = Yes OR tblRules.LateAction = Yes;"
    q.Execute
    
    Set q = Nothing
    Set r = Nothing
    Set i = Nothing
    Set t = Nothing

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cDatabaseUpdates.Apply11", eGDRaiseError_Raise

End Sub

'Delete this when RuleMgt change is complete
'This routine converts the global flag to yes for rules used across systems.
Public Sub ConvertRules()
On Error GoTo ErrSection:

    Dim rsSysRules  As Recordset
    Dim rsRules     As Recordset
    Dim strMDB      As String
    
'    Set g.WrkJet = CreateWorkspace("", "admin", "", dbUseJet)
'    strMDB = App.Path & "\Libraries.MDB"
'    Set g.dbNav = g.WrkJet.OpenDatabase(strMDB, False, False, "; pwd=" & DbPassword)
'    If Not LinkTableToDb(g.dbNav, "tblMarkets", App.Path & "\Market.mdb") Then
'        g.dbNav.Close
'        End
'    End If

    Set rsSysRules = g.dbNav.OpenRecordset("SELECT tblSystemRules.RuleID, Count(tblSystemRules.Selected) AS Cnt " & _
        "FROM tblSystemRules " & _
        "GROUP BY tblSystemRules.RuleID " & _
        "HAVING (((Count(tblSystemRules.Selected))>1))", dbOpenSnapshot)
    Set rsRules = g.dbNav.OpenRecordset("Select * from tblRules", dbOpenDynaset)

    Do Until rsRules.EOF
        rsSysRules.FindFirst "[RuleID]=" & rsRules!RuleID
        If Not rsSysRules.NoMatch Then
            rsRules.Edit
            rsRules!SystemNumber = 0
            rsRules.Update
        End If
        rsRules.MoveNext
    Loop
    
    Set rsSysRules = g.dbNav.OpenRecordset("SELECT tblSystemRules.RuleID, First(tblSystemRules.SystemNumber) AS SystemNumber, Count(tblSystemRules.Selected) AS Cnt " & _
        "FROM tblSystemRules " & _
        "GROUP BY tblSystemRules.RuleID " & _
        "HAVING (((Count(tblSystemRules.Selected)) = 1)) " & _
        "ORDER BY tblSystemRules.RuleID;", dbOpenSnapshot)
    Set rsRules = g.dbNav.OpenRecordset("Select * from tblRules", dbOpenDynaset)
    
    Do Until rsRules.EOF
        rsSysRules.FindFirst "[RuleID]=" & rsRules!RuleID
        If Not rsSysRules.NoMatch Then
            If rsRules!LibraryID <> 1 Then
                rsRules.Edit
                rsRules!SystemNumber = rsSysRules!SystemNumber
                rsRules.Update
            End If
        End If
        rsRules.MoveNext
    Loop
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cDatabaseUpdates.ConvertRules", eGDRaiseError_Raise

End Sub

Private Sub FixRulePreview()
On Error GoTo ErrSection:

    Dim i&, strText$, strFind$, strNew$
    Dim rs As Recordset
       
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [tblRules];", dbOpenDynaset)
    Do While Not rs.EOF
        strNew = ""
        strText = rs!PreviewRTF
        strFind = "~24007BUY at"
        i = InStr(strText, strFind)
        If i > 0 Then
            strNew = Left(strText, i - 1) & "~24016BUY Next Bar at" & Mid(strText, i + Len(strFind))
        Else
            strFind = "~24008SELL at"
            i = InStr(strText, strFind)
            If i > 0 Then
                strNew = Left(strText, i - 1) & "~24017SELL Next Bar at" & Mid(strText, i + Len(strFind))
            End If
        End If
        If Len(strNew) > 0 Then
            rs.Edit
            rs!PreviewRTF = strNew
            rs.Update
        End If
        rs.MoveNext
    Loop
    
ErrExit:
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    RaiseError "cDatabaseUpdates.FixRulePreview", eGDRaiseError_Raise

End Sub

Private Sub Apply25()
On Error GoTo ErrSection:

    Dim t As TableDef
    Dim i As Index
    Dim r As Relation
    Dim q As QueryDef
    Dim rs As Recordset
    Dim strName As String

    ' Add CustomerID and CheckSum fields to the Database table
    Set t = m.DB.TableDefs("tblDatabase")
    With t
        If Not ItemExists(.Fields, "CID") Then
            .Fields.Append .CreateField("CID", dbLong)
            With .Fields("CID")
                .DefaultValue = 0
                .Required = True
            End With
        End If
        If Not ItemExists(.Fields, "CheckSum") Then
            .Fields.Append .CreateField("CheckSum", dbDouble)
            With .Fields("CheckSum")
                .DefaultValue = 0
                .Required = True
            End With
        End If
        
        ' Fill in the Last Known Customer ID and calculate the CheckSum
        Set rs = m.DB.OpenRecordset("SELECT * FROM [tblDatabase];", dbOpenDynaset)
        With rs
            .Edit
            !CID = g.lLCD ' RI_GetLastDataServiceID \ 1000&
            !CheckSum = BuildCheckSum(rs, "tblDatabase")
            .Update
        End With
    End With
    
    ' Add a CheckSum field to the Librarys table
    Set t = m.DB.TableDefs("tblLibrarys")
    With t
        If Not ItemExists(.Fields, "CheckSum") Then
            .Fields.Append .CreateField("CheckSum", dbDouble)
            With .Fields("CheckSum")
                .DefaultValue = 0
                .Required = True
            End With
        End If
    
        ' Encrypt appropriate fields and calculate a checksum for all records
        Set rs = m.DB.OpenRecordset("SELECT * FROM [tblLibrarys];", dbOpenDynaset)
        With rs
            Do While Not .EOF
                .Edit
                EncryptField !Password, NullChk(!Password)
                EncryptField !RequiredMod, NullChk(!RequiredMod)
                !CheckSum = BuildCheckSum(rs, "tblLibrarys")
                .Update
                
                .MoveNext
            Loop
        End With
    End With
    
    ' Add a CheckSum field to the Systems table
    Set t = m.DB.TableDefs("tblSystems")
    With t
        If Not ItemExists(.Fields, "CheckSum") Then
            .Fields.Append .CreateField("CheckSum", dbDouble)
            With .Fields("CheckSum")
                .DefaultValue = 0
                .Required = True
            End With
        End If
    
        ' Encrypt appropriate fields and calculate a checksum for all records
        Set rs = m.DB.OpenRecordset("SELECT * FROM [tblSystems];", dbOpenDynaset)
        With rs
            Do While Not .EOF
                .Edit
                EncryptField !Password, NullChk(!Password)
                !CheckSum = BuildCheckSum(rs, "tblSystems")
                .Update
                
                .MoveNext
            Loop
        End With
    End With
    
    ' Add a CheckSum field to the Rules table
    Set t = m.DB.TableDefs("tblRules")
    With t
        If Not ItemExists(.Fields, "CheckSum") Then
            .Fields.Append .CreateField("CheckSum", dbDouble)
            With .Fields("CheckSum")
                .DefaultValue = 0
                .Required = True
            End With
        End If
    
        ' Encrypt appropriate fields and calculate a checksum for all records
        Set rs = m.DB.OpenRecordset("SELECT * FROM [tblRules];", dbOpenDynaset)
        With rs
            Do While Not .EOF
                .Edit
                EncryptField !PreviewRTF, NullChk(!PreviewRTF)
                EncryptField !CondRTF, NullChk(!CondRTF)
                EncryptField !CondCoded, NullChk(!CondCoded)
                EncryptField !PriceRTF, NullChk(!PriceRTF)
                EncryptField !PriceCoded, NullChk(!PriceCoded)
                EncryptField !Price2RTF, NullChk(!Price2RTF)
                EncryptField !Price2Coded, NullChk(!Price2Coded)
                EncryptField !Password, NullChk(!Password)
                !CheckSum = BuildCheckSum(rs, "tblRules")
                .Update
                
                .MoveNext
            Loop
        End With
    End With
    
    ' Add a CheckSum field to the Functions table
    Set t = m.DB.TableDefs("tblFunctions")
    With t
        If Not ItemExists(.Fields, "CheckSum") Then
            .Fields.Append .CreateField("CheckSum", dbDouble)
            With .Fields("CheckSum")
                .DefaultValue = 0
                .Required = True
            End With
        End If
    
        ' Encrypt appropriate fields and calculate a checksum for all records
        Set rs = m.DB.OpenRecordset("SELECT * FROM [tblFunctions];", dbOpenDynaset)
        With rs
            Do While Not .EOF
                .Edit
                EncryptField !Formatted, NullChk(!Formatted)
                EncryptField !FormattedWithFillWords, NullChk(!FormattedWithFillWords)
                EncryptField !CodedText, NullChk(!CodedText)
                EncryptField !Password, NullChk(!Password)
                !CheckSum = BuildCheckSum(rs, "tblFunctions")
                .Update
                
                .MoveNext
            Loop
        End With
    End With
    
    ' Add a CheckSum field to the SystemRules table
    Set t = m.DB.TableDefs("tblSystemRules")
    With t
        If Not ItemExists(.Fields, "CheckSum") Then
            .Fields.Append .CreateField("CheckSum", dbDouble)
            With .Fields("CheckSum")
                .DefaultValue = 0
                .Required = True
            End With
        End If
    
        ' Encrypt appropriate fields and calculate a checksum for all records
        Set rs = m.DB.OpenRecordset("SELECT * FROM [tblSystemRules];", dbOpenDynaset)
        With rs
            Do While Not .EOF
                .Edit
                !CheckSum = BuildCheckSum(rs, "tblSystemRules")
                .Update
                
                .MoveNext
            Loop
        End With
    End With
    
    strName = "qryPackageRules"
    If ItemExists(m.DB.QueryDefs, strName) Then
        m.DB.QueryDefs.Delete strName
    End If
    Set q = m.DB.CreateQueryDef(strName)
    q.SQL = "SELECT tblRules.Name, tblRules.PreviewRTF, tblRules.CondRTF, tblRules.CondCoded, tblRules.PriceRTF, tblRules.PriceCoded, tblRules.Price2RTF, tblRules.Price2Coded, tblRules.OrderPlacement, tblRules.ActionCodedName, tblRules.BuySell, tblRules.LateCondition, tblRules.LateAction, tblRules.RuleType, tblRules.SecurityLevel, tblRules.CannotDelete, tblRules.Password, tblRules.LastModified, tblRules.LibraryID, tblRules.Reverify, tblRules.ExitOnEntryBar, tblRules.ExitBasedOnEachTrade, tblRules.NumberContracts, tblRules.AsPercentOfPosition, tblRules.SystemNumber, tblRules.CheckSum " & _
            "FROM tblRules " & _
            "WHERE (((tblRules.LibraryID) = [SelLibraryID])) " & _
            "ORDER BY tblRules.Name;"

    strName = "qryPackageFunctions"
    If ItemExists(m.DB.QueryDefs, strName) Then
        m.DB.QueryDefs.Delete strName
    End If
    Set q = m.DB.CreateQueryDef(strName)
    q.SQL = "SELECT tblFunctions.FunctionName, tblFunctions.Description, tblFunctions.FunctionCategoryID, tblFunctions.ImplementationTypeID, tblFunctions.ReturnTypeID, tblFunctions.TradeSenseUsage, tblFunctions.Formatted, tblFunctions.FormattedWithFillWords, tblFunctions.CodedName, tblFunctions.CodedText, tblFunctions.Usage, tblFunctions.LateCalculating, tblFunctions.UsesOpenNextBar, tblFunctions.UsesMMData, tblFunctions.SecurityLevel, tblFunctions.CannotDelete, tblFunctions.Password, tblFunctions.LastModified, tblFunctions.LibraryID, tblFunctions.Reverify, tblFunctions.Abbrev, tblFunctions.UsesHLCNextBar, tblFunctions.ListID, tblFunctions.CheckSum " & _
            "FROM tblFunctions " & _
            "WHERE (((tblFunctions.LibraryID) = [SelLibraryID])) " & _
            "ORDER BY tblFunctions.CodedName;"
    
    strName = "qryPackageSystems"
    If ItemExists(m.DB.QueryDefs, strName) Then
        m.DB.QueryDefs.Delete strName
    End If
    Set q = m.DB.CreateQueryDef(strName)
    q.SQL = "SELECT tblSystems.SystemName, tblSystems.BarTimeFrame, tblSystems.Expenses, tblSystems.FromDate, tblSystems.ToDate, tblSystems.ToEndOfData, tblSystems.LibraryID, tblSystems.Developer, tblSystems.TradesPath, tblSystems.Notes, tblSystems.SecurityLevel, tblSystems.CannotDelete, tblSystems.Password, tblSystems.LastModified, tblSystems.BarsLoadedBeforeTrading, tblSystems.BarsTradedBeforeOrders, tblSystems.Reverify, tblSystems.MMID, tblSystems.Pyramid, tblSystems.TradeDepth, tblSystems.CanPyramidSameEntry, tblSystems.LinkInputs, tblSystems.CheckSum " & _
            "FROM tblSystems " & _
            "WHERE (((tblSystems.LibraryID)=[SelLibraryID]));"

    strName = "qryPackageSystemRules"
    If ItemExists(m.DB.QueryDefs, strName) Then
        m.DB.QueryDefs.Delete strName
    End If
    Set q = m.DB.CreateQueryDef(strName)
    q.SQL = "SELECT tblSystems.SystemName, tblRules.Name AS RuleName, tblSystemRules.Seq, tblSystemRules.Selected, tblSystemRules.Alternate, tblSystemRules.RuleUse, tblSystemRules.LastModifiedKnown, tblSystemRules.UnitsID, tblSystemRules.LinkedRules, tblSystemRules.ExitOnEntryBar, tblSystemRules.ExitBasedOnEachTrade, tblSystemRules.NumberContracts, tblSystemRules.AsPercentOfPosition, tblSystemRules.CheckSum " & _
            "FROM tblSystems INNER JOIN (tblRules INNER JOIN tblSystemRules ON tblRules.RuleID = tblSystemRules.RuleID) ON tblSystems.SystemNumber = tblSystemRules.SystemNumber " & _
            "WHERE (((tblSystems.LibraryID)=[SelLibraryID]));"

ErrExit:
    Set q = Nothing
    Set i = Nothing
    Set r = Nothing
    Set t = Nothing
    Exit Sub
    
ErrSection:
    Set q = Nothing
    Set i = Nothing
    Set r = Nothing
    Set t = Nothing
    RaiseError "cDatabaseUpdates.Apply25", eGDRaiseError_Raise
    
End Sub

Private Sub MoveAllRulesLocal(ByVal lItems As Long)
On Error GoTo ErrSection:

    Dim rs As Recordset
    Dim rs2 As Recordset
    Dim alSystemIds As New cGdArray
    Dim lPos As Long
    Dim lIndex As Long
    Dim System As cSystem
    Dim lID As Long
    
    ' Set up the Common Bridge...
    g.CommonBridge.AppPath = App.Path
    g.CommonBridge.CustomerID = g.lLCD
    g.CommonBridge.dbNavRef = m.DB
    
    alSystemIds.Create eGDARRAY_Longs
    
    ' 1) Save all rules currently not used in any systems to add to their
    '    favorites later...
    Set rs = m.DB.OpenRecordset("SELECT tblRules.*, tblLibrarys.BuiltIn " & _
            "FROM tblLibrarys INNER JOIN tblRules ON tblLibrarys.LibraryID = tblRules.LibraryID;", dbOpenDynaset)
    ValidateCheckSums rs, "tblRules"
    Do While Not rs.EOF
        If rs!BuiltIn = False And rs!CheckSum <> 0.5 Then
            Set rs2 = m.DB.OpenRecordset("SELECT * FROM [tblSystemRules] WHERE [RuleID]=" & rs!RuleID & ";", dbOpenDynaset)
            If rs2.EOF And rs2.BOF Then
                rs.Edit
                rs!SystemNumber = -1
                rs!CheckSum = BuildCheckSum(rs, "tblRules")
                rs.Update
            End If
        End If
        
        If frmAsk.Progress <= lItems Then frmAsk.Progress = frmAsk.Progress + 1
        rs.MoveNext
    Loop
    
    ' 2) Make local copies of all shared rules in the system rules table and then
    '    remove the system rules record for the shared rule...
    Set rs = m.DB.OpenRecordset("SELECT tblSystemRules.*, tblRules.SystemNumber " & _
            "FROM tblRules INNER JOIN tblSystemRules ON tblRules.RuleID = tblSystemRules.RuleID;", dbOpenDynaset)
    Do While Not rs.EOF
        If rs![tblRules.SystemNumber] = 0 Then
            lID = rs![tblSystemRules.SystemNumber]
            RuleCopy rs!RuleID, lID, lID
            rs.Delete
            
            If Not alSystemIds.BinarySearch(lID, lPos) Then
                alSystemIds.Add lID, lPos
            End If
        End If
        
        If frmAsk.Progress <= lItems Then frmAsk.Progress = frmAsk.Progress + 1
        rs.MoveNext
    Loop
    
    ' 3) Remove all shared rules not in the built-in libraries and move the previously
    '    saved rules from Step 1 to their "favorites"...
    Set rs = m.DB.OpenRecordset("SELECT tblRules.*, tblLibrarys.BuiltIn " & _
            "FROM tblLibrarys INNER JOIN tblRules ON tblLibrarys.LibraryID = tblRules.LibraryID;", dbOpenDynaset)
    Do While Not rs.EOF
        If rs!SystemNumber = 0 Then
            If rs!BuiltIn = False Then
                rs.Delete
            Else
                Set rs2 = m.DB.OpenRecordset("SELECT * FROM tblRuleParms " & _
                        "WHERE [RuleID]=" & rs!RuleID & ";", dbOpenDynaset)
                Do While Not rs2.EOF
                    m.DB.Execute "DELETE * FROM [tblSystemParms] WHERE [ParmID]=" & rs2!ParmID
                    m.DB.Execute "DELETE * FROM [tblSystemSecurities] WHERE [ParmID]=" & rs2!ParmID
                    rs2.MoveNext
                Loop
            End If
        ElseIf rs!SystemNumber = -1 Then
            rs.Edit
            rs!SystemNumber = 0
            rs!CheckSum = BuildCheckSum(rs, "tblRules")
            rs.Update
        End If
    
        If frmAsk.Progress <= lItems Then frmAsk.Progress = frmAsk.Progress + 1
        rs.MoveNext
    Loop

' don't do this anymore (it's making old systems to be invalid)
#If 0 Then
    ' 4) Make an attempt to Save any systems that were altered during this process...
    On Error Resume Next
    For lIndex = 0 To alSystemIds.Size - 1
        Set System = New cSystem
        
        With System
            .Load alSystemIds(lIndex)
            .Save
        End With
        If frmAsk.Progress <= lItems Then frmAsk.Progress = frmAsk.Progress + 1
    Next lIndex
    On Error GoTo ErrSection
#End If

ErrExit:
    Set rs = Nothing
    Set rs2 = Nothing
    Set alSystemIds = Nothing
    Set System = Nothing
    Exit Sub
    
ErrSection:
    Set rs = Nothing
    Set rs2 = Nothing
    Set alSystemIds = Nothing
    Set System = Nothing
    RaiseError "cDatabaseUpdates.MakeAllRulesLocal", eGDRaiseError_Raise
    
End Sub

Private Sub ConvertAllMarketsToCN()
On Error GoTo ErrSection:

    Dim rs As Recordset
    Dim Bars As New cGdBars

    Set rs = m.DB.OpenRecordset("SELECT * FROM [tblSystemSecurities];", dbOpenDynaset)
    Do While Not rs.EOF
        If rs!Format <> "CN" Then
            If rs!Format = "GT" Then
                Bars.Prop(eBARS_Symbol) = rs!Symbol & ".GT"
            Else
                Bars.Prop(eBARS_Symbol) = rs!Symbol
            End If
            
            rs.Edit
            rs!Symbol = Bars.Prop(eBARS_Symbol)
            rs!Path = AddSlash(App.Path) & "Data"
            rs!Format = "CN"
            rs!SecurityName = ""
            rs.Update
        Else
            rs.Edit
            rs!Path = AddSlash(App.Path) & "Data"
            rs.Update
        End If
            
        rs.MoveNext
    Loop

ErrExit:
    Set Bars = Nothing
    Set rs = Nothing
    Exit Sub
    
ErrSection:
    Set Bars = Nothing
    Set rs = Nothing
    RaiseError "cDatabaseUpdates.ConvertAllMarketsToCN", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ChangeDataType
'' Description: Change the data type on a field in a table
'' Inputs:      Table Name, Field Name, New Data Type
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ChangeDataType(ByVal strTable$, ByVal strField$, Optional ByVal nNewDataType As DataTypeEnum = dbLongBinary)
On Error GoTo ErrSection:

    ChangeAccessFieldDataType m.DB, strTable, strField, nNewDataType

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cTTUpdates.ChangeDataType"
    
End Sub

Private Sub FixTableChecksums(ByVal strTable$)
        
    Dim rs As Recordset
    
    Set rs = g.dbNav.OpenRecordset("SELECT * FROM [" & strTable & "];", dbOpenDynaset)
    Do While Not rs.EOF
        If rs!CheckSum <> 0.5 Then
            rs.Edit
            rs!CheckSum = BuildCheckSum(rs, strTable)
            rs.Update
        End If
        rs.MoveNext
    Loop

End Sub

Private Sub AddRuleCategories()
On Error GoTo ErrSection:

    Dim t As TableDef
    Dim i As Index
    Dim r As Relation
    Dim rs As Recordset
    Dim q As QueryDef
    Dim strName As String
    
    ' Add the CategoryID field to the Rules table...
    Set t = m.DB.TableDefs("tblRules")
    With t
        If Not ItemExists(.Fields, "CategoryID") Then
            .Fields.Append .CreateField("CategoryID", dbLong)
            .Fields("CategoryID").DefaultValue = 0
        End If
    End With

    ' Add the CategoryID field to the Package Rules query...
    strName = "qryPackageRules"
    If ItemExists(m.DB.QueryDefs, strName) Then
        m.DB.QueryDefs.Delete strName
    End If
    Set q = m.DB.CreateQueryDef(strName)
    q.SQL = "SELECT tblRules.Name, tblRules.PreviewRTF, tblRules.CondRTF, tblRules.CondCoded, tblRules.PriceRTF, tblRules.PriceCoded, tblRules.Price2RTF, tblRules.Price2Coded, tblRules.OrderPlacement, tblRules.ActionCodedName, tblRules.BuySell, tblRules.LateCondition, tblRules.LateAction, tblRules.RuleType, tblRules.SecurityLevel, tblRules.CannotDelete, tblRules.Password, tblRules.LastModified, tblRules.LibraryID, tblRules.Reverify, tblRules.ExitOnEntryBar, tblRules.ExitBasedOnEachTrade, tblRules.NumberContracts, tblRules.AsPercentOfPosition, IIf(IsNull([tblSystems].[SystemName]),0,[tblSystems].[SystemName]) AS SystemName, tblRules.CheckSum, tblRules.CategoryID " & _
            "FROM tblRules LEFT JOIN tblSystems ON tblRules.SystemNumber = tblSystems.SystemNumber " & _
            "WHERE (((tblRules.LibraryID) = [SelLibraryID])) " & _
            "ORDER BY tblRules.Name;"
    
    ' Add the Rule Categories lookup table...
    If Not ItemExists(m.DB.TableDefs, "tblRuleCategories") Then
        Set t = m.DB.CreateTableDef("tblRuleCategories")
        With t
            If Not ItemExists(.Fields, "CategoryID") Then
                .Fields.Append .CreateField("CategoryID", dbLong)
                .Fields("CategoryID").Required = True
                .Fields("CategoryID").Attributes = dbAutoIncrField
            End If
            If Not ItemExists(.Fields, "CategoryName") Then
                .Fields.Append .CreateField("CategoryName", dbText, 50)
                .Fields("CategoryName").Required = True
                .Fields("CategoryName").AllowZeroLength = False
            End If
            If Not ItemExists(.Fields, "Description") Then
                .Fields.Append .CreateField("Description", dbText)
                .Fields("Description").Required = False
                .Fields("Description").AllowZeroLength = True
            End If
            
            If Not ItemExists(.Indexes, "Primary") Then
                Set i = .CreateIndex("Primary")
                With i
                    .Fields.Append .CreateField("CategoryID")
                    .Primary = True
                End With
                .Indexes.Append i
            End If
        End With
        m.DB.TableDefs.Append t
    End If
    
    ' Add the relationship between the Rules table and the Rule Categories table...
    If Not ItemExists(m.DB.Relations, "RuleCategories") Then
        Set r = m.DB.CreateRelation("RuleCategories", "tblRuleCategories", "tblRules", _
            dbRelationDeleteCascade)
        With r
            .Fields.Append .CreateField("CategoryID")
            .Fields.Item("CategoryID").ForeignName = "CategoryID"
        End With
        m.DB.Relations.Append r
    End If
    
    ' Populate the Rules Category table...
    Set rs = m.DB.OpenRecordset("SELECT * FROM [tblRuleCategories];", dbOpenDynaset)
    If rs.EOF And rs.BOF Then
        AddRuleCategory 0, "Local Rule", "Rule already contained in a strategy"
        AddRuleCategory 1, "Other Entries", "Other Entries"
        AddRuleCategory 2, "Other Exits", "Other Exits"
        AddRuleCategory 3, "Breakouts", "Entries and Exits based on Breakouts"
        AddRuleCategory 4, "Cycles", "Entries and Exits based on Cycles"
        AddRuleCategory 5, "Trending", "Entries and Exits based on Trends"
        AddRuleCategory 6, "Patterns", "Entries and Exits based on Patterns"
        AddRuleCategory 7, "Volatility", "Entries and Exits based on Volatility"
        AddRuleCategory 8, "Swing Trading", "Entries and Exits based on Swing Trading"
        AddRuleCategory 9, "Profitable Exits", "Rules designed to exit a position at specific profit levels"
        AddRuleCategory 10, "Protective Stops", "Rules designed to limit the loss when trade goes against you"
    End If
    
    ' Set default values for preexisting rules...
    Set rs = m.DB.OpenRecordset("SELECT * FROM [tblRules];", dbOpenDynaset)
    Do While Not rs.EOF
        rs.Edit
        If rs!SystemNumber = 0 Then
            If rs!RuleType = 0 Then
                rs!CategoryID = 1
            Else
                rs!CategoryID = 2
            End If
        Else
            rs!CategoryID = 0
        End If
        rs.Update
        
        rs.MoveNext
    Loop

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cDatabaseUpdates.AddRuleCategories", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddRuleCategory
'' Description: Add a rule category to the database
'' Inputs:      ID, Name, and Description of the rule category
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddRuleCategory(ByVal lID As Long, ByVal strName As String, ByVal strDesc As String)
On Error GoTo ErrSection:

    Dim rs As Recordset

    Set rs = m.DB.OpenRecordset("SELECT * FROM [tblRuleCategories];", dbOpenDynaset)

    rs.AddNew
    rs!CategoryID = lID
    rs!CategoryName = strName
    rs!Description = strDesc
    rs.Update

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cDatabaseUpdates.AddRuleCategory", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    AddTradingItemTable
'' Description: Add the Automated Trading Item table to the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub AddTradingItemTable()
On Error GoTo ErrSection:

    Dim t As TableDef                   ' Temporary TableDef object
    Dim i As Index                      ' Temporary Database Index object

    If Not ItemExists(m.DB.TableDefs, "tblAutoTradingItem") Then
        Set t = m.DB.CreateTableDef("tblAutoTradingItem")
        With t
            If Not ItemExists(.Fields, "TradingItemID") Then
                .Fields.Append .CreateField("TradingItemID", dbLong)
                .Fields("TradingItemID").Attributes = dbAutoIncrField
            End If
            If Not ItemExists(.Fields, "StrategyID") Then
                .Fields.Append .CreateField("StrategyID", dbLong)
                .Fields("StrategyID").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "StrategyName") Then
                .Fields.Append .CreateField("StrategyName", dbText, 255)
                .Fields("StrategyName").AllowZeroLength = True
                .Fields("StrategyName").Required = False
                .Fields("StrategyName").DefaultValue = ""
            End If
            If Not ItemExists(.Fields, "SymbolID") Then
                .Fields.Append .CreateField("SymbolID", dbLong)
                .Fields("SymbolID").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "Symbol") Then
                .Fields.Append .CreateField("Symbol", dbText, 50)
                .Fields("Symbol").AllowZeroLength = True
                .Fields("Symbol").Required = False
                .Fields("Symbol").DefaultValue = ""
            End If
            If Not ItemExists(.Fields, "Period") Then
                .Fields.Append .CreateField("Period", dbText, 50)
                .Fields("Symbol").AllowZeroLength = True
                .Fields("Symbol").Required = False
                .Fields("Symbol").DefaultValue = ""
            End If
            If Not ItemExists(.Fields, "AccountID") Then
                .Fields.Append .CreateField("AccountID", dbLong)
                .Fields("AccountID").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "CurrentPosition") Then
                .Fields.Append .CreateField("CurrentPosition", dbLong)
                .Fields("CurrentPosition").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "QtyNextEntry") Then
                .Fields.Append .CreateField("QtyNextEntry", dbLong)
                .Fields("QtyNextEntry").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "OnCloseTime") Then
                .Fields.Append .CreateField("OnCloseTime", dbDouble)
                .Fields("OnCloseTime").DefaultValue = 0#
            End If
            If Not ItemExists(.Fields, "MinutesBeforeClose") Then
                .Fields.Append .CreateField("MinutesBeforeClose", dbLong)
                .Fields("MinutesBeforeClose").DefaultValue = 5&
            End If
            If Not ItemExists(.Fields, "ConfirmOrders") Then
                .Fields.Append .CreateField("ConfirmOrders", dbBoolean)
                .Fields("ConfirmOrders").DefaultValue = True
            End If
        
            If Not ItemExists(.Indexes, "Primary") Then
                Set i = .CreateIndex("Primary")
                With i
                    .Fields.Append .CreateField("TradingItemID")
                    .Primary = True
                End With
                .Indexes.Append i
            End If
            
            m.DB.TableDefs.Append t
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cDatabaseUpdates.AddTradingItemTable", eGDRaiseError_Raise
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion48
'' Description: Add an AccountPositionID field to the auto trading item table
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion48()
On Error GoTo ErrSection:

    Dim t As TableDef                   ' Temporary TableDef object
    Dim q As QueryDef

    If ItemExists(m.DB.TableDefs, "tblAutoTradingItem") Then
        Set t = m.DB.TableDefs("tblAutoTradingItem")
        With t
            If Not ItemExists(.Fields, "AccountPositionID") Then
                .Fields.Append .CreateField("AccountPositionID", dbLong)
                .Fields("AccountPositionID").DefaultValue = 0&
            
                Set q = m.DB.CreateQueryDef("")
                q.SQL = "UPDATE tblAutoTradingItem SET tblAutoTradingItem.AccountPositionID = 0;"
                q.Execute
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cDatabaseUpdates.ApplyVersion48"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion49
'' Description: Add a Deleted field to the auto trading item table
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion49()
On Error GoTo ErrSection:

    Dim t As TableDef                   ' Temporary TableDef object
    Dim q As QueryDef                   ' Query into the database

    If ItemExists(m.DB.TableDefs, "tblAutoTradingItem") Then
        Set t = m.DB.TableDefs("tblAutoTradingItem")
        With t
            If Not ItemExists(.Fields, "Deleted") Then
                .Fields.Append .CreateField("Deleted", dbBoolean)
                .Fields("Deleted").DefaultValue = False
            
                Set q = m.DB.CreateQueryDef("")
                q.SQL = "UPDATE tblAutoTradingItem SET tblAutoTradingItem.Deleted = No;"
                q.Execute
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cDatabaseUpdates.ApplyVersion49"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion50
'' Description: Add a On Close Exchange time field to the automated trading
''              item table
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion50()
On Error GoTo ErrSection:

    Dim t As TableDef                   ' Temporary TableDef object
    Dim q As QueryDef                   ' Query into the database

    If ItemExists(m.DB.TableDefs, "tblAutoTradingItem") Then
        Set t = m.DB.TableDefs("tblAutoTradingItem")
        With t
            If Not ItemExists(.Fields, "OnCloseTimeExch") Then
                .Fields.Append .CreateField("OnCloseTimeExch", dbDouble)
                .Fields("OnCloseTimeExch").DefaultValue = 0
            
                Set q = m.DB.CreateQueryDef("")
                q.SQL = "UPDATE tblAutoTradingItem SET tblAutoTradingItem.OnCloseTimeExch = 0;"
                q.Execute
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cDatabaseUpdates.ApplyVersion50"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion51
'' Description: Bring strategy baskets into the database
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion51()
On Error GoTo ErrSection:

    Dim t As TableDef                   ' Temporary TableDef object
    Dim i As Index                      ' Temporary Database Index object
    Dim r As Relation                   ' Relation between tables for the database
    Dim q As QueryDef                   ' Query definition for the database
    Dim strName As String               ' Name of the query definition

    If Not ItemExists(m.DB.TableDefs, "tblStrategyBaskets") Then
        g.bImportStrategyBaskets = True
        
        Set t = m.DB.CreateTableDef("tblStrategyBaskets")
        With t
            If Not ItemExists(.Fields, "StrategyBasketID") Then
                .Fields.Append .CreateField("StrategyBasketID", dbLong)
                .Fields("StrategyBasketID").Attributes = dbAutoIncrField
            End If
            If Not ItemExists(.Fields, "Name") Then
                .Fields.Append .CreateField("Name", dbText, 255)
                .Fields("Name").AllowZeroLength = True
                .Fields("Name").Required = False
                .Fields("Name").DefaultValue = ""
            End If
            If Not ItemExists(.Fields, "Description") Then
                .Fields.Append .CreateField("Description", dbText, 255)
                .Fields("Description").AllowZeroLength = True
                .Fields("Description").Required = False
                .Fields("Description").DefaultValue = ""
            End If
            If Not ItemExists(.Fields, "LastModified") Then
                .Fields.Append .CreateField("LastModified", dbDouble)
                .Fields("LastModified").DefaultValue = 0
            End If
            If Not ItemExists(.Fields, "RequiredMod") Then
                .Fields.Append .CreateField("RequiredMod", dbMemo)
                .Fields("RequiredMod").AllowZeroLength = True
                .Fields("RequiredMod").Required = False
                .Fields("RequiredMod").DefaultValue = ""
            End If
            If Not ItemExists(.Fields, "LibraryID") Then
                .Fields.Append .CreateField("LibraryID", dbLong)
                .Fields("LibraryID").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "SecurityLevel") Then
                .Fields.Append .CreateField("SecurityLevel", dbLong)
                .Fields("SecurityLevel").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "CannotDelete") Then
                .Fields.Append .CreateField("CannotDelete", dbBoolean)
                .Fields("CannotDelete").DefaultValue = False
            End If
            If Not ItemExists(.Fields, "Password") Then
                .Fields.Append .CreateField("Password", dbLongBinary)
            End If
            If Not ItemExists(.Fields, "CheckSum") Then
                .Fields.Append .CreateField("CheckSum", dbDouble)
                .Fields("CheckSum").DefaultValue = 0
            End If
        
            If Not ItemExists(.Indexes, "Primary") Then
                Set i = .CreateIndex("Primary")
                With i
                    .Fields.Append .CreateField("StrategyBasketID")
                    .Primary = True
                End With
                .Indexes.Append i
            End If
            
            m.DB.TableDefs.Append t
        End With
    End If

    If Not ItemExists(m.DB.TableDefs, "tblStrategyBasketItems") Then
        Set t = m.DB.CreateTableDef("tblStrategyBasketItems")
        With t
            If Not ItemExists(.Fields, "StrategyBasketItemID") Then
                .Fields.Append .CreateField("StrategyBasketItemID", dbLong)
                .Fields("StrategyBasketItemID").Attributes = dbAutoIncrField
            End If
            If Not ItemExists(.Fields, "StrategyBasketID") Then
                .Fields.Append .CreateField("StrategyBasketID", dbLong)
                .Fields("StrategyBasketID").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "SystemNumber") Then
                .Fields.Append .CreateField("SystemNumber", dbLong)
                .Fields("SystemNumber").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "SymbolGroupID") Then
                .Fields.Append .CreateField("SymbolGroupID", dbText, 255)
                .Fields("SymbolGroupID").AllowZeroLength = True
                .Fields("SymbolGroupID").Required = False
                .Fields("SymbolGroupID").DefaultValue = ""
            End If
            If Not ItemExists(.Fields, "SymbolGroupName") Then
                .Fields.Append .CreateField("SymbolGroupName", dbText, 255)
                .Fields("SymbolGroupName").AllowZeroLength = True
                .Fields("SymbolGroupName").Required = False
                .Fields("SymbolGroupName").DefaultValue = ""
            End If
            If Not ItemExists(.Fields, "SymbolID") Then
                .Fields.Append .CreateField("SymbolID", dbLong)
                .Fields("SymbolID").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "Symbol") Then
                .Fields.Append .CreateField("Symbol", dbText, 255)
                .Fields("Symbol").AllowZeroLength = True
                .Fields("Symbol").Required = False
                .Fields("Symbol").DefaultValue = ""
            End If
            If Not ItemExists(.Fields, "BarPeriod") Then
                .Fields.Append .CreateField("BarPeriod", dbText, 255)
                .Fields("BarPeriod").AllowZeroLength = True
                .Fields("BarPeriod").Required = False
                .Fields("BarPeriod").DefaultValue = ""
            End If
            If Not ItemExists(.Fields, "FromDate") Then
                .Fields.Append .CreateField("FromDate", dbLong)
                .Fields("FromDate").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "ToDate") Then
                .Fields.Append .CreateField("ToDate", dbLong)
                .Fields("ToDate").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "ToEndOfData") Then
                .Fields.Append .CreateField("ToEndOfData", dbBoolean)
                .Fields("ToEndOfData").DefaultValue = True
            End If
            If Not ItemExists(.Fields, "SplitAdjust") Then
                .Fields.Append .CreateField("SplitAdjust", dbBoolean)
                .Fields("SplitAdjust").DefaultValue = True
            End If
            If Not ItemExists(.Fields, "ContractMultiplier") Then
                .Fields.Append .CreateField("ContractMultiplier", dbDouble)
                .Fields("ContractMultiplier").DefaultValue = 1#
            End If
            If Not ItemExists(.Fields, "CheckSum") Then
                .Fields.Append .CreateField("CheckSum", dbDouble)
                .Fields("CheckSum").DefaultValue = 0
            End If
        
            If Not ItemExists(.Indexes, "Primary") Then
                Set i = .CreateIndex("Primary")
                With i
                    .Fields.Append .CreateField("StrategyBasketItemID")
                    .Primary = True
                End With
                .Indexes.Append i
            End If
            
            m.DB.TableDefs.Append t
        End With
    End If

    If Not ItemExists(m.DB.TableDefs, "tblStrategyBasketItemParms") Then
        Set t = m.DB.CreateTableDef("tblStrategyBasketItemParms")
        With t
            If Not ItemExists(.Fields, "StrategyBasketItemID") Then
                .Fields.Append .CreateField("StrategyBasketItemID", dbLong)
                .Fields("StrategyBasketItemID").DefaultValue = 0&
            End If
            If Not ItemExists(.Fields, "RuleName") Then
                .Fields.Append .CreateField("RuleName", dbText, 255)
                .Fields("RuleName").AllowZeroLength = True
                .Fields("RuleName").Required = False
                .Fields("RuleName").DefaultValue = ""
            End If
            If Not ItemExists(.Fields, "ParmName") Then
                .Fields.Append .CreateField("ParmName", dbText, 255)
                .Fields("ParmName").AllowZeroLength = True
                .Fields("ParmName").Required = False
                .Fields("ParmName").DefaultValue = ""
            End If
            If Not ItemExists(.Fields, "Value") Then
                .Fields.Append .CreateField("Value", dbText, 255)
                .Fields("Value").AllowZeroLength = True
                .Fields("Value").Required = False
                .Fields("Value").DefaultValue = ""
            End If
        
            If Not ItemExists(.Indexes, "Primary") Then
                Set i = .CreateIndex("Primary")
                With i
                    .Fields.Append .CreateField("StrategyBasketItemID")
                    .Fields.Append .CreateField("RuleName")
                    .Fields.Append .CreateField("ParmName")
                    .Primary = True
                End With
                .Indexes.Append i
            End If
            
            m.DB.TableDefs.Append t
        End With
    End If

    If Not ItemExists(m.DB.Relations, "LibrarysStrategyBasket") Then
        Set r = m.DB.CreateRelation("LibrarysStrategyBasket", "tblLibrarys", "tblStrategyBaskets", dbRelationDeleteCascade)
        With r
            .Fields.Append .CreateField("LibraryID")
            .Fields.Item("LibraryID").ForeignName = "LibraryID"
        End With
        
        m.DB.Relations.Append r
    End If

    If Not ItemExists(m.DB.Relations, "StrategyBasketsStrategyBasketItems") Then
        Set r = m.DB.CreateRelation("StrategyBasketsStrategyBasketItems", "tblStrategyBaskets", "tblStrategyBasketItems", dbRelationDeleteCascade)
        With r
            .Fields.Append .CreateField("StrategyBasketID")
            .Fields.Item("StrategyBasketID").ForeignName = "StrategyBasketID"
        End With
        
        m.DB.Relations.Append r
    End If

    If Not ItemExists(m.DB.Relations, "SystemsStrategyBasketItems") Then
        Set r = m.DB.CreateRelation("SystemsStrategyBasketItems", "tblSystems", "tblStrategyBasketItems", dbRelationDeleteCascade)
        With r
            .Fields.Append .CreateField("SystemNumber")
            .Fields.Item("SystemNumber").ForeignName = "SystemNumber"
        End With
        
        m.DB.Relations.Append r
    End If

    If Not ItemExists(m.DB.Relations, "StrategyBasketItemsStrategyBasketItemParms") Then
        Set r = m.DB.CreateRelation("StrategyBasketItemsStrategyBasketItemParms", "tblStrategyBasketItems", "tblStrategyBasketItemParms", dbRelationDeleteCascade)
        With r
            .Fields.Append .CreateField("StrategyBasketItemID")
            .Fields.Item("StrategyBasketItemID").ForeignName = "StrategyBasketItemID"
        End With
        
        m.DB.Relations.Append r
    End If
    
    strName = "qryPackageStrategyBaskets"
    If ItemExists(m.DB.QueryDefs, strName) Then
        m.DB.QueryDefs.Delete strName
    End If
    Set q = m.DB.CreateQueryDef(strName)
    q.SQL = "SELECT tblStrategyBaskets.StrategyBasketID, tblStrategyBaskets.Name, tblStrategyBaskets.Description, tblStrategyBaskets.LastModified, tblStrategyBaskets.RequiredMod, tblStrategyBaskets.LibraryID, tblStrategyBaskets.SecurityLevel, tblStrategyBaskets.CannotDelete, tblStrategyBaskets.Password, tblStrategyBaskets.CheckSum " & _
            "FROM tblStrategyBaskets " & _
            "WHERE (((tblStrategyBaskets.LibraryID) = [SelLibraryID])) " & _
            "ORDER BY tblStrategyBaskets.Name;"
    
    strName = "qryPackageStrategyBasketItems"
    If ItemExists(m.DB.QueryDefs, strName) Then
        m.DB.QueryDefs.Delete strName
    End If
    Set q = m.DB.CreateQueryDef(strName)
    q.SQL = "SELECT tblStrategyBasketItems.StrategyBasketItemID, tblStrategyBasketItems.StrategyBasketID, tblSystems.SystemName, tblStrategyBasketItems.SymbolGroupID, tblStrategyBasketItems.SymbolGroupName, tblStrategyBasketItems.SymbolID, tblStrategyBasketItems.Symbol, tblStrategyBasketItems.BarPeriod, tblStrategyBasketItems.FromDate, tblStrategyBasketItems.ToDate, tblStrategyBasketItems.ToEndOfData, tblStrategyBasketItems.SplitAdjust, tblStrategyBasketItems.ContractMultiplier, tblStrategyBasketItems.CheckSum " & _
            "FROM (tblStrategyBasketItems INNER JOIN tblStrategyBaskets ON tblStrategyBasketItems.StrategyBasketID=tblStrategyBaskets.StrategyBasketID) INNER JOIN tblSystems ON tblStrategyBasketItems.SystemNumber=tblSystems.SystemNumber " & _
            "WHERE (((tblStrategyBaskets.LibraryID) = [SelLibraryID]));"
    
    strName = "qryPackageStrategyBasketItemParms"
    If ItemExists(m.DB.QueryDefs, strName) Then
        m.DB.QueryDefs.Delete strName
    End If
    Set q = m.DB.CreateQueryDef(strName)
    q.SQL = "SELECT tblStrategyBasketItemParms.StrategyBasketItemID, tblStrategyBasketItemParms.RuleName, tblStrategyBasketItemParms.ParmName, tblStrategyBasketItemParms.Value " & _
            "FROM (tblStrategyBasketItemParms INNER JOIN (tblStrategyBasketItems INNER JOIN tblStrategyBaskets ON tblStrategyBasketItems.StrategyBasketID=tblStrategyBaskets.StrategyBasketID) ON tblStrategyBasketItemParms.StrategyBasketItemID=tblStrategyBasketItems.StrategyBasketItemID) " & _
            "WHERE (((tblStrategyBaskets.LibraryID) = [SelLibraryID]));"
    
    strName = "qryDep20"
    If ItemExists(m.DB.QueryDefs, strName) Then
        m.DB.QueryDefs.Delete strName
    End If
    Set q = m.DB.CreateQueryDef(strName)
    q.SQL = "SELECT tblLibrarys.LibraryName, First(tblLibrarys.LastModified) AS LastMod, First(tblLibrarys.Version) AS Vers " & _
            "FROM (tblStrategyBasketItems INNER JOIN (tblStrategyBaskets INNER JOIN tblLibrarys ON tblStrategyBaskets.LibraryID=tblLibrarys.LibraryID) ON tblStrategyBasketItems.StrategyBasketID=tblStrategyBaskets.StrategyBasketID) INNER JOIN tblSystems on tblStrategyBasketItems.SystemNumber=tblSystems.SystemNumber  " & _
            "GROUP BY tblLibrarys.LibraryName, tblStrategyBaskets.LibraryID, tblSystems.LibraryID " & _
            "HAVING (((tblSystems.LibraryID) <> [SelLibID]) And ((tblStrategyBaskets.LibraryID) = [SelLibID])) " & _
            "ORDER BY tblStrategyBaskets.LibraryID;"
        
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cDatabaseUpdates.ApplyVersion51"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion52
'' Description: Extra fields in automated trading items table for baskets
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion52()
On Error GoTo ErrSection:

    Dim t As TableDef                   ' Temporary TableDef object
    Dim q As QueryDef                   ' Query into the database
    Dim rs As Recordset                 ' Recordset into the database

    If ItemExists(m.DB.TableDefs, "tblAutoTradingItem") Then
        Set t = m.DB.TableDefs("tblAutoTradingItem")
        With t
            If Not ItemExists(.Fields, "ParentID") Then
                .Fields.Append .CreateField("ParentID", dbLong)
                .Fields("ParentID").DefaultValue = 0&
            
                Set q = m.DB.CreateQueryDef("")
                q.SQL = "UPDATE tblAutoTradingItem SET tblAutoTradingItem.ParentID = 0;"
                q.Execute
            End If
            If Not ItemExists(.Fields, "StrategyBasketID") Then
                .Fields.Append .CreateField("StrategyBasketID", dbLong)
                .Fields("StrategyBasketID").DefaultValue = 0&
            
                Set q = m.DB.CreateQueryDef("")
                q.SQL = "UPDATE tblAutoTradingItem SET tblAutoTradingItem.StrategyBasketID = 0;"
                q.Execute
            End If
            If Not ItemExists(.Fields, "StrategyBasketItemID") Then
                .Fields.Append .CreateField("StrategyBasketItemID", dbLong)
                .Fields("StrategyBasketItemID").DefaultValue = 0&
            
                Set q = m.DB.CreateQueryDef("")
                q.SQL = "UPDATE tblAutoTradingItem SET tblAutoTradingItem.StrategyBasketItemID = 0;"
                q.Execute
            End If
            If Not ItemExists(.Fields, "StrategyBasketItemMult") Then
                .Fields.Append .CreateField("StrategyBasketItemMult", dbLong)
                .Fields("StrategyBasketItemMult").DefaultValue = 1&
            
                Set q = m.DB.CreateQueryDef("")
                q.SQL = "UPDATE tblAutoTradingItem SET tblAutoTradingItem.StrategyBasketItemMult = 1;"
                q.Execute
            End If
            If Not ItemExists(.Fields, "StrategyBasketLastModified") Then
                .Fields.Append .CreateField("StrategyBasketLastModified", dbDouble)
                .Fields("StrategyBasketLastModified").DefaultValue = 0#
            
                Set q = m.DB.CreateQueryDef("")
                q.SQL = "UPDATE tblAutoTradingItem SET tblAutoTradingItem.StrategyBasketLastModified = 0;"
                q.Execute
            End If
            If Not ItemExists(.Fields, "Overrides") Then
                .Fields.Append .CreateField("Overrides", dbText, 255)
                .Fields("Overrides").AllowZeroLength = True
                .Fields("Overrides").Required = False
                .Fields("Overrides").DefaultValue = ""
            
                Set q = m.DB.CreateQueryDef("")
                q.SQL = "UPDATE tblAutoTradingItem SET tblAutoTradingItem.Overrides = '';"
                q.Execute
            End If
            If Not ItemExists(.Fields, "StrategyBasketItemKey") Then
                .Fields.Append .CreateField("StrategyBasketItemKey", dbText, 255)
                .Fields("StrategyBasketItemKey").AllowZeroLength = True
                .Fields("StrategyBasketItemKey").Required = False
                .Fields("StrategyBasketItemKey").DefaultValue = ""
            
                Set q = m.DB.CreateQueryDef("")
                q.SQL = "UPDATE tblAutoTradingItem SET tblAutoTradingItem.StrategyBasketItemKey = '';"
                q.Execute
            End If
        End With
    End If
    
    Set rs = m.DB.OpenRecordset("SELECT * FROM [tblAutoTradingItem];", dbOpenDynaset)
    Do While Not rs.EOF
        If rs!Deleted = 0 Then
            If InStr(rs!Name, "|") <> 0 Then
                rs.Edit
                rs!Name = Replace(rs!Name, "|", "_")
                rs.Update
            End If
        Else
            If InStr(rs!Name, "|") = 0 Then
                rs.Edit
                
                If Len(rs!Name) = 50 Then
                    rs!Name = Left(rs!Name, 49) & "|"
                Else
                    rs!Name = rs!Name & "|"
                End If
                
                rs.Update
            End If
        End If
        
        rs.MoveNext
    Loop
    
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cDatabaseUpdates.ApplyVersion52"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion53
'' Description: Add "Guru" security flag to strategies and baskets
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion53()
On Error GoTo ErrSection:

    Dim t As TableDef                   ' Temporary TableDef object
    Dim q As QueryDef                   ' Query into the database
    Dim strName As String               ' Name of the query definition

    If ItemExists(m.DB.TableDefs, "tblSystems") Then
        Set t = m.DB.TableDefs("tblSystems")
        With t
            If Not ItemExists(.Fields, "IsGuru") Then
                .Fields.Append .CreateField("IsGuru", dbBoolean)
                .Fields("IsGuru").DefaultValue = False
            
                Set q = m.DB.CreateQueryDef("")
                q.SQL = "UPDATE tblSystems SET tblSystems.IsGuru = 0;"
                q.Execute
            End If
        End With
    End If

    If ItemExists(m.DB.TableDefs, "tblStrategyBaskets") Then
        Set t = m.DB.TableDefs("tblStrategyBaskets")
        With t
            If Not ItemExists(.Fields, "IsGuru") Then
                .Fields.Append .CreateField("IsGuru", dbBoolean)
                .Fields("IsGuru").DefaultValue = False
            
                Set q = m.DB.CreateQueryDef("")
                q.SQL = "UPDATE tblStrategyBaskets SET tblStrategyBaskets.IsGuru = 0;"
                q.Execute
            End If
        End With
    End If

    If ItemExists(m.DB.TableDefs, "tblLibrarys") Then
        Set t = m.DB.TableDefs("tblLibrarys")
        With t
            If Not ItemExists(.Fields, "IsGuru") Then
                .Fields.Append .CreateField("IsGuru", dbBoolean)
                .Fields("IsGuru").DefaultValue = False
            
                Set q = m.DB.CreateQueryDef("")
                q.SQL = "UPDATE tblLibrarys SET tblLibrarys.IsGuru = 0;"
                q.Execute
            End If
            
            If Not ItemExists(.Fields, "Owners") Then
                .Fields.Append .CreateField("Owners", dbMemo)
                .Fields("Owners").AllowZeroLength = True
                .Fields("Owners").Required = False
                .Fields("Owners").DefaultValue = ""
                
                Set q = m.DB.CreateQueryDef("")
                q.SQL = "UPDATE tblLibrarys SET tblLibrarys.Owners = '';"
                q.Execute
            End If
        End With
    End If
    
    strName = "qryPackageSystems"
    If ItemExists(m.DB.QueryDefs, strName) Then
        m.DB.QueryDefs.Delete strName
    End If
    Set q = m.DB.CreateQueryDef(strName)
    q.SQL = "SELECT tblSystems.SystemName, tblSystems.BarTimeFrame, tblSystems.Expenses, tblSystems.FromDate, tblSystems.ToDate, tblSystems.ToEndOfData, tblSystems.LibraryID, tblSystems.Developer, tblSystems.TradesPath, tblSystems.Notes, tblSystems.SecurityLevel, tblSystems.CannotDelete, tblSystems.Password, tblSystems.LastModified, tblSystems.BarsLoadedBeforeTrading, tblSystems.BarsTradedBeforeOrders, tblSystems.Reverify, tblSystems.MMID, tblSystems.Pyramid, tblSystems.TradeDepth, tblSystems.CanPyramidSameEntry, tblSystems.LinkInputs, tblSystems.CheckSum " & _
            "FROM tblSystems " & _
            "WHERE (((tblSystems.LibraryID)=[SelLibraryID]));"

    strName = "qryPackageStrategyBaskets"
    If ItemExists(m.DB.QueryDefs, strName) Then
        m.DB.QueryDefs.Delete strName
    End If
    Set q = m.DB.CreateQueryDef(strName)
    q.SQL = "SELECT tblStrategyBaskets.StrategyBasketID, tblStrategyBaskets.Name, tblStrategyBaskets.Description, tblStrategyBaskets.LastModified, tblStrategyBaskets.RequiredMod, tblStrategyBaskets.LibraryID, tblStrategyBaskets.SecurityLevel, tblStrategyBaskets.CannotDelete, tblStrategyBaskets.Password, tblStrategyBaskets.CheckSum, tblStrategyBaskets.IsGuru " & _
            "FROM tblStrategyBaskets " & _
            "WHERE (((tblStrategyBaskets.LibraryID) = [SelLibraryID])) " & _
            "ORDER BY tblStrategyBaskets.Name;"
    
    strName = "qryDep20"
    If ItemExists(m.DB.QueryDefs, strName) Then
        m.DB.QueryDefs.Delete strName
    End If
    Set q = m.DB.CreateQueryDef(strName)
    q.SQL = "SELECT tblLibrarys.LibraryName, First(tblLibrarys.LastModified) AS LastMod, First(tblLibrarys.Version) AS Vers " & _
            "FROM tblStrategyBaskets INNER JOIN " & _
                    "(tblStrategyBasketItems INNER JOIN " & _
                        "(tblSystems INNER JOIN tblLibrarys ON tblSystems.LibraryID = tblLibrarys.LibraryID) " & _
                    "ON tblStrategyBasketItems.SystemNumber = tblSystems.SystemNumber) " & _
                "ON tblStrategyBaskets.StrategyBasketID = tblStrategyBasketItems.StrategyBasketID " & _
            "GROUP BY tblLibrarys.LibraryName, tblStrategyBaskets.LibraryID, tblSystems.LibraryID " & _
            "HAVING (((tblSystems.LibraryID) <> [SelLibID]) And ((tblStrategyBaskets.LibraryID) = [SelLibID])) " & _
            "ORDER BY tblStrategyBaskets.LibraryID;"
        
    strName = "qryDep21"
    If ItemExists(m.DB.QueryDefs, strName) Then
        m.DB.QueryDefs.Delete strName
    End If
    Set q = m.DB.CreateQueryDef(strName)
    q.SQL = "SELECT tblLibrarys.LibraryName, First(tblLibrarys.LastModified) AS LastMod, First(tblLibrarys.Version) AS Vers " & _
            "FROM tblStrategyBaskets INNER JOIN " & _
                    "(tblStrategyBasketItems INNER JOIN " & _
                        "(tblSystems INNER JOIN " & _
                            "(tblSystemRules INNER JOIN " & _
                                "(tblRules INNER JOIN tblLibrarys ON tblRules.LibraryID = tblLibrarys.LibraryID) " & _
                            "ON tblSystemRules.RuleID = tblRules.RuleID) " & _
                        "ON tblSystems.SystemNumber = tblSystemRules.SystemNumber) " & _
                    "ON tblStrategyBasketItems.SystemNumber = tblSystems.SystemNumber) " & _
                "ON tblStrategyBaskets.StrategyBasketID = tblStrategyBasketItems.StrategyBasketID " & _
            "GROUP BY tblLibrarys.LibraryName, tblStrategyBaskets.LibraryID, tblRules.LibraryID " & _
            "HAVING (((tblRules.LibraryID) <> [SelLibID]) And ((tblStrategyBaskets.LibraryID) = [SelLibID])) " & _
            "ORDER BY tblStrategyBaskets.LibraryID;"
                
    strName = "qryDep22"
    If ItemExists(m.DB.QueryDefs, strName) Then
        m.DB.QueryDefs.Delete strName
    End If
    Set q = m.DB.CreateQueryDef(strName)
    q.SQL = "SELECT tblLibrarys.LibraryName, First(tblLibrarys.LastModified) AS LastMod, First(tblLibrarys.Version) AS Vers " & _
            "FROM tblStrategyBaskets INNER JOIN " & _
                    "(tblStrategyBasketItems INNER JOIN " & _
                        "(tblSystems INNER JOIN " & _
                            "(tblSystemRules INNER JOIN " & _
                                "(tblRules INNER JOIN " & _
                                    "(tblFunctionRules INNER JOIN " & _
                                        "(tblFunctions INNER JOIN tblLibrarys ON tblFunctions.LibraryID = tblLibrarys.LibraryID) " & _
                                    "ON tblFunctionRules.FunctionIDRef = tblFunctions.FunctionID) " & _
                                "ON tblRules.RuleID = tblFunctionRules.RuleID) " & _
                            "ON tblSystemRules.RuleID = tblRules.RuleID) " & _
                        "ON tblSystems.SystemNumber = tblSystemRules.SystemNumber) " & _
                    "ON tblStrategyBasketItems.SystemNumber = tblSystems.SystemNumber) " & _
                "ON tblStrategyBaskets.StrategyBasketID = tblStrategyBasketItems.StrategyBasketID " & _
            "GROUP BY tblLibrarys.LibraryName, tblStrategyBaskets.LibraryID, tblFunctions.LibraryID " & _
            "HAVING (((tblFunctions.LibraryID) <> [SelLibID]) And ((tblStrategyBaskets.LibraryID) = [SelLibID])) " & _
            "ORDER BY tblStrategyBaskets.LibraryID;"
                
    strName = "qryDep23"
    If ItemExists(m.DB.QueryDefs, strName) Then
        m.DB.QueryDefs.Delete strName
    End If
    Set q = m.DB.CreateQueryDef(strName)
    q.SQL = "SELECT tblLibrarys.LibraryName, First(tblLibrarys.LastModified) AS LastMod, First(tblLibrarys.Version) AS Vers " & _
            "FROM tblStrategyBaskets INNER JOIN " & _
                    "(tblStrategyBasketItems INNER JOIN " & _
                        "(tblSystems INNER JOIN " & _
                            "(tblSystemRules INNER JOIN " & _
                                "(tblRules INNER JOIN " & _
                                    "(tblFunctionRules INNER JOIN " & _
                                        "(tblFunctions INNER JOIN " & _
                                            "(tblFunctionRefs INNER JOIN " & _
                                                "(tblFunctions AS tblFunctions_1 INNER JOIN tblLibrarys ON tblFunctions_1.LibraryID = tblLibrarys.LibraryID) " & _
                                            "ON tblFunctionRefs.FunctionIDRef = tblFunctions_1.FunctionID) " & _
                                        "ON tblFunctions.FunctionID = tblFunctionRefs.FunctionID) " & _
                                    "ON tblFunctionRules.FunctionIDRef = tblFunctions.FunctionID) " & _
                                "ON tblRules.RuleID = tblFunctionRules.RuleID) " & _
                            "ON tblSystemRules.RuleID = tblRules.RuleID) " & _
                        "ON tblSystems.SystemNumber = tblSystemRules.SystemNumber) " & _
                    "ON tblStrategyBasketItems.SystemNumber = tblSystems.SystemNumber) " & _
                "ON tblStrategyBaskets.StrategyBasketID = tblStrategyBasketItems.StrategyBasketID " & _
            "GROUP BY tblLibrarys.LibraryName, tblStrategyBaskets.LibraryID, tblFunctions_1.LibraryID " & _
            "HAVING (((tblFunctions_1.LibraryID) <> [SelLibID]) And ((tblStrategyBaskets.LibraryID) = [SelLibID])) " & _
            "ORDER BY tblStrategyBaskets.LibraryID;"
                
ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cDatabaseUpdates.ApplyVersion53"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion54
'' Description: Fix the qryPackageSystems that got screwed up in version 53
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion54()
On Error GoTo ErrSection:

    Dim q As QueryDef                   ' Query into the database
    Dim strName As String               ' Name of the query definition

    strName = "qryPackageSystems"
    If ItemExists(m.DB.QueryDefs, strName) Then
        m.DB.QueryDefs.Delete strName
    End If
    Set q = m.DB.CreateQueryDef(strName)
    q.SQL = "SELECT tblSystems.SystemName, tblSystems.BarTimeFrame, tblSystems.Expenses, tblSystems.FromDate, tblSystems.ToDate, tblSystems.ToEndOfData, tblSystems.LibraryID, tblSystems.Developer, tblSystems.TradesPath, tblSystems.Notes, tblSystems.SecurityLevel, tblSystems.CannotDelete, tblSystems.Password, tblSystems.LastModified, tblSystems.BarsLoadedBeforeTrading, tblSystems.BarsTradedBeforeOrders, tblSystems.Reverify, tblSystems.MMID, tblSystems.Pyramid, tblSystems.TradeDepth, tblSystems.CanPyramidSameEntry, tblSystems.LinkInputs, tblSystems.CheckSum, tblSystems.NumShares, tblSystems.StockExpenses, tblSystems.AllowReverse, tblSystems.ForceLimitThrough, tblSystems.DollarsPerTrade, tblSystems.UseSharesPerTrade, tblSystems.ForexExpenses, tblSystems.IsGuru " & _
            "FROM tblSystems " & _
            "WHERE (((tblSystems.LibraryID)=[SelLibraryID]));"

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cDatabaseUpdates.ApplyVersion54"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion55
'' Description: Add "Override" and "Expose" flags to strategy basket item parameters
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion55()
On Error GoTo ErrSection:

    Dim t As TableDef                   ' Temporary TableDef object
    Dim q As QueryDef                   ' Query into the database
    Dim strName As String               ' Name of the query definition

    If ItemExists(m.DB.TableDefs, "tblStrategyBasketItemParms") Then
        Set t = m.DB.TableDefs("tblStrategyBasketItemParms")
        With t
            If Not ItemExists(.Fields, "IsOverridden") Then
                .Fields.Append .CreateField("IsOverridden", dbBoolean)
                .Fields("IsOverridden").DefaultValue = False
            
                Set q = m.DB.CreateQueryDef("")
                q.SQL = "UPDATE tblStrategyBasketItemParms SET tblStrategyBasketItemParms.IsOverridden = 1;"
                q.Execute
            End If
            If Not ItemExists(.Fields, "IsExposed") Then
                .Fields.Append .CreateField("IsExposed", dbBoolean)
                .Fields("IsExposed").DefaultValue = False
            
                Set q = m.DB.CreateQueryDef("")
                q.SQL = "UPDATE tblStrategyBasketItemParms SET tblStrategyBasketItemParms.IsExposed = 0;"
                q.Execute
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cDatabaseUpdates.ApplyVersion55"
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    ApplyVersion56
'' Description: Add "ExitAtEndOfDay" flag to automated trading items
'' Inputs:      None
'' Returns:     None
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ApplyVersion56()
On Error GoTo ErrSection:

    Dim t As TableDef                   ' Temporary TableDef object
    Dim q As QueryDef                   ' Query into the database

    If ItemExists(m.DB.TableDefs, "tblAutoTradingItem") Then
        Set t = m.DB.TableDefs("tblAutoTradingItem")
        With t
            If Not ItemExists(.Fields, "ExitAtEndOfDay") Then
                .Fields.Append .CreateField("ExitAtEndOfDay", dbBoolean)
                .Fields("ExitAtEndOfDay").DefaultValue = False
            
                Set q = m.DB.CreateQueryDef("")
                q.SQL = "UPDATE tblAutoTradingItem SET tblAutoTradingItem.ExitAtEndOfDay = 0;"
                q.Execute
            End If
        End With
    End If

ErrExit:
    Exit Sub
    
ErrSection:
    RaiseError "cDatabaseUpdates.ApplyVersion56"
    
End Sub

