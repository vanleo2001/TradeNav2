VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGdTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' FEATURES of cGdTree ...
' - works as a tree, or as a simple collection
'    (if just stick with "sibling" relationships)
' - can store multiple object types (if desired),
'    or can check for same object type (default)
' - can optionally store a user-defined "node type"
'    (as a string) with each item
' - auto-creates unique keys when not supplied
' - can get and reset the key for an existing item
' - can get the index of an item or a "relative"
' - can reset the item at an existing position
' - descendants are automatically removed with an item
'    (done in reverse so children are removed before parents)
' - can move a part of the tree (along with all
'    its descendants) to another part of the tree
' - can enumerate through the tree ("For Each x In Tree")
' - efficiency: retrieves keyed items faster than a
'    VB collection, about the same when adding items
'
' DEPENDENCIES ...
' - requires FlexBag.DLL to be registered and in project references
' - requires G32_GD.DLL to be in system folder
' - requires mGdDll.Bas module in project

Option Explicit

' Tree Relationship enumerations
' ("###" means can be used when adding nodes)
Enum eTreeRelationship
    'MYSELF: means the node passed in
    eTREE_Myself = 0       '###
    
    'PARENT: will not exist for 0-level nodes
    eTREE_Parent = 1
    'FIRST CHILD: will exist if any children
    eTREE_FirstChild = 2   '###
    'LAST CHILD: will exist if any children
    eTREE_LastChild = 3    '###
    
    'FIRST SIBLING: always exists (could be node passed in)
    eTREE_FirstSibling = 4 '###
    'LAST SIBLING: always exists (could be node passed in)
    eTREE_LastSibling = 5  '###
    'NEXT SIBLING: will not always exist
    eTREE_NextSibling = 6  '###
    'PREV SIBLING: will not always exist
    eTREE_PrevSibling = 7  '###
    
    'LAST DESCENDANT: always exists (could be node passed in)
    ' (this is the last node which will get
    ' moved or removed along with an ancestor)
    eTREE_LastDescendant = 8
    'ROOT: 0-level ancestor will exist for all
    ' nodes (could be the node passed in)
    eTREE_Root = 9
End Enum

' Module-level data ...
Private Type mPrivate
    ' local variables to hold collections
    ' (using second collection to hold level for
    '  each node -- the two collections are kept
    '  in sync whenever adding, removing, etc.)
    Objects As FlexBag
    NodeInf As FlexBag
    
    ' strObjectType used to store first object type
    ' (other objects must be of the same type ...
    strObjectType As String
    '  ... unless this is set to True)
    bAllowMultipleObjectTypes As Boolean
    
    ' true if strObjectType is an object
    ' (faster method to know if use "Set" or "Let")
    bItemsAreObjects As Boolean
    
    ' gdArray of sorted keys (this is used to check if
    ' a key exists, since FlexBag is very fast for
    ' everything -- except checking for a non-existant
    ' key, where it is almost 100 times slower!)
    hSortedKeys As Long
    
    ' keep local Count for speed
    nCount As Long
End Type
Private m As mPrivate


' Create the collections when this class is created.
Private Sub Class_Initialize()
    
    Randomize '(to seed Rnd used in auto-key)
    
    Set m.Objects = New FlexBag
    m.Objects.AllowDuplicates = False
    m.Objects.AssignmentCreates = False '(we need to handle in the Item property)
    m.Objects.CaseSensitiveKeys = False
    m.Objects.UsePositionalIndices = True
        
    Set m.NodeInf = New FlexBag
    m.NodeInf.AllowDuplicates = m.Objects.AllowDuplicates
    m.NodeInf.AssignmentCreates = m.Objects.AssignmentCreates
    m.NodeInf.CaseSensitiveKeys = m.Objects.CaseSensitiveKeys
    m.NodeInf.UsePositionalIndices = m.Objects.UsePositionalIndices
    
    m.bItemsAreObjects = True '(default)
    m.nCount = 0
    
    m.hSortedKeys = gdCreateArray(eGDARRAY_Strings)
    
End Sub
    
' Destroy collections when this class is terminated.
Private Sub Class_Terminate()
    Clear
    Set m.NodeInf = Nothing
    Set m.Objects = Nothing
    gdDestroyArray m.hSortedKeys
End Sub

' To clear all items from the tree.
Public Sub Clear()
    Dim idx&
    ' remove items in reverse (so children get removed first)
    For idx = m.nCount To 1 Step -1
        m.Objects.Remove idx
        m.NodeInf.Remove idx
    Next
    m.Objects.Clear
    m.NodeInf.Clear
    m.nCount = 0
    m.strObjectType = ""
    gdClear m.hSortedKeys, False
End Sub

' To see if key exists in sorted string array
Private Function KeySearch(ByVal strKey$, nSortedPos&) As Boolean
    Dim hKey&
    hKey = gdCreateString(Len(strKey))
    gdSetStr hKey, 0, strKey
    If gdBinarySearch(m.hSortedKeys, hKey, nSortedPos, eGdSort_IgnoreCase, 0, -1) Then
        KeySearch = True
    Else
        KeySearch = False
    End If
    gdDestroyString hKey
End Function

' Call this to find out if a key (or index) exists.
Public Property Get Exists(ByVal vIndexOrKey As Variant) As Boolean
    ' NOTE:  FlexBag is fast at everything -- EXCEPT when asked
    ' if a non-existant key exists (where it is almost 100 times
    ' slower -- a bug?).  To get around this, we are keeping our
    ' own sorted list of keys just for this purpose.  The
    ' surprisingly small additional overhead to do this when
    ' adding and removing keys is FAR worth it for the benefit
    ' we get here by doing our own existence check rather than
    ' calling FlexBag's "Exists" routine for non-existant keys.
    Dim nPos&
    If VarType(vIndexOrKey) = vbString Then
        Exists = KeySearch(vIndexOrKey, nPos)
    ElseIf vIndexOrKey > 0 And vIndexOrKey <= m.nCount Then
        Exists = True
    Else
        Exists = False
    End If
End Property

' If a key is not supplied, we will make a unique key
' (get a random number that does not yet exist)
Private Function AutoKey(nSortedPos&) As String
    Dim nRandom&, strAutoKey$
    'get random # between 0 and 2 billion
    nRandom = Int(20000 * CDbl(Rnd)) * 100000 + Int(100000 * CDbl(Rnd))
    Do
        nRandom = nRandom + 1
        strAutoKey = "AutoKey# " & CStr(nRandom)
    Loop While KeySearch(strAutoKey, nSortedPos)
    AutoKey = strAutoKey
End Function

Private Sub IndexErrorMsg(ByVal vIndexOrKey As Variant)

    Dim strErr$
    strErr = "Invalid Index:  " & CStr(vIndexOrKey)
    If m.bAllowMultipleObjectTypes = False Then
        strErr = strErr & "|(object:  " & m.strObjectType & ")"
    End If
    InfBox strErr, "e", , "ERROR in cGdTree"

End Sub

' To get the index (position) of an item.
' - returns 0 if index is not valid
Public Property Get Index(ByVal vIndexOrKey As Variant) As Long
    
    On Error GoTo IndexError
    Dim idx&
    If VarType(vIndexOrKey) <> vbString Then
        idx = vIndexOrKey
        If idx < 0 Or idx > m.nCount Then idx = 0
    ElseIf KeySearch(vIndexOrKey, idx) Then
        idx = m.Objects.Index("~" & vIndexOrKey)
    Else
        idx = 0
    End If
IndexResume:
    Index = idx
    Exit Property
    
IndexError:
    IndexErrorMsg vIndexOrKey
    idx = 0
    Resume IndexResume
End Property

' To reference an item in the tree.
' - returns "Nothing" if index is not valid
Public Property Get Item(ByVal vIndexOrKey As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    
    On Error GoTo ItemError
    Dim idx&, iVarType&
    ' get idx for item
    idx = Index(vIndexOrKey)
    If idx > 0 Then
        ' see if need to use "Set" or not
        ' (if not multiple item types, use the local
        ' .bItemsAreObjects variable since it's much
        ' faster than calling "VarType" for each item)
        If m.bAllowMultipleObjectTypes Then
            ' need to check each item
            iVarType = VarType(m.Objects(idx))
            If iVarType = vbObject Or iVarType <= vbNull Then
                Set Item = m.Objects(idx)
            Else
                Item = m.Objects(idx)
            End If
        ElseIf m.bItemsAreObjects Then
            Set Item = m.Objects(idx)
        Else
            Item = m.Objects(idx)
        End If
    Else
        Set Item = Nothing
    End If
ItemResume:
    Exit Property

ItemError:
    IndexErrorMsg vIndexOrKey
    Set Item = Nothing
    Resume ItemResume
End Property

' To set an item in the tree (can also set to "Nothing").
Public Property Let Item(ByVal vIndexOrKey As Variant, ByVal objSet As Variant)
    
    Dim idx&
    ' verify object type
    If BadObjectType(objSet) Then Exit Property
        
    If Exists(vIndexOrKey) Then
        ' if exists, just reset its item
        idx = Index(vIndexOrKey)
        If idx > 0 Then m.Objects(idx) = objSet
    ElseIf VarType(vIndexOrKey) = vbString Then
        ' if not exist, then just add it
        Add objSet, CStr(vIndexOrKey)
    Else
        ' add at index# (fill in extra space first)
        For idx = m.nCount + 1 To vIndexOrKey - 1
            Add Nothing
        Next
        Add objSet
    End If

End Property
'(just have Set Item call Let Item -- so can do either)
Public Property Set Item(ByVal vIndexOrKey As Variant, ByVal objSet As Variant)
    Item(vIndexOrKey) = objSet
End Property

Private Function FindAddAfterRelative(ByVal idxRelative&, ByVal eAsRelation As eTreeRelationship, nAddLevel As Long) As Long
    
    Dim nRelativeLevel&, idxAdd&, idx&
    
    If idxRelative > 0 And idxRelative <= m.nCount Then
        nRelativeLevel = NodeLevel(idxRelative)
        Select Case eAsRelation
            Case eTREE_FirstChild:
                ' add as the first child
                ' (i.e. next position)
                idxAdd = idxRelative + 1
                nAddLevel = nRelativeLevel + 1 ' child
            Case eTREE_LastChild:
                ' add as the last child
                ' (after current last descendant)
                idxAdd = RelativeIndex(idxRelative, eTREE_LastDescendant) + 1
                nAddLevel = nRelativeLevel + 1 ' child
            Case eTREE_FirstSibling:
                ' add as first sibling
                ' (right before current first sibling)
                idxAdd = RelativeIndex(idxRelative, eTREE_FirstSibling)
                nAddLevel = nRelativeLevel ' sibling
            Case eTREE_LastSibling:
                ' add as last sibling
                ' (after current last sibling's last descendant)
                idx = RelativeIndex(idxRelative, eTREE_LastSibling)
                idxAdd = RelativeIndex(idx, eTREE_LastDescendant) + 1
                nAddLevel = nRelativeLevel ' sibling
            Case eTREE_NextSibling:
                ' add as next sibling
                ' (after my last descendant)
                idxAdd = RelativeIndex(idxRelative, eTREE_LastDescendant) + 1
                nAddLevel = nRelativeLevel ' sibling
            Case eTREE_PrevSibling, eTREE_Myself:
                ' add as previous sibling
                ' (i.e. previous position)
                idxAdd = idxRelative
                nAddLevel = nRelativeLevel ' sibling
            Case Else:
                'bogus relationship for adding!
                idxAdd = 0 ' i.e. not added
        End Select
    End If

    FindAddAfterRelative = idxAdd
End Function


' To add an existing item to the tree.
' (returns current index of new item)
Public Function Add(ByVal objAdd As Variant, _
        Optional ByVal strKey As String = "", _
        Optional ByVal vRelativesIndexOrKey As Variant = -1, _
        Optional ByVal eAsRelation As eTreeRelationship = eTREE_NextSibling) As Long
    
    Dim idxAdd&, nAddLevel&, nSortedPos&
    Dim bFindRelative As Boolean
    
    ' verify object type
    If BadObjectType(objAdd) Then
        Add = 0
        Exit Function
    End If
    
    ' defaults
    idxAdd = m.nCount + 1 ' at end of tree
    nAddLevel = 0   ' root node (level = 0)
                
    ' see if valid relative
    bFindRelative = True
    If m.nCount = 0 Then
        bFindRelative = False
    ElseIf VarType(vRelativesIndexOrKey) <> vbString Then
        If vRelativesIndexOrKey < 0 Or _
            vRelativesIndexOrKey > m.nCount Then
                ' just add to end of tree
                bFindRelative = False
        ElseIf vRelativesIndexOrKey = 0 Then
            ' add to beginning of tree
            idxAdd = 1
            bFindRelative = False
        End If
    End If
    ' find where to add based on relationship
    If bFindRelative Then
        idxAdd = FindAddAfterRelative(Index(vRelativesIndexOrKey), _
                    eAsRelation, nAddLevel)
        If idxAdd = 0 Then
            'bogus relationship for adding!
            Add = 0 ' i.e. not added
            Exit Function
        End If
    End If
        
    ' get sorted position for new key
    If Len(strKey) = 0 Then
        ' make a key if one was not given
        strKey = AutoKey(nSortedPos)
    ElseIf KeySearch(strKey, nSortedPos) Then
        ' duplicate key exists!
        Add = 0
        Exit Function
    End If
    ' insert key into sorted array (see Note on "Exists")
    If nSortedPos >= gdGetSize(m.hSortedKeys) Then
        gdSetStr m.hSortedKeys, nSortedPos, strKey
    Else
        gdInsertStr m.hSortedKeys, strKey, nSortedPos
    End If

    ' Keys don't seem to work if they are a string of numbers
    ' (a Flexbag bug?).  We prepend a non-numeric character to
    ' all keys (just for internal use) which fixes this problem.
    strKey = "~" & strKey

    ' add object and level to collection
    If idxAdd <= m.nCount Then
        ' insert before existing item at that position
        m.Objects.Add objAdd, strKey, idxAdd
        m.NodeInf.Add Trim(Str(nAddLevel)), strKey, idxAdd
    Else
        ' just append to end
        m.Objects.Add objAdd, strKey
        m.NodeInf.Add CStr(nAddLevel), strKey
    End If
    m.nCount = m.Objects.Count
    
    Add = idxAdd 'return index where inserted
End Function

' Returns index of a Relative
' (returns 0 if does not exist)
Public Property Get RelativeIndex(ByVal vIndexOrKeyFrom As Variant, _
        ByVal eRelative As eTreeRelationship) As Long
 
    Dim idxRelative&, idx&, nLevel&
    
    idxRelative = 0 ' default: 0 = not exists
    ' find specified relative
    idx = Index(vIndexOrKeyFrom)
    If idx > 0 Then
        nLevel = Val(m.NodeInf(idx))
        Select Case eRelative
            Case eTREE_Myself:
                idxRelative = idx
                
            'PARENT: will exist for all nodes
            ' except for 0-level nodes
            Case eTREE_Parent:
                ' search backwards for a lower level
                If nLevel > 0 Then
                    For idx = idx - 1 To 1 Step -1
                        If Val(m.NodeInf(idx)) < nLevel Then
                            idxRelative = idx
                            Exit For
                        End If
                    Next
                End If
            
            'ROOT: will exist for all nodes
            ' (could be the node passed in)
            Case eTREE_Root:
                ' search backward for first 0-level
                For idx = idx To 1 Step -1
                    If Val(m.NodeInf(idx)) = 0 Then
                        idxRelative = idx
                        Exit For
                    End If
                Next
            
            'FIRST CHILD: if children exist
            Case eTREE_FirstChild:
                ' just check next node (see if higher level)
                If idx + 1 <= m.nCount Then
                    If Val(m.NodeInf(idx + 1)) = nLevel + 1 Then
                        idxRelative = idx + 1
                    End If
                End If
            
            'LAST CHILD: if children exist
            Case eTREE_LastChild:
                ' search forward for last node which is
                ' one level higher, until hit same level
                For idx = idx + 1 To m.nCount
                    If Val(m.NodeInf(idx)) = nLevel + 1 Then
                        idxRelative = idx
                    ElseIf Val(m.NodeInf(idx)) <= nLevel Then
                        Exit For
                    End If
                Next
            
            'LAST DESCENDANT: will always exist
            ' (could be node passed in)
            Case eTREE_LastDescendant:
                ' search for last node at a higher level
                ' (this is the last node which will get
                ' moved or removed along with an ancestor)
                idxRelative = idx
                For idx = idx + 1 To m.nCount
                    If Val(m.NodeInf(idx)) > nLevel Then
                        idxRelative = idx
                    Else
                        Exit For
                    End If
                Next
            
            'FIRST SIBLING: will always exist
            ' (could be node passed in)
            Case eTREE_FirstSibling:
                ' search backwards for furthest back node
                ' at the same level, until hit lower level
                For idx = idx To 1 Step -1
                    If Val(m.NodeInf(idx)) = nLevel Then
                        idxRelative = idx
                    ElseIf Val(m.NodeInf(idx)) < nLevel Then
                        Exit For
                    End If
                Next
            
            'LAST SIBLING: will always exist
            ' (could be node passed in)
            Case eTREE_LastSibling:
                ' search forward for last node at the
                ' same level, until hit lower level
                For idx = idx To m.nCount
                    If Val(m.NodeInf(idx)) = nLevel Then
                        idxRelative = idx
                    ElseIf Val(m.NodeInf(idx)) < nLevel Then
                        Exit For
                    End If
                Next
            
            'NEXT SIBLING: will not always exist
            Case eTREE_NextSibling:
                ' search forward for next node at same
                ' level, until hit lower level
                For idx = idx + 1 To m.nCount
                    If Val(m.NodeInf(idx)) = nLevel Then
                        idxRelative = idx
                        Exit For
                    ElseIf Val(m.NodeInf(idx)) < nLevel Then
                        Exit For
                    End If
                Next
            
            'PREV SIBLING: will not always exist
            Case eTREE_PrevSibling:
                ' search backward for next node at
                ' same level, until hit lower level
                For idx = idx - 1 To 1 Step -1
                    If Val(m.NodeInf(idx)) = nLevel Then
                        idxRelative = idx
                        Exit For
                    ElseIf Val(m.NodeInf(idx)) < nLevel Then
                        Exit For
                    End If
                Next
        End Select
    End If

    RelativeIndex = idxRelative
End Property

' Returns a Relative (returns "Nothing" if does not exist)
Public Property Get RelativeItem(ByVal vIndexOrKeyFrom As Variant, _
        ByVal eRelative As eTreeRelationship) As Variant
    
    Dim idx&, iVarType&
    idx = RelativeIndex(vIndexOrKeyFrom, eRelative)
    If idx > 0 Then
        ' see if need to use "Set" or not
        If m.bAllowMultipleObjectTypes Then
            ' need to check each item
            iVarType = VarType(m.Objects(idx))
            If iVarType = vbObject Or iVarType <= vbNull Then
                Set RelativeItem = m.Objects(idx)
            Else
                RelativeItem = m.Objects(idx)
            End If
        ElseIf m.bItemsAreObjects Then
            Set RelativeItem = m.Objects(idx)
        Else
            RelativeItem = m.Objects(idx)
        End If
    Else
        Set RelativeItem = Nothing
    End If

End Property

' Returns index of the Ancestor at a specified level
' - if nAtLevel >= 0: find ancestor at this level (absolute)
' - if nAtLevel < 0: find ancestor this many levels from me (relative)
' (returns 0 if does not exist)
Public Property Get AncestorIndex(ByVal vIndexOrKeyFrom As Variant, _
        ByVal nAtLevel As Long) As Long
 
    Dim idxAncestor As Long, idx As Long
    
    idxAncestor = 0 ' default: 0 = not exists
    ' find specified relative
    idx = Index(vIndexOrKeyFrom)
    If idx > 0 Then
        If nAtLevel < 0 Then
            ' specified relative to current level
            nAtLevel = Val(m.NodeInf(idx)) + nAtLevel
        End If
        If nAtLevel <= Val(m.NodeInf(idx)) And nAtLevel >= 0 Then
            ' search backward for first node at level
            For idx = idx To 1 Step -1
                If Val(m.NodeInf(idx)) = nAtLevel Then
                    idxAncestor = idx
                    Exit For
                End If
            Next
        End If
    End If

    AncestorIndex = idxAncestor
End Property

' To remove an item from the tree.
' (removes all descendants in reverse, so
'  children get removed before their parents)
Public Function Remove(ByVal vIndexOrKey As Variant) As Boolean
    
    Dim idx&, idxRemove&, idxLastDesc&, nSortedPos&, strKey$
    
    ' find item (make sure item exists)
    nSortedPos = -1
    If VarType(vIndexOrKey) = vbString Then
        If KeySearch(vIndexOrKey, nSortedPos) Then
            idxRemove = Index(vIndexOrKey)
        End If
    ElseIf vIndexOrKey > 0 And vIndexOrKey <= m.nCount Then
        idxRemove = vIndexOrKey
    End If

    If idxRemove > 0 Then
        ' get number to remove: with all descendants
        idxLastDesc = RelativeIndex(idxRemove, eTREE_LastDescendant)
        If idxLastDesc <> idxRemove Then nSortedPos = -1 '(must research for first key)
        
        ' remove all objects in reverse
        For idx = idxLastDesc To idxRemove Step -1
            ' find each key, then remove
            strKey = Mid(m.Objects.Key(idx), 2)
            If nSortedPos < 0 Then '(may already have it from beginning of routine)
                If Not KeySearch(strKey, nSortedPos) Then
                    nSortedPos = -1
                End If
            End If
            If nSortedPos >= 0 Then
                gdDeleteItems m.hSortedKeys, nSortedPos, 1
            End If
            nSortedPos = -1
            ' remove objects
            m.Objects.Remove idx
            m.NodeInf.Remove idx
        Next
        m.nCount = m.Objects.Count
        idx = gdGetSize(m.hSortedKeys)
        
        Remove = True
    Else
        Remove = False ' not found
    End If

End Function

' To get the number of items in the tree.
Public Property Get Count() As Long
    Count = m.nCount
End Property

' To get the key for an item.
Public Property Get Key(ByVal vIndexOrKey As Variant) As String
    'strip prefix (first character) off the internal key
    Dim idx&
    idx = Index(vIndexOrKey)
    If idx > 0 Then
        Key = Mid(m.Objects.Key(idx), 2)
    Else
        Key = ""
    End If
End Property

' To set the key for an item.
Public Property Let Key(ByVal vIndexOrKey As Variant, ByVal strNewKey As String)
    
    Dim idx&, nSortedPos&, strOldKey$
    
    idx = Index(vIndexOrKey)
    If idx > 0 Then
        strOldKey = Mid(m.Objects.Key(idx), 2)
        If strOldKey <> strNewKey Then
            ' make sure new key doesn't already exist
            If Len(strNewKey) = 0 Then
                strNewKey = AutoKey(nSortedPos)
            ElseIf KeySearch(strNewKey, nSortedPos) Then
                nSortedPos = -1
            End If
            If nSortedPos >= 0 Then
                ' add new key to sorted list
                gdInsertStr m.hSortedKeys, strNewKey, nSortedPos
                ' remove old key from sorted list
                If KeySearch(strOldKey, nSortedPos) Then
                    gdDeleteItems m.hSortedKeys, nSortedPos, 1
                End If
                ' change key in collections
                m.Objects.Key(idx) = "~" & strNewKey
                m.NodeInf.Key(idx) = "~" & strNewKey
            End If
        End If
    End If
    
End Property

' Returns true if tree has been modified since last queried.
Public Property Get Modified() As Boolean
    Modified = m.Objects.Modified
End Property

' This hidden property allows you to enumerate
' the tree items with the For...Each syntax.
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m.Objects.[_NewEnum]
End Property

' to adjust node levels for block of nodes
Private Sub AdjustNodeLevels(ByVal idxStart&, ByVal idxEnd&, ByVal nLevelAdjust&)
    Dim idx&, iPos&, nLevel&, strInf$
    If nLevelAdjust <> 0 Then
        For idx = idxStart To idxEnd
            strInf = m.NodeInf(idx)
            nLevel = Val(strInf) + nLevelAdjust
            ' keep current ObjectType with new Node Level
            iPos = InStr(strInf, Chr(0))
            If iPos > 0 Then
                m.NodeInf(idx) = CStr(nLevel) & Chr(0) & Mid(strInf, iPos + 1)
            Else
                m.NodeInf(idx) = CStr(nLevel)
            End If
        Next
    End If
End Sub

' Returns 0-based Level of specified node
' (returns -1 for an invalid index)
Public Property Get NodeLevel(ByVal vIndexOrKey As Variant) As Long
    
    Dim idx&
    idx = Index(vIndexOrKey)
    If idx > 0 Then
        NodeLevel = Val(m.NodeInf(idx))
    Else
        ' invalid index
        NodeLevel = -1 '0
    End If

End Property

' Sets node level for an item (but only if in valid range
' -- cannot move a node more than one level higher than
' the previous node), and adjusts for all its descendants.
Public Property Let NodeLevel(ByVal vIndexOrKey As Variant, ByVal nLevel As Long)
    
    Dim idx&, idxLast&, nLevelAdjust&
    idx = Index(vIndexOrKey)
    If idx > 0 Then
        If idx = 1 Then
            nLevel = 0
        ElseIf nLevel > Val(m.NodeInf(idx - 1)) + 1 Then
            ' invalid node level
            Exit Property
        End If
        ' adjust node level for all descendants
        nLevelAdjust = nLevel - Val(m.NodeInf(idx))
        If nLevelAdjust <> 0 Then
            idxLast = RelativeIndex(idx, eTREE_LastDescendant)
            AdjustNodeLevels idx, idxLast, nLevelAdjust
        End If
    End If

End Property

' Returns the node type for an item in the tree.
Public Property Get NodeType(ByVal vIndexOrKey As Variant) As String
    
    Dim s$, i&, idx&
    idx = Index(vIndexOrKey)
    If idx > 0 Then
        s = m.NodeInf(idx)
        i = InStr(s, Chr(0))
        If i > 0 Then
            s = Mid(s, i + 1)
        Else
            s = ""
        End If
    End If
    
    NodeType = s
End Property

' Sets the node type for an item in the tree.
Public Property Let NodeType(ByVal vIndexOrKey As Variant, ByVal strType As String)
    
    Dim nLevel&, idx&
    idx = Index(vIndexOrKey)
    If idx > 0 Then
        nLevel = Val(m.NodeInf(idx))
        m.NodeInf(idx) = CStr(nLevel) & Chr(0) & strType
    End If

End Property

' Returns true if node has children
Public Function HasChildren(ByVal vIndexOrKey As Variant) As Boolean
    Dim idx&
    idx = Index(vIndexOrKey)
    If idx > 0 And idx < m.nCount Then '(don't look at last item since can't have children)
        If NodeLevel(idx + 1) > NodeLevel(idx) Then
            HasChildren = True
        End If
    End If
End Function

' To move an item and all its descendants (along
' with their keys) to another location in the tree.
' - returns index where item got moved to
' - returns 0 if an invalid move (e.g. trying to move
'       a node into the middle of its own descendants)
Public Function Move(ByVal vToMove As Variant, _
        ByVal vRelativesIndexOrKey As Variant, _
        ByVal eAsRelation As eTreeRelationship) As Long

    Dim idxOld&, idxNew&, i&, nNumMove&, nLevel&
    Dim strMovedKey$, strKey$
    
    ' get number to move, where from and where to
    idxOld = Index(vToMove)
    If idxOld > 0 Then
        strMovedKey = Key(idxOld) '(store for later)
        nNumMove = RelativeIndex(idxOld, eTREE_LastDescendant) _
                - idxOld + 1
        ' where to move
        idxNew = Index(vRelativesIndexOrKey)
        idxNew = FindAddAfterRelative(idxNew, eAsRelation, nLevel)
    End If
    
    ' make sure a valid move (e.g. can't move part of a
    ' tree into the middle of its own descendants!)
    If nNumMove > 0 And idxNew > 0 And _
        (idxNew <= idxOld Or idxNew >= idxOld + nNumMove) Then

        ' adjust the node levels for the block
        AdjustNodeLevels idxOld, idxOld + nNumMove - 1, _
                nLevel - NodeLevel(idxOld)

        If idxNew < idxOld Or idxNew > idxOld + nNumMove Then
            ' move items in collections
            For i = 1 To nNumMove
                ' save key, then remove key from nodes (so no duplicates)
                strKey = m.Objects.Key(idxOld)
                m.Objects.RemoveKey idxOld
                m.NodeInf.RemoveKey idxOld
                ' add to new position with original key
                If idxNew > idxOld Then
                    ' moving towards end
                    m.Objects.Add m.Objects(idxOld), strKey, , idxNew - 1
                    m.NodeInf.Add m.NodeInf(idxOld), strKey, , idxNew - 1
                    ' remove from old position
                    m.Objects.Remove idxOld
                    m.NodeInf.Remove idxOld
                Else
                    ' moving towards beginning
                    m.Objects.Add m.Objects(idxOld), strKey, idxNew
                    m.NodeInf.Add m.NodeInf(idxOld), strKey, idxNew
                    ' (need to adjust indexes since just added below)
                    idxOld = idxOld + 1
                    idxNew = idxNew + 1
                    ' remove from old position
                    m.Objects.Remove idxOld
                    m.NodeInf.Remove idxOld
                End If
            Next
        End If
        
        ' return index of where got moved to
        Move = Index(strMovedKey)
    Else
        Move = 0
    End If
    
End Function

' To get and set the "AllowMultipleObjectTypes" property
Public Property Get AllowMultipleObjectTypes() As Boolean
    AllowMultipleObjectTypes = m.bAllowMultipleObjectTypes
End Property
Public Property Let AllowMultipleObjectTypes(ByVal bAllowed As Boolean)
    m.bAllowMultipleObjectTypes = bAllowed
End Property

' Checks the object type (if multiple object types are not allowed)
Private Function BadObjectType(Obj As Variant) As Boolean
    
    Dim strObjectType$, bError As Boolean
    If Not m.bAllowMultipleObjectTypes Then
        strObjectType = TypeName(Obj)
        If UCase(strObjectType) = "NOTHING" Then
            ' just ignore it, it's OK
        ElseIf Len(m.strObjectType) = 0 Then
            ' if first object added, store object type
            m.strObjectType = strObjectType
            If VarType(Obj) = vbObject Then
                m.bItemsAreObjects = True
            Else
                m.bItemsAreObjects = False
            End If
        ElseIf strObjectType <> m.strObjectType Then
            ' Error: not same object type as first one added
            bError = True
            Err.Raise vbObjectError + 1, "Class: cGdTree", _
                "Trying to add multiple object types to cGdTree:" & Chr(13) & m.strObjectType & "  and  " & strObjectType
        End If
    End If
    
    BadObjectType = bError
End Function

' Helper for making a copy
Friend Sub SetPrivateData(mCopyFrom As mPrivate)
    Dim i&, hArray&
    ' get copy of all private data (but save and restore array handle)
    hArray = m.hSortedKeys
    m = mCopyFrom
    m.hSortedKeys = hArray
    ' make this a copy of the gdArray
    gdCopy m.hSortedKeys, mCopyFrom.hSortedKeys
    ' now make a copy of the collections
    Set m.Objects = New FlexBag
    Set m.NodeInf = New FlexBag
    For i = 1 To mCopyFrom.Objects.Count
        m.Objects.Add mCopyFrom.Objects(i), mCopyFrom.Objects.Key(i)
        m.NodeInf.Add mCopyFrom.NodeInf(i), mCopyFrom.NodeInf.Key(i)
    Next
    m.nCount = m.Objects.Count
End Sub

' To return a true "copy" of this object
Public Function MakeCopy() As cGdTree
    Dim aCopy As New cGdTree
    aCopy.SetPrivateData m
    Set MakeCopy = aCopy
End Function


Public Function ErrorCheck() As Long

    Dim iErr&, i&, strKey$
    
    If m.nCount <> m.Objects.Count Then
        iErr = 1
    ElseIf m.nCount <> m.NodeInf.Count Then
        iErr = 2
    ElseIf m.nCount <> gdGetSize(m.hSortedKeys) Then
        iErr = 3
        i = gdGetSize(m.hSortedKeys)
    Else
        For i = 1 To m.nCount
            strKey = gdGetStr(m.hSortedKeys, i - 1)
            If Not m.Objects.Exists("~" & strKey) Then
                iErr = 4
                Exit For
            End If
        Next
    End If

    ErrorCheck = iErr
End Function

Public Sub FromKeyValueString(ByVal strKeyValueString As String, Optional ByVal strFieldDelimiter As String = vbTab, Optional ByVal strKeyValueDelimiter As String = "=")

    Dim astrFields As cGdArray          ' Array of fields from the main string
    Dim lIndex As Long                  ' Index into a for loop
    Dim iIndexOfDelim As Integer        ' Index of the equals sign in the string
    Dim strKey As String                ' Key for the string
    Dim strValue As String              ' Value for the string
    
    Set astrFields = New cGdArray
    astrFields.SplitFields strKeyValueString, strFieldDelimiter
    
    For lIndex = 0 To astrFields.Size - 1
        iIndexOfDelim = InStr(astrFields(lIndex), strKeyValueDelimiter)
        If iIndexOfDelim = 0 Then
            strKey = astrFields(lIndex)
            strValue = ""
        Else
            strKey = Left(astrFields(lIndex), iIndexOfDelim - 1)
            strValue = Mid(astrFields(lIndex), iIndexOfDelim + 1)
        End If
        
        If Exists(strKey) Then
            Item(strKey) = strValue
        Else
            Add strValue, strKey
        End If
    Next lIndex
    
End Sub
