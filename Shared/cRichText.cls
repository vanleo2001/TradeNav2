VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRichText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'BENEFITS ...
'- MUCH faster for formatting lengthy text
'- very efficient for appending and inserting new text
'- easy to assign formatting at the time text is added
'- builds RTF text only when need to assign to control
'- flexibility to allow all types of rich text formatting
'   (allows for common presets as well as custom formats)

'USAGE ...
'1) Set the .RTB member to the RichTextBox control
'2) Use .AddText to add each section of text with its formatting
'   (formatting can either be a combination of presets or a custom format)
'   e.g.  .AddText "Bold and blue text", rtfBold Or rtfBlue
'         .AddText " italicized text to add", rtfItalic
'3) Call .BuildRTF to build the TextRTF and assign it to the control

'Custom Color (one of the presets which can be customized) ...
'- To set rtfCustomColor: .CustomColor = RGB(???,???,???)
'- To use it: .AddText "Bold and Colored weird", rtfCustomColor Or rtfBold

'Custom Formatting (optional) ...
'- Can create up to 127 custom formats, each of which can store whatever
'   rich text formatting you want -- custom formats cannot, however,
'   be combined with the presets (i.e. can't do "FormatID Or rtfXXX")
'- To create custom formats:
'   1) call .CreateCustomFormat FormatID  (where FormatID < 128)
'   2) then immediately do the RTB custom formatting for this ID
'   e.g.    .CreateCustomFormat 3  'for third custom format
'           With .RTBox
'               .SetBold = True
'               .SetStrikeThru = True
'               .SetColor = RGB(121,205,33) 'my weird color
'           End With
'- To use custom formats: format = "rtfUseCustomFormat + FormatID"
'   e.g.  .AddText "Custom formatted", rtfUseCustomFormat + 3
'- Can select one custom format to use as a preset (so
'       it can be freely combined with the other presets):
'   1) create the custom format (e.g.  .CreateCustomFormat 3)
'   2) assign that Custom FormatID to be the CustomPreset
'       (e.g.  .CustomFormatIDforCustomPreset = 3)
'   3) now can combine that custom format with other presets
'       (e.g.  .AddText "custom bold", rtfCustomPreset Or rtfBold)


'========================================================
'IMPLEMENTATION DESIGN ...
'- maintain Text and Formatting strings (kept same length),
'   where each byte of Formatting string is the FormatID (0-255)
'   for the corresponding byte of the Text string
'- use FormatID's as follows:  0 to 127 is used as a bitmask of 7
'   common presets, while 128 to 255 are available for custom formats
'- the rtf for each FormatID is generated from the control itself
'   (rather than hard-coded) and stored in private FormatsRTF string
'- to build TextRTF:
'   1) build array of FormatID's (for each FormatID, get
'       rtf for that format from the FormatsRTF string)
'   2) start TextRtf text with same header as FormatsRTF
'   3) add section of text (with same formatting) at a time
'       from Text to TextRTF (prepend the formatting)
'   4) assign TextRTF to property of control
Option Explicit

Private Const CLASS_VERSION = 2 '(for serializing class data)
Private Const RTB_CHAR = "`"

'Constants used as presets (can be "OR'd" with each other)
Enum eRTF_Formats
    rtfUnchanged = -1  'same format as previous character
    rtfPlain = 0  'black, non-bold, non-italic
    rtfBold = 1
    rtfItalic = 2
    rtfUnderline = 4
    rtfBlue = 8
    rtfRed = 16
    rtfCustomColor = 32
    rtfCustomPreset = 64  'to combine one Custom Format ID with other presets
    rtfUseCustomFormat = 128 'cannot be OR'd with other presets
End Enum

'Constants used for .SetFormat
Enum eRTF_FormatMode
    rtfTurnOff = 0   'turns off the specified formats
    rtfTurnOn = 1    'adds the specified formats to existing
    rtfOverwrite = 2 'replaces existing with the specified formats
End Enum

'RichTextBox associated with this class
Private mRTB As RichTextBox

'Stores the text and its formatting (both are kept same length)
Private mText As cMemBuffer
Private mFormatting As cMemBuffer

'Local storage of FormatsRTF (string containing special
'formatting codes for each FormatID)
Private mFormatsRTF As String
Private mbFormatsRTFLoaded As Boolean

'for customizing
Private mCustomColor As Long
Private mCustomPreset As Long 'CustomFormatID to use as preset
Private mMaxID As Long  'keeps track of highest custom format set

'Constructor
Private Sub Class_Initialize()
    Set mText = New cMemBuffer
    Set mFormatting = New cMemBuffer
    mCustomColor = RGB(0, 0, 128) '(default to dark blue)
    mCustomPreset = 0 'default
End Sub

'Destructor
Private Sub Class_Terminate()
    Set mRTB = Nothing
    Set mText = Nothing
    Set mFormatting = Nothing
End Sub

'Returns RichTextBox control associated with this class
Public Property Get RTBox() As RichTextBox
    Set RTBox = mRTB
End Property

'Sets RichTextBox control associated with this class
Public Property Let RTBox(ByVal pRTBox As RichTextBox)
    Set mRTB = pRTBox
    UpdateFormatsRTF
End Property
Public Property Set RTBox(ByVal pRTBox As RichTextBox)
    Let RTBox = pRTBox
End Property

'Clears the current text and formatting
Public Sub Clear()
    mText.Clear
    mFormatting.Clear
End Sub

'Returns length of current text
Public Property Get TextLength() As Long
    TextLength = mText.Length
End Property

'Returns the current text
Public Property Get Text() As String
    Text = mText.Buffer
End Property

'Appends or inserts a section of text along with its formatting
'(nAtOffset is zero-based, set to -1 to append to existing text)
Public Sub AddText(ByVal strText$, _
        Optional ByVal eFormats As eRTF_Formats = rtfPlain, _
        Optional ByVal nAtOffset& = -1)

    Dim i&
    If nAtOffset = -1 Or nAtOffset > mText.Length Then
        nAtOffset = mText.Length 'append to end
    End If
    If nAtOffset >= 0 Then
        If eFormats = rtfUnchanged Then
            If nAtOffset = 0 Then
                eFormats = rtfPlain 'just set to plain
            Else
                'get format of previous character
                eFormats = mFormatting.GetByte(nAtOffset - 1)
            End If
        End If
        'insert/append text
        mText.Insert strText, nAtOffset
        'insert/append formatting string (same length as text)
        'where each character indicates the "FormatID" for the text
        mFormatting.Insert String(Len(strText), eFormats), nAtOffset
    End If

End Sub

'Returns the color of rtfCustomColor
Public Property Get CustomColor() As Long
    CustomColor = mCustomColor
End Property
'Sets the color of rtfCustomColor
Public Property Let CustomColor(ByVal nNewColor As Long)
    mCustomColor = nNewColor
End Property

'Returns the CustomFormatID assigned to rtfCustomPreset
Public Property Get CustomFormatIDforCustomPreset() As Long
    CustomFormatIDforCustomPreset = mCustomPreset
End Property
'Sets the CustomFormatID to use for rtfCustomPreset
Public Property Let CustomFormatIDforCustomPreset(ByVal nNewID As Long)
    If nNewID >= 0 And nNewID <= 255 Then
        '(store as 0-127 ... 128 will be added to it later)
        If nNewID >= 128 Then nNewID = nNewID - 128
        mCustomPreset = nNewID
    End If
End Property

'To create a custom format (call this, then set formatting of RTBox)
Public Sub CreateCustomFormat(ByVal CustomFormatID As Integer)
    'convert ID to number between 128 and 255
    If CustomFormatID < 128 Then CustomFormatID = CustomFormatID + 128
    If CustomFormatID >= 128 And CustomFormatID <= 255 Then
        UpdateFormatsRTF
        With mRTB
            .SelStart = (CustomFormatID - 124) * 2
            .SelLength = 1
            If CustomFormatID > mMaxID Then mMaxID = CustomFormatID
        End With
    End If
End Sub

'Returns the format mask at the specified offset
Public Property Get Format(ByVal nAtOffset&) As eRTF_Formats
    If nAtOffset >= 0 And nAtOffset < mFormatting.Length Then
        Format = mFormatting.GetByte(nAtOffset)
    Else
        Format = 0
    End If
End Property
'Sets the format mask at the specified offset
Public Property Let Format(ByVal nAtOffset&, ByVal vFormat As eRTF_Formats)
    If nAtOffset >= 0 And nAtOffset < mFormatting.Length _
        And vFormat >= 0 And vFormat <= 255 Then
            mFormatting.PutByte vFormat, nAtOffset
    End If
End Property

'Sets the formatting for a section of text -- "mode" can be ...
'- rtfTurnOn: combines the specified format(s) with existing formats
'- rtfTurnOff: sets just the specified format(s) off for existing formats
'- rtfOverwrite: replaces existing formats with specified format(s)
Public Sub SetFormat(ByVal eFormats As eRTF_Formats, ByVal eFormatMode As eRTF_FormatMode, _
        ByVal nStartAtOffset&, Optional ByVal nLength& = -1)

    Dim iPos&, buf$
    Dim CurMask As Byte, SetMask As Byte, NewMask As Byte
    With mFormatting
        ' fix length
        If nLength = -1 Or nStartAtOffset + nLength > .Length Then
            nLength = .Length - nStartAtOffset
        End If
        If nStartAtOffset >= 0 And nStartAtOffset < .Length And nLength > 0 _
                And eFormats >= 0 And eFormats <= 255 Then
            SetMask = eFormats
            buf = .GetStr(nLength, nStartAtOffset)
            For iPos = 1 To Len(buf)
                CurMask = Asc(Mid(buf, iPos, 1))
                Select Case eFormatMode
                    Case rtfTurnOn:
                        'set bits in mask
                        NewMask = CurMask Or SetMask
                    Case rtfTurnOff:
                        'clear bits from mask
                        NewMask = CurMask And Not SetMask
                    Case rtfOverwrite:
                        'set format to this
                        NewMask = SetMask
                End Select
                If NewMask <> CurMask Then
                    Mid(buf, iPos, 1) = Chr(NewMask)
                End If
            Next
            .PutStr buf, nStartAtOffset
        End If
    End With
End Sub


'Update local copy of FormatsRTF (string with codes for each FormatID)
Private Sub UpdateFormatsRTF()
    
    Dim strRTF$, i&
    
    If Not mRTB Is Nothing Then
        With mRTB
            'make sure FormatsRTF is loaded in control
            If Not mbFormatsRTFLoaded Then
                .TextRTF = mFormatsRTF
                mbFormatsRTFLoaded = True
            End If
    
            'initialize presets, if not already done
            If Len(mFormatsRTF) = 0 Then
                'must first make sure RTBox is cleared
                '(just put "\plain" into rich text)
                .Text = ""
                strRTF = .TextRTF
                i = InStr(LCase(strRTF), "\pard\")
                If i > 0 Then
                    .TextRTF = Left(strRTF, i + 4) & "\plain }"
                End If
                
                'now put enough of special characters into string
                .Text = String(512, RTB_CHAR) '(one character for each FormatID)
                '.SelStart = rtfBold * 2
                '.SelLength = 1
                '.SelBold = True
                '.SelStart = rtfItalic * 2
                '.SelLength = 1
                '.SelItalic = True
                '.SelStart = rtfUnderline * 2
                '.SelLength = 1
                '.SelUnderline = True
                .SelStart = 2 'rtfBlue * 2
                .SelLength = 1
                .SelColor = vbBlue
                .SelStart = 4 'rtfRed * 2
                .SelLength = 1
                .SelColor = vbRed
            End If
            
            'make sure custom color is set correctly
            .SelStart = 6 ' rtfCustomColor * 2
            .SelLength = 1
            .SelColor = mCustomColor
            
            'update local copy of FormatsRTF
            mFormatsRTF = .TextRTF
            .SelLength = 0
        End With
    End If

End Sub

'Load array with the formatting rtf for each FormatID (0-255)
Private Sub LoadFormatsArray(aFormats() As String, strHdr$, strPlain$)
    
    Dim strRTF$, iFormat&, iPreset&, strTextRTF$
    Dim iStart&, iPos&
    
    strHdr = ""
    strPlain = ""
    
    'Make sure FormatsRTF is updated and loaded into control
    UpdateFormatsRTF
    
    'Size array to store RTF for each FormatID
    If mMaxID < 127 Then mMaxID = 127
    ReDim aFormats(0 To mMaxID) As String
    
'Form1.Text1 = mRTB.TextRTF
    
    'Get RTF for Header (through "\pard")
    strTextRTF = mRTB.TextRTF
    iPos = InStr(LCase(strTextRTF), "\pard\")
    If iPos = 0 Then Exit Sub 'ERROR???
    iStart = iPos + 5
    strHdr = Left(strTextRTF, iStart - 1)
    'and for Plain text (from "\pard" through first RTB_CHAR)
    iPos = InStr(iStart, strTextRTF, RTB_CHAR)
    If iPos = 0 Then Exit Sub
    strPlain = Trim(Mid(strTextRTF, iStart, iPos - iStart))
    If Left(LCase(strPlain), 6) <> "\plain" Then
        strPlain = "\plain" & strPlain
    End If
    
    'Get RTF for color and custom formats
    For iFormat = 128 - 3 To mMaxID
        'look for next "off" char (odd position)
        iStart = iPos + 1
        iPos = InStr(iStart, strTextRTF, RTB_CHAR)
        If iPos <= 0 Then Exit For
        iStart = iPos + 1
        'look for next "on" char (even position)
        iPos = InStr(iStart, strTextRTF, RTB_CHAR)
        If iPos <= 0 Then Exit For
        'rtf is what's between
        strRTF = Trim(Mid(strTextRTF, iStart, iPos - iStart))
        If iFormat >= 128 Then
            'custom format
            aFormats(iFormat) = strRTF
        Else
            'color: search for "\cf?"
            iStart = InStr(LCase(strRTF), "\cf")
            If iStart > 0 Then
                strRTF = Mid(strRTF, iStart)
                For iStart = 4 To 9
                    If Mid(strRTF, iStart, 1) < "0" Or Mid(strRTF, iStart, 1) > "9" Then
                        strRTF = Left(strRTF, iStart - 1)
                        Exit For
                    End If
                Next
                Select Case iFormat
                    Case 125:
                        aFormats(rtfBlue) = strRTF
                    Case 126:
                        aFormats(rtfRed) = strRTF
                    Case 127:
                        aFormats(rtfCustomColor) = strRTF
                End Select
            End If
        End If
    Next
    
    aFormats(rtfBold) = "\b"
    aFormats(rtfItalic) = "\i"
    aFormats(rtfUnderline) = "\ul"
    
    'get rtf for CustomPreset from its assigned CustomFormat
    iFormat = mCustomPreset
    If iFormat < 128 Then iFormat = iFormat + 128
    If iFormat >= 128 And iFormat <= mMaxID Then
        aFormats(rtfCustomPreset) = aFormats(iFormat)
    End If
    
    'Now do all the combinations of presets (0-127)
    For iFormat = mMaxID To 1 Step -1
        If iFormat >= 128 Then
            strRTF = aFormats(iFormat)
        Else
            'concatenate the presets according to the bitmask
            '(start with CustomPreset since it includes plain rtf)
            strRTF = ""
            For iPreset = 6 To 0 Step -1 '(64,32,16,8,4,2,1)
                If iFormat And (2 ^ iPreset) Then
                    strRTF = strRTF & aFormats(2 ^ iPreset)
                End If
            Next
        End If
        'make sure each format starts with "Plain"
        If Left(strRTF, Len(strPlain)) <> strPlain Then
            strRTF = strPlain & strRTF
        End If
        aFormats(iFormat) = strRTF
    Next
    
    'clear control's highlighted selection
    With mRTB
        .SelStart = 0
        .SelLength = 0
    End With
    
    If 0 Then
        strRTF = ""
        For iFormat = 0 To mMaxID
            strRTF = strRTF & Right("000" & Str(iFormat), 3) & ": " & aFormats(iFormat) & vbCrLf
        Next
        'Form1.Text1 = strRTF
    End If
    
End Sub

'Builds the RichTextFormat string and assigns it to the RichTextBox
Public Sub BuildRTF(Optional ByVal FreeTextWhenDone As Boolean = True)
    
    Dim strRTF$, strHdr$, strPlain$
    Dim i&, iPos&, iStartPos&, iLen&
    Dim nFormatID&, nNextFormatID&, nChar&
    Dim aFormats() As String
    Dim bufRTF As cMemBuffer
    
    'Can free extra memory since text appending appears to be done
    mText.FreeExtra
    mFormatting.FreeExtra
    
    'Lock RTBox's updating to avoid flicker
    LockWindowUpdate mRTB.hWnd
    
    'Load array with formatting codes for each FormatID
    LoadFormatsArray aFormats(), strHdr, strPlain
        
    If Len(strHdr) = 0 Then
        'ERROR with rtf, so just do plain text
        mRTB.Text = mText.Buffer
        mbFormatsRTFLoaded = False
    Else
        'Now build the RTF for the actual text
        Set bufRTF = New cMemBuffer
        With bufRTF
            'preallocate buffer to what will probably be big enough
            .Allocated = mText.Length * 2 '(but will grow if need to)
            'start with RTF "Header"
            .Buffer = strHdr
            If mText.Length > 0 Then
                'get FormatID for first character
                nFormatID = mFormatting.GetByte(0)
                iStartPos = 0 '(starting character for current formatting)
                For iPos = 0 To mText.Length - 1
                    'get FormatID for next character
                    If iPos + 1 < mText.Length Then
                        nNextFormatID = mFormatting.GetByte(iPos + 1)
                    Else
                        nNextFormatID = -2 '(end of text)
                    End If
                    'get current character in text
                    nChar = mText.GetByte(iPos)
                    'When hit last character for current formatting
                    'or when hit end of paragraph then do section ...
                    If nFormatID <> nNextFormatID Or nChar = 10 Then
                        'add the formatting codes for this section of text
                        If nFormatID > 0 And nFormatID <= mMaxID Then
                            .PutStr aFormats(nFormatID) & " "
                        Else 'do plain for invalid or no format
                            .PutStr strPlain & " "
                        End If
                        'add this section of text
                        iLen = iPos - iStartPos + 1
                        .PutFromMemory mText.MemPtr + iStartPos, iLen
                        'must add "\par " for vbCrLf
                        If nChar = 10 Then .PutStr "\par "
                        'reset the flags
                        nFormatID = nNextFormatID
                        iStartPos = iPos + 1 '(starting position for next formatting)
                    End If
                Next
                'RTF ends with "}"
                .PutStr "}"
            End If
            
            'Assign RTF text to the control
            mRTB.TextRTF = .Buffer
            mbFormatsRTFLoaded = False
        End With
    End If
    
    'FileFromString "c:\chk.rtf", mFormatsRTF
    
    'cleanup
    LockWindowUpdate 0 '(unlock RTBox's updating)
    Set bufRTF = Nothing
    If FreeTextWhenDone Then Clear
    ReDim aFormats(0) As String

End Sub


'Returns class data so contents can be stored
Public Property Get SerializeClassData() As String

    Dim buf As cMemBuffer, strRTF$
    Set buf = New cMemBuffer
    With buf
        If Not mbFormatsRTFLoaded And Not mRTB Is Nothing Then
            strRTF = mRTB.TextRTF 'save control's text
        End If
        UpdateFormatsRTF
        If Len(strRTF) > 0 Then
            mRTB.TextRTF = strRTF 'restore control's text
        End If
        If mText.Length > 0 And mText.Length = mFormatting.Length Then
            .Allocated = mText.Length + mFormatting.Length _
                    + Len(mFormatsRTF) + 100
            'write Version
            .PutLong CLASS_VERSION + 123000000, 0
            'write Text and Formatting
            .PutLong mText.Length
            .PutFromMemory mText.MemPtr, mText.Length
            .PutFromMemory mFormatting.MemPtr, mFormatting.Length
            'write FormatsRTF
            .PutLong Len(mFormatsRTF)
            .PutStr mFormatsRTF
            'write other stuff
            .PutInt mMaxID '(2 bytes)
            .PutInt mCustomPreset '(2 bytes)
            .PutLong mCustomColor
        End If
        SerializeClassData = .Buffer
    End With
    Set buf = Nothing
    
End Property

'Restores class data after previously being stored
Public Property Let SerializeClassData(ByVal strClassData As String)
    
    Dim strTemp$, n&
    Dim buf As cMemBuffer
    Set buf = New cMemBuffer
    With buf
        .Buffer = strClassData
        'make sure same class version
        n = .GetLong(0)
        If n = CLASS_VERSION + 123000000 Then
            'get text and formatting
            n = .GetLong
            mText.Buffer = .GetStr(n)
            mFormatting.Buffer = .GetStr(n)
            'get FormatsRTF
            n = .GetLong
            mFormatsRTF = .GetStr(n)
            mbFormatsRTFLoaded = False '(so new mFormatsRTF will get loaded)
            'get other stuff
            mMaxID = .GetInt '(2 bytes)
            mCustomPreset = .GetInt '(2 bytes)
            mCustomColor = .GetLong
        End If
    End With
    
End Property
