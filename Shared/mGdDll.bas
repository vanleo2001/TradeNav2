Attribute VB_Name = "mGdDll"
' NOTE: This module requires "G32_GD.DLL" (should be in System folder)
Option Explicit

Global Const USE_DEFAULT_NULL = -1.54321E-123 ' just something not likely to be used!
Global Const kNullData = -999999#

Enum eGdArray_Type
    eGDARRAY_NoArray = 0    'no array yet
    eGDARRAY_Doubles = 68   'D: 8 bytes
    eGDARRAY_Floats = 70    'F: 4 bytes
    eGDARRAY_Longs = 76     'L: 4 bytes (+/- 2 billion)
    eGDARRAY_Shorts = 83    'S: 2 bytes -32767 to 32767 (null=-32768)
    eGDARRAY_TinyInts = 84  'T: 1 byte: -127 to 127 (null=-128)
    eGDARRAY_Strings = 33   '!: string array
    eGDARRAY_gdString = 36  '$: a gdString
    eGDARRAY_Objects = 42   '*: array of objects
End Enum

Enum eGdSort_Flags
    eGdSort_Default = 0
    ' below 0x10000 reserved for "at_string_pos" (for string array sorting)
    eGdSort_Descending = &H10000
    eGdSort_DeleteNullValues = &H20000 ' deletes all null values
    eGdSort_IgnoreCase = &H40000 ' string arrays only
    eGdSort_Stable = &H80000
    eGdSort_DeleteDuplicates = &H100000 ' leaves only unique values
    eGdSort_MatchUsingSearchStringLength = &H1000000 ' only for searching a string array
End Enum


' used by "gdCalcStatistic" and "gdCalcMovingStatistic"
Enum eGdStatisticType
    eGdStat_Sum = 1             ' Sum of each item
    eGdStat_Average = 2         ' Average = Sum / Count
    eGdStat_SumOfSquares = 10   ' Sum of the square of each item
    eGdStat_StdDev = 11         ' Standard Deviation
    eGdStat_Variance = 12       ' Variance = StdDev squared
    eGdStat_CoefOfVariation = 13  ' Coefficient of Variation = 100 * StdDev / Average
    eGdStat_StdDevSample = 21         '(for sample, not total population)
    eGdStat_VarianceSample = 22       '(for sample, not total population)
    eGdStat_CoefOfVariationSample = 23  '(for sample, not total population)
End Enum

' used by "gdCalcStatistics"
Public Type gdArrayStatistics
    ' 256 byte structure
    Count As Long
    iReserved As Long ' (future use)
    MinValue As Double
    MaxValue As Double
    Range As Double
    Sum As Double
    SumOfSquares As Double
    Average As Double
    StdDev As Double
    Variance As Double
    CoefOfVariation As Double
    AvgDev As Double
    SumDevSq As Double
    Skewness As Double
    Kurtosis As Double
    dReserved(18) As Double ' (future use)
End Type

' used by gdFile matching functions
Public Type gdFileMatchingTotals
    dMatchedBytes As Double
    nMatchedFiles As Long
    nMatchedFolders As Long
    nFoldersSearched As Long
    nReserved As Long
End Type

' following is obsolete -- now using gdTickCountVB in G32_ZIP.DLL
'Declare Function gdTickCount Lib "G32_GD.dll" () As Double

' for profiling: can use ID's from 0-999
Private Declare Sub DLL_gdResetProfiles Lib "G32_GD.dll" Alias "gdResetProfiles" (ByVal nFromID&, ByVal nToID&)
Declare Sub gdStartProfile Lib "G32_GD.dll" (ByVal nID&)
Declare Sub gdStopProfile Lib "G32_GD.dll" (ByVal nID&)
Private Declare Function gdGetProfilesString Lib "G32_GD.dll" (ByVal nFromID&, ByVal nToID&, ByVal strDelimiter$) As Long

'Declare Function pass_variant& Lib "G32_GD.dll" (var1 As Variant, var2 As Variant)

' returns formatted string to display a price
Declare Function gdFormatPriceString& Lib "G32_GD.dll" (ByVal Buffer$, ByVal iMaxCount&, ByVal dPrice#, ByVal dTickMove#, ByVal dMinMoveInTicks#, ByVal sDecimalDigits%)

' for debugging
Declare Function gdTrackingItem& Lib "G32_GD.dll" (ByVal TrackingItem%)
Declare Function gdCheckMemoryLeaks& Lib "G32_GD.dll" (ByVal chk_mode&)
Declare Function gdTestFunction& Lib "G32_GD.dll" (ByVal hArray&)

' To create and initialize an array (returns a "handle" to the newly
' created array/string that the caller is responsible to destroy)
Declare Function gdCreateString& Lib "G32_GD.dll" (ByVal Size&)
Private Declare Function DLL_gdCreateArray& Lib "G32_GD.dll" Alias "gdCreateArray" (ByVal array_type As Byte, ByVal Size&, ByVal null_value#)
' To destroy an array when done with it.
Declare Sub gdDestroyArray Lib "G32_GD.dll" (hArray&)
Declare Sub gdDestroyString Lib "G32_GD.dll" (hString&)

' Returns array type.
Declare Function gdGetType Lib "G32_GD.dll" (ByVal hArray&) As Byte

' Returns # of items currently allocated to hold.
Declare Function gdGetAllocated& Lib "G32_GD.dll" (ByVal hArray&)

' Returns TRUE if there are no items in the array.
Declare Function gdIsEmpty& Lib "G32_GD.dll" (ByVal hArray&)

' Reinitializes array (makes it empty).
Declare Sub gdClear Lib "G32_GD.dll" (ByVal hArray&, ByVal free_extra_memory&)

' Allocate enough memory to hold specified # of items.
Declare Function gdReserve& Lib "G32_GD.dll" (ByVal hArray&, ByVal num_allocate&, ByVal preserve_data&)

' Set size of the array (# of items).
Declare Function gdSetSize& Lib "G32_GD.dll" (ByVal hArray&, ByVal Size&, ByVal preserve_data&)

' Frees up any extra memory at end of array.
Declare Sub gdFreeExtra Lib "G32_GD.dll" (ByVal hArray&)

' Resets one or more items to NullValue.
' (this function will NOT cause the array to grow or change size)
Declare Function gdNullTheItems& Lib "G32_GD.dll" (ByVal hArray&, ByVal from_item&, ByVal to_item&)

' for debugging: displays info about the array
Declare Sub gdDump Lib "G32_GD.dll" (ByVal hArray&, ByVal msg$, ByVal num_items&)

' Returns current size of array (# of items).
Declare Function gdGetSize& Lib "G32_GD.dll" (ByVal hArray&)

' Returns TRUE if this is a "constant-value" array.
Declare Function gdIsConstantValue& Lib "G32_GD.dll" (ByVal hArray&)

' Returns address to the actual data itself (advanced feature!).
Declare Function gdGetDataPtr& Lib "G32_GD.dll" (ByVal hArray&)

' Get offset "shift".
Declare Function gdGetShifted& Lib "G32_GD.dll" (ByVal hArray&)
' Set the offset "shift".
Declare Function gdSetShifted& Lib "G32_GD.dll" (ByVal hArray&, ByVal new_shift&, ByVal relative_to_current&)

' Get the shared-data mode property.
Declare Function gdGetShared& Lib "G32_GD.dll" (ByVal hArray&)
' Set the shared-data mode property.
Declare Function gdSetShared& Lib "G32_GD.dll" (ByVal hArray&, ByVal shared_data_mode&)

' Deletes one or more items from array (making the array smaller).
Declare Sub gdDeleteItems Lib "G32_GD.dll" (ByVal hArray&, ByVal from_item&, ByVal delete_count&)

' Moves one or more items up or down in the array.
' - block_start: starting offset of block to move
' - block_size: number of items to move
' - shift_amount: number of positions to move block up/down
'      (if > 0, moves toward end; if < 0, moves toward beginning)
' - returns true if performed a move
Declare Function gdMoveItems& Lib "G32_GD.dll" (ByVal hArray&, ByVal block_start&, ByVal block_size&, ByVal shift_amount&)

' Appends items/records from another array/table (data types must match).
' - returns true if one or more items were appended
Declare Function gdAppendFrom& Lib "G32_GD.dll" (ByVal hArrayTo&, ByVal hArrayFrom&, ByVal from_item&, ByVal to_item&)

' Parse numeric or string array items from a string.
Declare Function gdSplitFields& Lib "G32_GD.dll" (ByVal hArray&, ByVal string_to_parse$, ByVal delimiters$, ByVal num_fields&)

' Join string array items into a single string.
'(gdJoinFields returns a handle to a newly created gdString that caller is responsible for destroying)
Declare Function gdJoinFields& Lib "G32_GD.dll" (ByVal hArray&, ByVal delimiters$)

' Use "quicksort" algorithm to sort the array.
' (commonly accepted as most efficient sorting algorithm)
Declare Function gdSort& Lib "G32_GD.dll" (ByVal hArray&, ByVal eSortFlags As eGdSort_Flags, ByVal Bottom&, ByVal Top&)
Declare Function gdSortAsIndex& Lib "G32_GD.dll" (ByVal hArray&, ByVal hSortByArray&, ByVal initialize_index&, ByVal eSortFlags As eGdSort_Flags, ByVal Bottom&, ByVal Top&)

' Performs a binary search on a SORTED array.
' - if FOUND: returns TRUE, iPos is position of match
' - if NOT FOUND: returns FALSE, iPos is position to insert at
Declare Function gdBinarySearch& Lib "G32_GD.dll" (ByVal hArray&, ByVal search_for#, iPos&, ByVal eSortFlags As eGdSort_Flags, ByVal Bottom&, ByVal Top&)
Declare Function gdBinarySearchAsIndex& Lib "G32_GD.dll" (ByVal hArray&, ByVal hSortedByArray&, ByVal search_for#, iPos&, ByVal eSortFlags As eGdSort_Flags, ByVal Bottom&, ByVal Top&)

' Builds an indexed list (an array used as an index to another array)
' sorted by another array and optionally using another array as a filter.
''Declare Function gdIndexList& Lib "G32_GD.dll" (ByVal hListArray&, ByVal hFilterArray&, ByVal hSortByArray&, ByVal eSortFlags As eGdSort_Flags)

' Shuffles (randomizes) elements of array.
Declare Sub gdShuffle Lib "G32_GD.dll" (ByVal hArray&, ByVal Bottom&, ByVal Top&)

Declare Function gdRandomNumber& Lib "G32_GD.dll" (ByVal min_num&, ByVal max_num&)

' Returns the array's NullValue.
Declare Function gdNullValue# Lib "G32_GD.dll" (ByVal hArray&)

' Makes the array a "constant-value".
Declare Function gdMakeConstantValue& Lib "G32_GD.dll" (ByVal hArray&, ByVal Value#, ByVal Size&)

' Get the item at the specified offset.
Declare Function gdGetNum# Lib "G32_GD.dll" (ByVal hArray&, ByVal offset&)

' Set an item to the specified value.
' (will grow, set size, and alloc/destroy objects)
Declare Function gdSetNum& Lib "G32_GD.dll" (ByVal hArray&, ByVal offset&, ByVal Value#)

' Adds an item to the array (returns offset added).
Declare Function gdAddNum& Lib "G32_GD.dll" (ByVal hArray&, ByVal Value#)

' Inserts an item into the array at the specified location.
Declare Function gdInsertNum& Lib "G32_GD.dll" (ByVal hArray&, ByVal Value#, ByVal insert_at&)

Declare Function gdGetChar Lib "G32_GD.dll" (ByVal hArray&, ByVal offset&) As Byte

Declare Function gdSetChar& Lib "G32_GD.dll" (ByVal hArray&, ByVal offset&, ByVal Value As Byte)

Private Declare Function DLL_gdGetStr& Lib "G32_GD.dll" Alias "gdGetStr" (ByVal hArray&, ByVal offset&, ByVal strGet$, ByVal max_length&)

Declare Function gdSetStr& Lib "G32_GD.dll" (ByVal hArray&, ByVal offset&, ByVal strSet$)

Declare Function gdAddStr& Lib "G32_GD.dll" (ByVal hArray&, ByVal strAdd$)

Declare Function gdInsertStr& Lib "G32_GD.dll" (ByVal hArray&, ByVal strInsert$, ByVal insert_at&)

' Get the item at the specified offset.
Declare Function gdGetVariant& Lib "G32_GD.dll" (ByVal hArray&, ByVal offset&, obj As Variant)

' Set an item to the specified value.
' (will grow, set size, and alloc/destroy objects)
Declare Function gdSetVariant& Lib "G32_GD.dll" (ByVal hArray&, ByVal offset&, obj As Variant)

' Adds an item to the array (returns offset added).
Declare Function gdAddVariant& Lib "G32_GD.dll" (ByVal hArray&, obj As Variant)

' Inserts an item into the array at the specified location.
Declare Function gdInsertVariant& Lib "G32_GD.dll" (ByVal hArray&, obj As Variant, ByVal insert_at&)


' Assignment operator (reattach to an existing array).
Declare Function gdCopy& Lib "G32_GD.dll" (ByVal hToArray&, ByVal hFromArray&)

Declare Function gdIndexOfMaxValue& Lib "G32_GD.dll" (ByVal hArray&, ByVal nFromItem&, ByVal nToItem&)
Declare Function gdIndexOfMinValue& Lib "G32_GD.dll" (ByVal hArray&, ByVal nFromItem&, ByVal nToItem&)
Declare Function gdMaxValue# Lib "G32_GD.dll" (ByVal hArray&, ByVal nFromItem&, ByVal nToItem&)
Declare Function gdMinValue# Lib "G32_GD.dll" (ByVal hArray&, ByVal nFromItem&, ByVal nToItem&)
Declare Function gdRange# Lib "G32_GD.dll" (ByVal hArray&, ByVal nFromItem&, ByVal nToItem&)
Declare Function gdCount& Lib "G32_GD.dll" (ByVal hArray&, ByVal nFromItem&, ByVal nToItem&)
Declare Function gdCountOf& Lib "G32_GD.dll" (ByVal hArray&, ByVal dMatchValue#, ByVal nFromItem&, ByVal nToItem&)

' Calculates the "best fit" line for all or a portion of the array
' using the common "least squares" linear regression method ...
' - returns true if able to calculate (at least 2 valid points)
' - also returns the "slope" and "intercept" of best fit line,
'      such that can then easily calculate any point on the line:
'          BestFitPoint = at_offset * slope + intercept
' - can also return "std_dev" for confidence band (std dev of "errors")
Declare Function gdBestFitLine& Lib "G32_GD.dll" (ByVal hArray&, dSlope As Double, dIntercept As Double, dStdDev As Double, ByVal nFromItem&, ByVal nToItem&)

' Calculates the "best fit" curve for all or a portion of the array
' using the second-order "least squares" regression method ...
' - returns true if able to calculate (at least 2 valid points)
' - also returns the coefficients of best fit curve,
'      such that can then easily calculate any point on the curve:
'      BestFitPoint = pow(at_offset, 2) * sq_coeff + at_offset * slope + intercept
' - can also return "std_dev" for confidence band (std dev of "errors")
Declare Function gdBestFitCurve& Lib "G32_GD.dll" (ByVal hArray&, dSqCoeff As Double, dSlope As Double, dIntercept As Double, dStdDev As Double, ByVal nFromItem&, ByVal nToItem&)

' Calculates various statistics for an array (count, min, max, range, average,
' standard deviation, coeffiecient of variation, skewness, kurtosis, etc.)
' - set bSampleOfPopulation = True if array is a sampling (does not contain all the data)
' - for whole array: pass nFromItem = 0, nToItem = -1
Declare Function gdCalcStatistics& Lib "G32_GD.dll" (ByVal hArray&, Stats As gdArrayStatistics, ByVal bSampleOfPopulation&, ByVal nFromItem&, ByVal nToItem&)

' Calculates a statistic for an array
' - eType: sum, average, sum of squares, standard deviation, variance, etc.
' - for whole array: pass nFromItem = 0, nToItem = -1
Declare Function gdCalcStatistic Lib "G32_GD.dll" (ByVal hInputArray&, ByVal eStatType As eGdStatisticType, ByVal nFromItem&, ByVal nToItem&) As Double

' Calculates various moving statistics for an input array (for a rolling "window")
' - eType: sum, average, sum of squares, standard deviation, variance, etc.
' - dPeriods: size of the rolling window, or 0 = from beginning
Declare Function gdCalcMovingStatistic& Lib "G32_GD.dll" (ByVal hResultArray&, ByVal hInputArray&, ByVal eStatType As eGdStatisticType, ByVal nPeriods#)

' Serialize array (read/write to binary file).
Declare Function gdSerializeArray& Lib "G32_GD.dll" (ByVal hArray&, ByVal hFile&, ByVal bPut&)
' Serialize data arrays of a gdBars (read/write to binary file).
Declare Function gdSerializeBarsArrays& Lib "G32_GD.dll" (ByVal hBars&, ByVal hFile&, ByVal bPut&)
' File operations
Declare Function gdFileOpen& Lib "G32_GD.dll" (ByVal strFileName$, ByVal strMode$)
Declare Sub gdFileClose Lib "G32_GD.dll" (hFile&)
Declare Function gdFileBinaryIO& Lib "G32_GD.dll" (ByVal hFile&, vPtr As Any, ByVal nBytes&, ByVal bPut&)
Declare Function gdFileStringIO& Lib "G32_GD.dll" (ByVal hFile&, ByVal strString$, ByVal nBytes&, ByVal bPut&)
Declare Sub gdFileFlush Lib "G32_GD.dll" (ByVal hFile&)


' Array operation -- supports following operations:
' Math:  +  -  *  /  MOD  POWER
' Comparison:  =  <>  <  >  <=  >=
' Logical:  AND  OR  XOR(Exclusive OR)  NOR(Neither)  NAND(Not Both)
' Functions:  MIN  MAX  ROUND(rounds Op1 to Op2 decimals)
' Shift (Op1 shifted by Op2):  SHIFT(forward and backward shifts)  BACK(backward shifts only)
' Unary functions (Op2 n/a):  IS  NOT  ABS  SIGN  INT  FRACTPART  FLOOR  CEILING
'                               SINE  COSINE  NATLOG  LOG10
Declare Function gdArrayOperate& Lib "G32_GD.dll" (ByVal hResultArray&, ByVal hArray1&, ByVal strOperation$, ByVal hArray2&)
Declare Function gdArrayItemOperate Lib "G32_GD.dll" (ByVal operand1#, ByVal strOperation$, ByVal operand2#) As Double


' To efficiently compare gdStrings (alphabetic comparison):
' - if strCompareOperator is passed (e.g. "=", ">", "<=", "<>", etc.),
'      then return is true/false
' - else if strCompareOperator is empty, return is as follows:
'      negative if str1 < str2, 0 if equal, positive if str1 > str2
' - if bIgnoreCase is true, comparison is not case-sensitive
' - if MaxCount > 0, compares only up to specified number of characters (0 = compare all)
Declare Function gdStringCompare& Lib "G32_GD.dll" (ByVal hStr1&, ByVal strCompareOperator$, ByVal hStr2&, ByVal bIgnoreCase&, ByVal MaxCount&)


'==========================================
' Table
'(gdCreateTable returns a handle to a newly created gdTable that caller is responsible for destroying)
Declare Function gdCreateTable& Lib "G32_GD.dll" (ByVal nNumRecords&)
Declare Sub gdDestroyTable Lib "G32_GD.dll" (hTable&)
Declare Function gdClearField& Lib "G32_GD.dll" (ByVal hTable&, ByVal nField&)
Declare Function gdNumRecords& Lib "G32_GD.dll" (ByVal hTable&)
Declare Function gdNumFields& Lib "G32_GD.dll" (ByVal hTable&)
Declare Function gdSetNumRecords& Lib "G32_GD.dll" (ByVal hTable&, ByVal nNumRecords&)
Declare Function gdFieldArrayHandle& Lib "G32_GD.dll" (ByVal hTable&, ByVal nField&)
Declare Function gdSetFieldName& Lib "G32_GD.dll" (ByVal hTable&, ByVal nField&, ByVal strFieldName$)
Declare Function gdFieldType Lib "G32_GD.dll" (ByVal hTable&, ByVal nField&) As Byte
Declare Function gdFieldNum& Lib "G32_GD.dll" (ByVal hTable&, ByVal strFieldName$)
Declare Function gdCreateField& Lib "G32_GD.dll" (ByVal hTable&, ByVal eArrayType As Byte, ByVal nField&, ByVal strFieldName$)
Declare Function gdAttachField& Lib "G32_GD.dll" (ByVal hTable&, ByVal hArrayHandle&, ByVal nField&, ByVal strFieldName$)
Declare Function gdSetTableNum& Lib "G32_GD.dll" (ByVal hTable&, ByVal nField&, ByVal nRecord&, ByVal dNumber#)
Declare Function gdSetTableStr& Lib "G32_GD.dll" (ByVal hTable&, ByVal nField&, ByVal nRecord&, ByVal strText$)
Declare Function gdGetTableNum# Lib "G32_GD.dll" (ByVal hTable&, ByVal nField&, ByVal nRecord&)
'(gdGetTableStr returns a handle to a newly created gdString that caller is responsible for destroying)
' If you want a VB String back instead, use the gdGetTableString wrapper
Private Declare Function gdGetTableStr& Lib "G32_GD.dll" (ByVal hTable&, ByVal nField&, ByVal nRecord&)
'(gdFieldName returns a handle to a newly created gdString that caller is responsible for destroying)
Declare Function gdFieldName& Lib "G32_GD.dll" (ByVal hTable&, ByVal nField&)
''Declare Function gdCreateTableIndex& Lib "G32_GD.dll" (ByVal hTable&, ByVal nSortByField&, ByVal eSortFlags As eGdSort_Flags, ByVal nFilterByField&)
''Declare Function gdBuildTableIndex& Lib "G32_GD.dll" (ByVal hTable&, ByVal hIndex&, ByVal nSortByField&, ByVal eSortFlags As eGdSort_Flags, ByVal InitializeIndex&, ByVal nFilterByField&)

'(gdCreateTableIndex returns a handle to a newly created gdArray of Longs that caller is responsible for destroying)
Declare Function gdCreateTableIndex& Lib "G32_GD.dll" (ByVal hTable&, ByVal nFilterByField&)
Declare Function gdSortTableIndex& Lib "G32_GD.dll" (ByVal hTable&, ByVal hIndex&, ByVal nSortByField&, ByVal eSortFlags As eGdSort_Flags)
Declare Function gdSerializeTable& Lib "G32_GD.dll" (ByVal hTable&, ByVal hFile&, ByVal bPut&)
Declare Function gdTableIsNullAt& Lib "G32_GD.dll" (ByVal hTable&, ByVal nField&, ByVal nRecord&)

' gets a record of the table as a delimited string
Declare Function gdGetTableRecord& Lib "G32_GD.dll" (ByVal hTable&, ByVal nRecord&, ByVal strFieldDelim$)
' sets a record of the table from a delimited string
Declare Function gdSetTableRecord& Lib "G32_GD.dll" (ByVal hTable&, ByVal strRecord$, ByVal nRecord&, ByVal strFieldDelim$)
' adds/inserts a record of the table from a delimited string
Declare Function gdAddTableRecord& Lib "G32_GD.dll" (ByVal hTable&, ByVal strRecord$, ByVal nInsertAtRecord&, ByVal strFieldDelim$)
' deletes one or more records of the table
Declare Function gdDeleteTableRecords& Lib "G32_GD.dll" (ByVal hTable&, ByVal nFromRecord&, ByVal nDeleteCount&)
' gets data for table as a double-delimited string
Declare Function gdTableToString& Lib "G32_GD.dll" (ByVal hTable&, ByVal strRecordDelim$, ByVal strFieldDelim$)
' sets data for table from a double-delimited string
Declare Function gdTableFromString& Lib "G32_GD.dll" (ByVal hTable&, ByVal strTableData$, ByVal strRecordDelim$, ByVal strFieldDelim$)


'=========================================
' Bars

Enum eBarsArray
    eBARS_DateTime = &H1&
    eBARS_Open = &H4&
    eBARS_High = &H8&
    eBARS_Low = &H10&
    eBARS_Close = &H20&
    eBARS_Vol = &H40&
    eBARS_OI = &H80&
    eBARS_ContVol = &H100&
    eBARS_ContOI = &H200&

    eBARS_UpTicks = &H400&
    eBARS_DownTicks = &H800&

    eBARS_Bid = &H1000&
    eBARS_BidSize = &H2000&
    eBARS_Ask = &H4000&
    eBARS_AskSize = &H8000&
    eBARS_HighBid = &H10000
    eBARS_LowBid = &H20000
    eBARS_HighAsk = &H40000
    eBARS_LowAsk = &H80000
    eBARS_Flags = &H100000
    eBARS_BidVol = &H200000
    eBARS_AskVol = &H400000

    eBARS_Prices = (eBARS_DateTime Or eBARS_Open Or eBARS_High Or eBARS_Low Or eBARS_Close)
    eBARS_VolOI = (eBARS_Vol Or eBARS_OI Or eBARS_ContVol Or eBARS_ContOI)
    eBARS_BidAsk = (eBARS_Bid Or eBARS_BidSize Or eBARS_Ask Or eBARS_AskSize Or eBARS_HighBid Or eBARS_LowBid Or eBARS_HighAsk Or eBARS_LowAsk)

    eBARS_Eod = (eBARS_Prices Or eBARS_VolOI)
    eBARS_EodBidAsk = (eBARS_Eod Or eBARS_BidAsk)
    eBARS_Intraday = (eBARS_Prices Or eBARS_Vol Or eBARS_UpTicks Or eBARS_DownTicks Or eBARS_BidVol Or eBARS_AskVol)
    eBARS_TickByTick = (eBARS_DateTime Or eBARS_Close Or eBARS_Vol Or eBARS_Flags)
    eBARS_Minutized = (eBARS_DateTime Or eBARS_Close Or eBARS_Vol Or eBARS_UpTicks Or eBARS_DownTicks Or eBARS_BidVol Or eBARS_AskVol)
    
    ' TLB 8/22/2012: a special configuration only used for market profile data (i.e. not charting, etc)
    eBARS_Profiled = (eBARS_DateTime Or eBARS_Close Or eBARS_Vol Or eBARS_Flags Or eBARS_BidVol Or eBARS_AskVol)
End Enum

Enum eBarsPropID
    ' numeric:
    eBARS_SymbolID = 1
    eBARS_IfMarketDefined = 2
    eBARS_TickMove = 3
    eBARS_TickValue = 4
    eBARS_MinMoveInTicks = 5
    eBARS_Margin = 6
    eBARS_ContractSize = 7
    eBARS_hDataMgr = 8
    eBARS_ConvFactor = 9
    eBARS_Periodicity = 10      ' Periodicity = PeriodType + PeriodsPerBar
    eBARS_PeriodType = 11       ' see eBarsPeriodType
    eBARS_PeriodsPerBar = 12    ' e.g. 1 for "daily", 3 for "3 days", 20 for "20 minutes"
    eBARS_SecurityType = 13     ' read-only
    eBARS_Contract = 14         ' read-only
    eBARS_IsOption = 15         ' read-only
    eBARS_LastTickTime = 16     ' minutes from midnight for last tick (if snapshot data)
    eBARS_ArrayMask = 17        ' 32-bit mask of which arrays are being used
    eBARS_CrossoverTime = 18    ' minutes from midnight
    eBARS_StartTime = 19        ' minutes from midnight (can be custom-set)
    eBARS_EndTime = 20          ' minutes from midnight (can be custom-set)
    eBARS_CsiNumber = 21        ' commodity # used by CSI
    eBARS_SuspendTime = 22      ' minutes from midnight (end of first session)
    eBARS_ResumeTime = 23       ' minutes from midnight (start of second session)
    eBARS_OddBarsUp = 24        ' true if odd bars are up (only for Point&Figure, Kagi, Renko)
    eBARS_ExpiresPriorMonth = 25 ' true if symbol expires in month prior to delivery
    eBARS_Session = 26          'D'=day, 'C'=combined, 'S'=synthetic day (for electronic)
    eBARS_LastTickDown = 27     ' if last tick was down (below previous)
    eBARS_LastTickVol = 28      ' vol of last trade (if known)
    eBARS_IsExternalSymbol = 29 ' true if ExternalSecType flag is set (e.g. data from other vendors)
    eBARS_PriceHasSettled = 30  ' true if last close represents a "settled" price
    eBARS_TimeInSeconds = 31    ' true if time in DateTime resolved to seconds (instead of minutes)
    eBARS_DefaultStartTime = 32 ' minutes from midnight (from BaseSyms)
    eBARS_DefaultEndTime = 33   ' minutes from midnight (from BaseSyms)
    eBARS_FractZen = 34         ' true if these are "merged FractZen" bars

    ' strings:
    eBARS_Symbol = 101
    eBARS_Desc = 102
    eBARS_MarketSymbol = 103
    eBARS_CustomString = 104    ' can be used for temporary storage purposes
    eBARS_PeriodicityStr = 105  ' e.g. "Daily", "2 days", "30 minute", "100 ticks"
    eBARS_BaseSymbol = 106      ' read-only
    eBARS_OptionSymbol = 107    ' read-only
    eBARS_PriceDisplayFormat = 108
    eBARS_Exchange = 109
    eBARS_ExchangeTimeZoneInf = 110
End Enum

' Note:  Periodicity = PeriodType + PeriodsPerBar
Enum eBarsPeriodType
    ' Intraday types
    ePRD_EachTick = &H1000000   '  16777216
    ePRD_Ticks = &H2000000      '  33554432
    ePRD_Minutes = &H3000000    '  50331648
    ePRD_SMP = &H4000000        '  67108864 - Steidlmayer Market Profile bars
    ePRD_IntBreakout = &H5000000 ' 83886080
    ePRD_IntRenko = &H6000000   ' 100663296
    ePRD_IntKagi = &H7000000    ' 117440512
    ePRD_IntPF = &H8000000      ' 134217728
    ePRD_IntVol = &H9000000     ' 150994944
    ' End-of-day types
    ePRD_Days = &H11000000      ' 285212672
    ePRD_Weeks = &H12000000     ' 301989888
    ePRD_Months = &H13000000    ' 318767104
    ePRD_Quarters = &H14000000  ' 335544320
    ePRD_Years = &H15000000     ' 352321536
    ePRD_EodRenko = &H16000000  ' 369098752
    ePRD_EodKagi = &H17000000   ' 385875968
    ePRD_EodPF = &H18000000     ' 402653184
    ePRD_EodVol = &H19000000    ' 419430400
    ePRD_EodBreakout = &H20000000 '536870912
End Enum

Enum eTickFlags
    eTICK_Unknown = 0
    eTICK_AtBid = 1     ' i.e. SELL Volume
    eTICK_AtAsk = 2     ' i.e. BUY Volume
    eTICK_InRange = 3
    eTICK_Settle = 4
    'eTICK_Mask = 7
End Enum

'(gdCreateBars returns a handle to a newly created gdBars that caller is responsible for destroying)
Declare Function gdCreateBars& Lib "G32_GD.dll" (ByVal nSize&, ByVal WhichArrays As eBarsArray)
Declare Sub gdDestroyBars Lib "G32_GD.dll" (hBars&)

Declare Function gdBarsArray& Lib "G32_GD.dll" (ByVal hBars&, ByVal WhichArray As eBarsArray)
Declare Function gdBarsData# Lib "G32_GD.dll" (ByVal hBars&, ByVal WhichArray As eBarsArray, ByVal offset&)

' To read bars from a data file.
' - strFormat: "CSI", "MS7", "GEN" (Genesis EOD), "GT" (GenTick)
' - strPath: path where data file is located
' - strSymbol: pass either the symbol (leave blank to use the current bars.Symbol)
'      or pass the name of the data file to load (e.g. "F104.DTA")
' - strPeriod: "D" = daily (default), "W" = weekly, "10" = 10 minute bars, etc.
' - AppendMode: >0 = append at bar record, 0 = by dates, -1 = by file record #
' - returns # of bars loaded, or negative number for error
Declare Function gdBarsFromFile& Lib "G32_GD.dll" (ByVal hBars&, _
        ByVal strFormat$, ByVal strPath$, ByVal strSymbol$, _
        ByVal strPeriod$, ByVal FromDate&, ByVal ToDate&, _
        ByVal AppendMode&, ByVal hErrMsg&, ByVal hAlignDates&)

' To write bars to a data file.
' - strFormat: "CSI", "MS7", "GEN" (Genesis EOD), "GT" (GenTick)
' - strPath: path where data file is located
' - strSymbol: pass either the symbol (leave blank to use the current bars.Symbol)
'      or pass the name of the data file to load (e.g. "F104.DTA")
' - strDesc: if strSymbol and strDesc are blank, will use current bars.Desc
' - ConvFactor: if < -5, will use current bars.ConvFactor
' - AddMode: 0=overwrite file, 1=add to existing file, 2=add + auto-detect conv_factor
' - returns # of bars written, or negative number for error
Declare Function gdBarsToFile& Lib "G32_GD.dll" (ByVal hBars&, _
        ByVal strFormat$, ByVal strPath$, ByVal strSymbol$, ByVal strDesc$, _
        ByVal ConvFactor&, ByVal AddMode&, ByVal hErrMsg&)

' To build bars (e.g. Daily, Weekly, Monthly, Quarterly, Yearly)
' - strPeriod: as a string (optional number then letter/word)
'   - "Daily", "Day", "D"  =  Daily  (same for Weekly, Monthly, etc.)
'   - "2 days", "2d"  =  2 days per bar (same for weeks per bar, etc.)
'   - "30 minute", "30m", "30"  =  30 minute bars
'   - "100 ticks", "100t"  =  100 ticks per bar
'   - "500 vol", "500v"  =  1000 intraday volume per bar
'   - "500k vol", "500kv"  =  500000 end-of-day volume per bar
' - hFromBars: if NULL, will use current hToBars as source
' - bAppendFromBars: to append FromBars to source bars
' - hSplitDates: if used, will separate weekly/monthly bars at splits
' - returns success
Declare Function gdBuildBars2& Lib "G32_GD.dll" (ByVal hToBars&, ByVal strPeriod As String, ByVal hFromBars&, ByVal bAppendFromBars&, ByVal hSplitDates&)

' To add empty forecast bars (with dates filled in) to end of bars
' - nNumForecastBars: number of empty bars to add
' - hHolidays: optional handle to array of holidays (gdArrayD)
Declare Function gdAddForecastBars& Lib "G32_GD.dll" (ByVal hBars&, ByVal nNumForecastBars&, ByVal hHolidays&)

' To align bars to an array of dates or another set of bars.
' - returns # of bars now sized to
Declare Function gdAlignBars& Lib "G32_GD.dll" (ByVal hBars&, ByVal hAlignTo&, ByVal nFlags&)

Declare Sub gdRemoveOvernightPriceGap Lib "G32_GD.dll" (ByVal hBars&)

Declare Function gdGetBarsNumProp# Lib "G32_GD.dll" (ByVal hBars&, ByVal ePropID As eBarsPropID)
Declare Function gdSetBarsNumProp& Lib "G32_GD.dll" (ByVal hBars&, ByVal ePropID As eBarsPropID, ByVal dNumber#)
'(gdGetBarsStrProp returns a handle to a newly created gdString that caller is responsible for destroying)
Declare Function gdGetBarsStrProp& Lib "G32_GD.dll" (ByVal hBars&, ByVal ePropID As eBarsPropID)
Declare Function gdSetBarsStrProp& Lib "G32_GD.dll" (ByVal hBars&, ByVal ePropID As eBarsPropID, ByVal strString$)

Declare Function gdGetBarsMinMove# Lib "G32_GD.dll" (ByVal hBars&, ByVal dForDate As Double)
'Declare Function gdBarsPriceDisplay& Lib "G32_GD.dll" (ByVal hBars&, ByVal dPrice As Double, ByVal nDisplayFlags As Long)
Declare Function gdBarsPriceDisplay2& Lib "G32_GD.dll" (ByVal hBars&, ByVal dPrice As Double, ByVal nDisplayFlags As Long, ByVal nForSessionDate As Long)
Declare Function gdBarsPriceFromString# Lib "G32_GD.dll" (ByVal hBars&, ByVal strPrice$)
Declare Sub gdBarsFixPrices Lib "G32_GD.dll" (ByVal hBars&)
Declare Sub gdBarsFixPrices2 Lib "G32_GD.dll" (ByVal hBars&, ByVal nStartBar&, ByVal dPriceAdjust#)
Declare Sub gdBarsFixTicksHelper Lib "G32_GD.dll" (ByVal hBars&, ByVal nStart&, ByVal dPastGoodPrice#)

' Rounding now done in C++ (more consistent and more efficient)
Declare Function RoundToMinMove Lib "G32_GD.dll" Alias "gdRoundPriceToMinMove" (ByVal dPrice#, ByVal dMinMove#) As Double

Declare Sub gdDeleteFirstBars Lib "G32_GD.dll" (ByVal hBars&, ByVal nCount&)
Declare Sub gdDeleteSomeBars Lib "G32_GD.dll" (ByVal hBars&, ByVal nFromBar&, ByVal nCount&)
Declare Sub gdAppendBars Lib "G32_GD.dll" (ByVal hBars&, ByVal nFromBars&, ByVal bPrepend&)

' Performs a binary search on the DateTime array.
' - returns bar number where the date matches or would be inserted
'      (i.e. bar number with the lowest date >= dtFindDateTime)
' - returns -1 if bExactMatch = true and no exact match is found
Declare Function gdBarsFindDateTime& Lib "G32_GD.dll" (ByVal hBars&, ByVal dDateTimeToFind#, ByVal bExactMatch&)

' converts DateTime of bar to the specified time zone (see "ConvertTimeZone" for format spec)
Declare Function gdBarsDateTimeConvert# Lib "G32_GD.dll" (ByVal hBars&, ByVal nBar As Long, ByVal strToTimeZone$)

' Returns Session Date for the specified date/time (tomorrow if time is after crossover)
' But when asked to validate, it will return 0 if the date/time is invalid:
' - if the trading session is a Sat or Sun
' - if it's before the custom start time (when start time different from default start)
' - if it's after the custom end time (when end time different from default end)
Declare Function gdBarsSessionDate& Lib "G32_GD.dll" (ByVal hBars&, ByVal dDateTime#, ByVal bValidate&)

' converts bars data to Heikin-Ashi (modifies the OHLC data)
Declare Sub gdBarsConvertToHeikinAshi Lib "G32_GD.dll" (ByVal hBars&)

' builds a spread from other bars
Declare Function gdCalcSpread& Lib "G32_GD.dll" (ByVal hBarsHandles&, ByVal strSpreadDfn$, ByVal uFlags&)


'=========================================
' gdArgs
Declare Function gdGetArgCount& Lib "G32_GD.dll" (ByVal hArgs&)
Declare Function gdGetArgFromBar& Lib "G32_GD.dll" (ByVal hArgs&)
Declare Function gdGetArgDrawingCommands& Lib "G32_GD.dll" (ByVal hArgs&)
Declare Function gdGetArgType& Lib "G32_GD.dll" (ByVal hArgs&, ByVal nArgNum&)
Declare Function gdGetArgAsNumber& Lib "G32_GD.dll" (ByVal hArgs&, ByVal nArgNum&, dNumber#)
Declare Function gdGetArgAsHandle& Lib "G32_GD.dll" (ByVal hArgs&, ByVal nArgNum&, hHandle&)
Declare Function gdGetArgInstanceMemPtr& Lib "G32_GD.dll" (ByVal hArgs&)
Declare Sub gdSetArgInstanceMemPtr Lib "G32_GD.dll" (ByVal hArgs&, ByVal hMemPtr&)

'=========================================
' gdFile stuff
Declare Function gdRegisterFiles Lib "G32_GD.dll" (ByVal strFileSpec$, Totals As Any) As Long
Declare Function gdSetFileDate Lib "G32_GD.dll" (ByVal strFileSpec$, ByVal dDateTime#) As Long

' DLL encrypting function (used by "gdEncrypt")
Private Declare Function gdVBE Lib "G32_GD.dll" (ByVal bEncrypt&, pMemory As Any, ByVal nMemLen&, ByVal hKey&) As Long
' Returns the MD5 value for a chunk of memory
Declare Function gdCalcMemMD5 Lib "G32_GD.dll" (ByVal strMemory As String, ByVal nMemLen&, ByVal strMD5$, ByVal bAsHex&) As Long
' Returns the CRC value for a chunk of memory
Declare Function gdCalcMemCRC32 Lib "G32_GD.dll" (ByVal nMemPtr As Long, ByVal nMemLen As Long) As Long
Declare Function gdCalcCumulativeCRC32 Lib "G32_GD.dll" (ByVal nMemPtr As Long, ByVal nMemLen As Long, ByVal nInitialCRC As Long) As Long
Declare Function gdCalcStrCRC32 Lib "G32_GD.dll" Alias "gdCalcMemCRC32" (ByVal strMemory As String, ByVal nMemLen As Long) As Long
' Calculates the CRC value for a file -- returns true/false for success
Declare Function gdCalcFileCRC32 Lib "G32_GD.dll" (ByVal strFileName As String, CRC As Long) As Long
' Converts memory to/from a hex string:
' - if iMemLen > 0 then builds strHex from pMemory (returns length of hex string)
' - else converts strHex to pMemory (returns length of memory, or -1 if an error)
Declare Function gdHexMemory Lib "G32_GD.dll" (ByVal strHex As String, ByVal nMemPtr As Long, ByVal nMemLen As Long) As Long

' Uses search engine to find all the matching files
' - strFileSpec: filemask with special search options
' - dFlags = bitmask for various options:
'      bit 1 = to include full path with returned filenames
'      bit 2 = to include folder names
'      bit 3 = to include extra tab-delimited fields (name, size, date.time, attribs)
' - hFilesArray: optional string array to be filled with the filenames
' - pTotals: for optional summary information
' - returns the number of matches found
Declare Function gdGetMatchingFiles2 Lib "G32_GD.dll" (ByVal strFileSpec$, ByVal dwFlags&, ByVal hFilenameArray&, MatchedTotals As gdFileMatchingTotals) As Long

' see function: GetMasterFileMatches
Declare Sub gdGetMasterFileMatches Lib "G32_GD.dll" (ByVal hTable&, ByVal strPath$, ByVal strFormat$, ByVal strSymbol$)


' To convert a gdString to a VB string
' - if hArray is a string array, offset is the item # in the array
' - if hArray is a gdString, offset is ignored
Public Function gdGetStr(ByVal hArray&, Optional ByVal offset& = 0) As String

    Dim s$, i&, max_length_guess&

    ' first get length of string
    i = gdGetNum(hArray, offset)
    If i > 0 Then
        ' allocate space for string
        s = Space(i)
        i = DLL_gdGetStr(hArray, offset, s, i)
        If i <> Len(s) Then s = Left(s, i)
    Else
        'OLD METHOD: keep for awhile to be backward compatible
        max_length_guess = 250
        Do While True
            s = Space(max_length_guess + 1)
            i = DLL_gdGetStr(hArray, offset, s, max_length_guess + 1)
            ' make sure max_length_guess was big enough
            If i <= max_length_guess Then
                If i < 0 Then i = 0
                s = Left(s, i)
                Exit Do
            End If
            ' else double size and try again
            max_length_guess = max_length_guess * 2
        Loop
    End If

    gdGetStr = s
End Function

' Formats a price for display purposes
' (will first round price to nearest MinMove)
' sDecimalDigits usage:
'   0  = display as trading units, e.g. 79.09375 = 79^03 for bonds
'   -1 = use default # of digits to right of decimal (based on MinMove)
'   >0 = user-specified # of digits to right of decimal
Public Function gdFormatPrice(ByVal dPrice#, ByVal dTickMove#, _
    Optional ByVal dMinMoveInTicks# = 1#, Optional ByVal sDecimalDigits% = -1) As String
    
    Dim strPrice$, rc&
    strPrice = Space(50)
    rc = gdFormatPriceString(strPrice, Len(strPrice), _
        dPrice, dTickMove, dMinMoveInTicks, sDecimalDigits)
    If rc < 0 Then rc = 0
    gdFormatPrice = Left(strPrice, rc)
        
End Function

' Creates a new gdArray and returns its handle
' (caller is responsible for destroying when done)
Public Function gdCreateArray(ByVal eArrayType As eGdArray_Type, _
        Optional ByVal nSize As Long = 0, _
        Optional ByVal dNullValue# = USE_DEFAULT_NULL) As Long
    
    gdCreateArray = DLL_gdCreateArray(eArrayType, nSize, dNullValue)
End Function

Public Function gdGetObject(ByVal hArray&, ByVal nOffset&) As Object
    Dim v As Variant
    Set v = Nothing
    If gdGetVariant(hArray, nOffset, v) Then
        If VarType(v) = vbEmpty Then Set v = Nothing
    End If
    Set gdGetObject = v
    Set v = Nothing
End Function

Public Function gdSetObject(ByVal hArray&, ByVal nOffset&, obj As Object) As Boolean
    Dim v As Variant
    Set v = obj
    If gdSetVariant(hArray, nOffset, v) Then
        gdSetObject = True
    Else
        gdSetObject = False
    End If
    Set v = Nothing
End Function


Public Function GetPeriodType(ByVal nPeriodicity As Long) As eBarsPeriodType
    GetPeriodType = nPeriodicity And &HFF000000
End Function

Public Function GetPeriodsPerBar(ByVal nPeriodicity As Long) As Long
    GetPeriodsPerBar = nPeriodicity And &HFFFFFF
End Function

Public Function GetPeriodStr(ByVal Periodicity As Variant) As String
    Dim hBars&, hString&
    hBars = gdCreateBars(0, eBARS_Close)
    If VarType(Periodicity) = vbString Then
        gdSetBarsStrProp hBars, eBARS_PeriodicityStr, Periodicity
    Else
        gdSetBarsNumProp hBars, eBARS_Periodicity, Periodicity
    End If
    hString = gdGetBarsStrProp(hBars, eBARS_PeriodicityStr)
    GetPeriodStr = gdGetStr(hString)
    gdDestroyString hString
    gdDestroyBars hBars
End Function

Public Function GetPeriodicity(ByVal strPeriod$) As Long
    Dim hBars&
    hBars = gdCreateBars(0, eBARS_Close)
    gdSetBarsStrProp hBars, eBARS_PeriodicityStr, strPeriod
    GetPeriodicity = gdGetBarsNumProp(hBars, eBARS_Periodicity)
    gdDestroyBars hBars
End Function

Public Function IsIntraday(ByVal nPeriodicity As Long) As Boolean
    
    If nPeriodicity >= ePRD_Days Or nPeriodicity = 0 Then
        IsIntraday = False
    Else
        IsIntraday = True
    End If
    
End Function

' To encrypt/decrypt a chunk of memory
Public Function gdEncrypt(ByVal bEncrypt As Boolean, mbMemory As cMemBuffer, mbPassword As cMemBuffer) As Long

    ' we'll pass the key as a gdArray of doubles
    ' (so can't be easily deciphered if intercepted)
    Dim i&, hKey&, iFlags&, dMult#, dValue#
    hKey = gdCreateArray(eGDARRAY_Doubles, mbPassword.Length, 0)
    dMult = hKey
    For i = 0 To mbPassword.Length - 1
        dMult = dMult + i + 1
        gdSetNum hKey, i, mbPassword.GetByte(i) * dMult + mbMemory.Length
    Next
    If bEncrypt Then
        iFlags = 1
    End If
    gdEncrypt = gdVBE(iFlags, ByVal mbMemory.MemPtr, mbMemory.Length, hKey)
    gdDestroyArray hKey
    
End Function

' Returns the MD5 value for a chunk of memory
' - bAsHex = True:  will return a 32-byte hex string
' - bAsHex = False: will return a 16-byte binary string
Public Function gdCalcMD5(ByVal pFromMemory$, Optional ByVal bAsHex As Boolean = True) As String

    Dim strMD5$, i&
    strMD5 = Space(32)
    i = gdCalcMemMD5(pFromMemory, Len(pFromMemory), strMD5, bAsHex)
    gdCalcMD5 = Left(strMD5, i)

End Function

' Returns number of matching files
Public Function gdNumMatchingFiles(ByVal strFileSpec$, _
        Optional ByVal bIncludeFolderNames As Boolean = False, _
        Optional dTotalBytes As Double = 0) As Long

    Dim MatchedTotals As gdFileMatchingTotals
    Dim dwFlags As Long
    If bIncludeFolderNames Then dwFlags = dwFlags Or &H4
    gdNumMatchingFiles = gdGetMatchingFiles2(strFileSpec, dwFlags, ByVal 0&, MatchedTotals)
    dTotalBytes = MatchedTotals.dMatchedBytes

End Function

Public Function gdGetProfiles(Optional ByVal nFromID& = 0, Optional ByVal nToID& = 999, _
        Optional ByVal strDelimiter$ = vbCrLf) As String
        
    Dim hString&
    hString = gdGetProfilesString(nFromID, nToID, strDelimiter)
    If hString <> 0 Then
        gdGetProfiles = gdGetStr(hString)
        gdDestroyString hString
    End If

End Function

Public Sub gdResetProfiles(Optional ByVal nFromID& = 0, Optional ByVal nToID& = 999)
    DLL_gdResetProfiles nFromID, nToID
End Sub

' This function should be used to get a string from a table (it will destroy the temp string correctly)
Public Function gdGetTableString(ByVal hTable&, ByVal nField&, ByVal nRecord&) As String
    
    Dim hString&
    hString = gdGetTableStr(hTable, nField, nRecord)
    gdGetTableString = gdGetStr(hString)
    gdDestroyString hString

End Function

' Returns one or more "matches" for symbol from MS/CSI master file
' - if Symbol is empty, fills table with all records
' - Symbol can be either a symbol or a filename (e.g. F001.DTA)
' - returns table: 0=Symbol, 1=Period (d/w/m/##), 2=Desc, 3=Filename, 4=Format, 5=File_Num, 6=ConvFact, 7=SecFlag
Public Function GetMasterFileMatches(ByVal strPath$, ByVal strFormat$, Optional ByVal strSymbol$ = "") As cGdTable

    Dim tblMatches As New cGdTable
    gdGetMasterFileMatches tblMatches.TableHandle, strPath, strFormat, strSymbol
    Set GetMasterFileMatches = tblMatches
    
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    EncryptToHex
'' Description: Encrypt a string then convert it to a hexidecimal string
'' Inputs:      String
'' Returns:     Hex String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function EncryptToHex(ByVal strString As String, Optional ByVal strKey As String = "") As String
On Error GoTo ErrSection:

    Dim lLength As Long                 ' Length of the string
    Dim strHex As String                ' Hexidecimal string to return
    Dim mb As New cMemBuffer            ' Membuffer to use for encrypted string
    Dim mbKey As New cMemBuffer         ' Key for encryption

    If Len(strString) > 0 Then
        ' set the encryption key
        If Len(strKey) > 0 Then
            mbKey.Buffer = strKey
        Else ' default key:
            mbKey.PutByte 71
            mbKey.PutByte 202
            mbKey.PutByte 123
            mbKey.PutByte 63
            mbKey.PutByte 176
            mbKey.PutByte 2
            mbKey.PutByte 70
            mbKey.PutByte 198
            mbKey.PutByte 169
            mbKey.PutByte 85
            mbKey.PutByte 10
        End If
        ' encrypt the string
        mb.Buffer = strString
        gdEncrypt True, mb, mbKey
        ' convert the encrypted buffer to hex
        strHex = Space(mb.Length * 2 + 2)
        lLength = gdHexMemory(strHex, ByVal mb.MemPtr, mb.Length)
        If lLength > 0 Then
            EncryptToHex = Left(strHex, lLength)
        End If
    End If
       
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mGdDll.EncryptToHex"
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Function:    DecryptFromHex
'' Description: Convert a hex string to a string, then decrypt it
'' Inputs:      Hex String
'' Returns:     String
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DecryptFromHex(ByVal strHex As String, Optional ByVal strKey As String = "") As String
On Error GoTo ErrSection:

    Dim lLength As Long                 ' Length of the string
    Dim mb As New cMemBuffer            ' Mem buffer for decrypting
    Dim mbKey As New cMemBuffer         ' Key for decryption

    If Len(strHex) > 0 And Len(strHex) Mod 2 = 0 Then
        ' set the encryption key
        If Len(strKey) > 0 Then
            mbKey.Buffer = strKey
        Else ' default key:
            mbKey.PutByte 71
            mbKey.PutByte 202
            mbKey.PutByte 123
            mbKey.PutByte 63
            mbKey.PutByte 176
            mbKey.PutByte 2
            mbKey.PutByte 70
            mbKey.PutByte 198
            mbKey.PutByte 169
            mbKey.PutByte 85
            mbKey.PutByte 10
        End If
        
        lLength = Len(strHex) \ 2
        mb.Length = lLength
        gdHexMemory strHex, ByVal mb.MemPtr, 0
        gdEncrypt False, mb, mbKey
        DecryptFromHex = mb.Buffer
    End If
       
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mGdDll.DecryptFromHex"
End Function

#If 0 Then
' TLB: this was the old function, but it didn't actually work as well for some things,
' e.g. calling RoundToMinMove(1.39728, 0.00001) did not equal Val("1.39728"),
' whereas the C++ version uses SigDigits when inverting MinMove and thus works better
Public Function RoundToMinMove_OLD(ByVal dValue#, ByVal dMinMove#) As Double
On Error GoTo ErrSection:
   
    Dim dReturn As Double               ' Return value from the function
    
    If dMinMove >= 1 Then
        dReturn = Int(CDbl(dValue / dMinMove) + 0.5) * dMinMove
    ElseIf dMinMove > 0 Then
        ' If the min move happens to be below 1 (especially 0.1), it works much
        ' better to multiply first then divide in the line below because of
        ' rounding issues, so invert the min move here.  11/16/2006 DAJ
        dMinMove = 1 / dMinMove
        dReturn = Int(CDbl(dValue * dMinMove) + 0.5) / dMinMove
    Else
        dReturn = dValue
    End If
    
    RoundToMinMove_OLD = dReturn
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mGdDll.RoundToMinMove_OLD"
End Function

Public Function RoundToMinMove(ByVal dValue#, ByVal dMinMove#) As Double
On Error GoTo ErrSection:
   
    RoundToMinMove = gdRoundPriceToMinMove(dValue, dMinMove)
    
    If IsIDE Then
        If dMinMove > 0.000011 Then
            If RoundToMinMove <> RoundToMinMove_OLD(dValue, dMinMove) Then
                DebugLog "RoundToMinMove difference for: " & Str(dValue) & ", " & Str(dMinMove)
                'InfBox Str(dValue) & ", " & Str(dMinMove), , , "MinMove diff for ..."
            End If
        End If
    End If
    
ErrExit:
    Exit Function
    
ErrSection:
    RaiseError "mGdDll.RoundToMinMove"
End Function
#End If

' To write a line of text to a file opened with gdFileOpen
' (returns # of bytes written)
Public Function gdFileWriteLine(ByVal hFile&, ByVal strLine$, Optional ByVal bFlush As Boolean = False) As Long

    If hFile = 0 Then
        gdFileWriteLine = 0
    Else
        ' need to convert vbCrLf to vbLf for C++ file i/o to work correctly
        strLine = Replace(strLine, vbCrLf, vbLf) & vbLf
        gdFileWriteLine = gdFileStringIO(hFile, strLine, Len(strLine), True)
        If bFlush Then
            gdFileFlush hFile
        End If
    End If

End Function
