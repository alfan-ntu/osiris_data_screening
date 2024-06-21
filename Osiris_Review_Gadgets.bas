Attribute VB_Name = "Osiris_Review_Gadgets"
'
'   Description: A module containing Osiris data review associated gadgets
'
'   Date: 2024/6/16
'   Author: maoyi.fan@yapro.com.tw
'   Ver.: 0.1i
'   Revision History:
'       - 2024/6/16, 0.1i: Adjusted record 'Status' related process due to inconsistent report format
'       - 2024/6/15, 0.1h: Adjusted constant arrangement to accommodate dual operation conditions
'       - 2024/6/13, 0.1g: Fixed the issue jumping to the first unscreened record when all records have been screened
'       - 2024/5/14, 0.1f: Created Screening_Worksheet and populate comparable state formula, country code... in
'                          PLI Screening Worksheet
'       - 2024/4/23, 0.1e: Added gadgets to find range of selected area
'       - 2024/4/16, 0.1c: Added function DoComparableQuartile() to calculate quartile numbers
'                          of comparable companies and other minor fixes
'       - 2024/4/15, 0.1b: First added
'
'   ToDo's:
'       1) Issue: condition to identify a unscreened record more than just checking if the status column is a check mark.
'                 In VBA, an uninitialized cell or variable is actually an Emtpy variable. Need to check IsEmtpy in the
'                 function ScreenStatistics()
'
Option Explicit

'
' Description: Check the column "N" of Screening_Worksheet and determine number of OK companies
'
Type Screening_Statistics
    okCount             As Integer
    conditionCount      As Integer
    rejectCount         As Integer
    unscreenedCount     As Integer
    totalCount          As Integer
End Type

'
' Description: Type definition for PLI Quartiles
'
Type Quartile_Data_Type
    minQuartile         As Double
    lowerQuartile       As Double
    medianQuartiile     As Double
    upperQuartile       As Double
    maxQuartile         As Double
End Type

'
' Description: Type of range and validated status
'
Type Validated_Range
    srcRangeStr         As String
    valid               As Boolean
End Type

'
' Description: Country code dictionary
'
Public countryCodeDict As New Scripting.Dictionary


'
' Description: setup the country code dictionary
' Coding Date: 2024/5/14
'
Sub setupCountryCodeDictionary()
    countryCodeDict.Add Key:="TW", Item:="台灣"
    countryCodeDict.Add Key:="CN", Item:="中國"
    countryCodeDict.Add Key:="IN", Item:="印度"
    countryCodeDict.Add Key:="TH", Item:="泰國"
    countryCodeDict.Add Key:="HK", Item:="香港"
    countryCodeDict.Add Key:="JP", Item:="日本"
    countryCodeDict.Add Key:="ID", Item:="印尼"
    countryCodeDict.Add Key:="AU", Item:="澳大利亞"
    countryCodeDict.Add Key:="KR", Item:="韓國"
    countryCodeDict.Add Key:="VN", Item:="越南"
    countryCodeDict.Add Key:="MY", Item:="馬來西亞"
    countryCodeDict.Add Key:="SG", Item:="新加坡"
    countryCodeDict.Add Key:="NZ", Item:="紐西蘭"
    ' Debug.Print "Country code list of size " & countryCodeDict.Count & " is created!"
End Sub


'
' Description: Look up country name in Chinese for country ISO Code using Scripting.Dictionary
' Coding Date: 2024/5/14
'
Function findCountryNameInChinese(ByVal cntyISOCode As String) As String
    findCountryNameInChinese = Osiris_Review_Gadgets.countryCodeDict(cntyISOCode)
End Function


'
' Description: Exercise DoComparableQuartile subroutine
'
Sub ComparableQuartileTest()
    Dim targetWorksheet         As Worksheet
    Dim targetWorksheetName     As String
    Dim targetRange             As Range
    Dim q                       As Quartile_Data_Type
    
    targetWorksheetName = Osiris_Review_Constant.OM_DETAILS_SHEET
    Set targetWorksheet = Worksheets(targetWorksheetName)
    '
    ' Actual exerciser has to determine the actual range passed to the function
    ' DoComparableQuartile()'
    
    Set targetRange = targetWorksheet.Range("D15:I80")
    
    q = DoComparableQuartile(targetRange, Osiris_Review_Constant.BMK_CURRENT_YEAR)
    Debug.Print "Min: " & q.minQuartile
    Debug.Print "Lower Quartile: " & q.lowerQuartile
    Debug.Print "Median Quartile: " & q.medianQuartiile
    Debug.Print "Upper Quartile: " & q.upperQuartile
    Debug.Print "Max: " & q.maxQuartile

End Sub

'
' Description: Calculate the quartile numbers of comparables
' Coding Date: 2024/4/16
' Input:
'   - benchmarkRange: Range of the PLI average, values of last three years and comparable flag column;
'   - yearIndex: 0: average; 1: current year; 2: last year; 4: year before last year
'
Function DoComparableQuartile(benchmarkRange As Range, yearIndex As Integer) As Quartile_Data_Type
    Dim tmpRange                                                                As Range
    Dim PLI()                                                                   As Double
    Dim minPLI, lowerQuartilePLI, medianQuartilePLI, upperQuartilePLI, maxPLI   As Double
    Dim comparableCount, rowCount, r                                            As Long
    Dim PLIValue, comparableFlag                                                As String
    Dim PLIOffset, comparableOffset                                             As Integer
    Dim q                                                                       As Quartile_Data_Type
    
    ' Set the Top-Left cell of the benchmark range
    Set tmpRange = benchmarkRange.Cells(1, 1)
    rowCount = benchmarkRange.Rows.Count
    comparableCount = 0
    
    ' Column offset per yearIndex
    PLIOffset = yearIndex
    comparableOffset = Osiris_Review_Constant.BMK_COMPARABLE_OFFSET
    '
    ' Traverse the benchmark rows to form PLI list of comparable companies
    '
    For r = 1 To rowCount
        If IsNumeric(benchmarkRange.Cells(r, PLIOffset + 1)) Then
            PLIValue = Format(benchmarkRange.Cells(r, PLIOffset + 1).Value, "##0.00")
        Else
            PLIValue = benchmarkRange.Cells(r, PLIOffset + 1).Value
        End If
        comparableFlag = benchmarkRange.Cells(r, comparableOffset + 1).Value
        ' update PLI array if the company is categorized as a comparable company
        If comparableFlag = Osiris_Review_Constant.CONST_COMPARABLE_STATE_OK Then
            ReDim Preserve PLI(comparableCount)
            PLI(comparableCount) = CDbl(PLIValue)
            comparableCount = comparableCount + 1
        End If
    Next r
    
    ' Update customized variable of data type Quartile_Data_Type
    If comparableCount > 0 Then
        With q
            .minQuartile = WorksheetFunction.Quartile(PLI, 0)
            .lowerQuartile = WorksheetFunction.Quartile(PLI, 1)
            .medianQuartiile = WorksheetFunction.Quartile(PLI, 2)
            .upperQuartile = WorksheetFunction.Quartile(PLI, 3)
            .maxQuartile = WorksheetFunction.Quartile(PLI, 4)
        End With
    Else
        With q
            .minQuartile = 0
            .lowerQuartile = 0
            .medianQuartiile = 0
            .upperQuartile = 0
            .maxQuartile = 0
        End With
    End If
    DoComparableQuartile = q
End Function

'
' Description: do Screening_Worksheet screening statistics and return an enumeration
'              of Screening_Statistics
' Coding Date: 2024/4/15
'
Function ScreenStatistics(ByVal screenWorksheet As Worksheet) As Screening_Statistics
    Dim ss As Screening_Statistics
    Dim selectedRange As Range
    Dim lRow, r As Long
    Dim dbg As Long
    
    Set selectedRange = screenWorksheet.Range(Osiris_Review_Constant.SCREENING_WORKSHEET_BASE_RANGE)
    lRow = FindMaximumRow(selectedRange)
    ' Debug.Print "Maximum number of rows I found: " & lRow
    
    With ss
        .conditionCount = 0
        .okCount = 0
        .rejectCount = 0
        .unscreenedCount = 0
        .totalCount = 0
    End With
    
    dbg = 0
    
    For r = 3 To lRow
        Set selectedRange = screenWorksheet.Cells(r, Osiris_Review_Constant.SCREENING_WORKSHEET_STATUS_COLUMN)
        If selectedRange.Value = Osiris_Review_Constant.CONST_COMPARABLE_STATE_CONDITION Then
            ss.conditionCount = ss.conditionCount + 1
        ElseIf selectedRange.Value = Osiris_Review_Constant.CONST_COMPARABLE_STATE_NG Then
            ss.rejectCount = ss.rejectCount + 1
        ElseIf selectedRange.Value = Osiris_Review_Constant.CONST_COMPARABLE_STATE_OK Then
            ss.okCount = ss.okCount + 1
        ElseIf selectedRange.Value = Osiris_Review_Constant.CONST_COMPARABLE_STATE_TBD Then
            ss.unscreenedCount = ss.unscreenedCount + 1
        ElseIf selectedRange.Value = Osiris_Review_Constant.CONST_COMPARABLE_STATE_EMPTY Then
            ss.unscreenedCount = ss.unscreenedCount + 1
        ElseIf AscW(selectedRange.Value) = Osiris_Review_Constant.UNICODE_CHECK Then
            '
            ' Issue: condition to identify a unscreened record more than just checking if the
            '      status column is a check mark
            ' ToDo (2024/6/19): to fix this bug
            '
            ss.unscreenedCount = ss.unscreenedCount + 1
        Else
            dbg = dbg + 1
            Debug.Print "Debug(" & Str(dbg) & ");  problematic row: " & Str(r)
        End If
    Next r
    ScreenStatistics = ss
End Function

'
' Description: returns the minimum allowable row number to ensure move previous functioning correlctly
'
Function FindMinimumRow(ByVal targetRange As Range) As Long
    FindMinimumRow = 3
End Function

'
' Description: returns the maximum allowable row number to ensure move next functioning correlctly
'
Function FindMaximumRow(ByVal targetRange As Range) As Long
    Dim lRow As Long
    
    ' Boundary check
    If targetRange.Offset(1, 0).Value = "" Then
        lRow = targetRange.Row
    Else
        lRow = targetRange.End(xlDown).Row
    End If
    FindMaximumRow = lRow
End Function


'
' Description: Return bottom-right cell of a range
' Coding Date: 2024/4/23
'
Function getBottomRightCell(ByVal targetWs As Worksheet, ByVal rt As Range) As Range
    Dim lRow, lCol  As Long
    Dim targetRange As Range
    
    lRow = rt.End(xlDown).Row
    lCol = rt.End(xlToRight).Column
    Set getBottomRightCell = targetWs.Cells(lRow, lCol)
End Function

'
' Description: Returns comparable state label, especially handles those non-ASCII code included in Osiris database search results
' Code Date: 2024/06/16
'
Function ReturnStateLabel(ByVal comparableState As String) As String
    ' return comparable state code according to the non-ASCII code selected by Osiris database
    If IsEmpty(comparableState) Or comparableState = Osiris_Review_Constant.CONST_COMPARABLE_STATE_EMPTY Then
        ReturnStateLabel = Osiris_Review_Constant.CONST_COMPARABLE_STATE_TBD    ' added on 2024/6/16
    ElseIf AscW(comparableState) = Osiris_Review_Constant.UNICODE_CHECK Then
        ReturnStateLabel = Osiris_Review_Constant.CONST_COMPARABLE_STATE_TBD
    ElseIf AscW(comparableState) = Osiris_Review_Constant.UNICODE_FORBIDDEN Then
        ReturnStateLabel = Osiris_Review_Constant.CONST_COMPARABLE_STATE_NG
    Else
        ReturnStateLabel = comparableState
    End If
End Function


'
' Description: Converts PLI label to PLI switch to support calling company screening row-by-row
'
Function PLILabelToSwitch(PLILabel As String) As String
    If PLILabel = Osiris_Review_Constant.CONST_OM_PLI_LABEL Then
        PLILabelToSwitch = Osiris_Review_Constant.CONST_OM_PLI
    Else
        PLILabelToSwitch = Osiris_Review_Constant.CONST_NCP_PLI
    End If
End Function

'
' Description: remove non-printable characters from the original string
'
Public Function CleanMessyString(ByVal Str As String) As String

    'Removes non-printable characters from a string

    Dim CleanString As String
    Dim i As Integer

    CleanString = Str
    For i = Len(CleanString) To 1 Step -1
        'Debug.Print Asc(Mid(Str, i, 1))

        Select Case Asc(Mid(Str, i, 1))
            Case 1 To 31, Is >= 127
                'Bad stuff
                'https://www.ionos.com/digitalguide/server/know-how/ascii-codes-overview-of-all-characters-on-the-ascii-table/
                CleanString = Left(CleanString, i - 1) & " " & Mid(CleanString, i + 1)
            Case Else
                'Keep

        End Select
    Next i
    CleanMessyString = CleanString
End Function

'
' Description: returns column letter for a given column number
' Source: https://stackoverflow.com/questions/12796973/function-to-convert-column-number-to-letter
'
Function Col_Letter(lngCol As Long) As String
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function

'
' Description: Scans the company column of the worksheet Osiris_Review_Constant.MASTER_SHEET, i.e. "列表 (2)" or "Screening_Worksheet"
'              to determine the number of companies
' ToDo's: Can be stored as a global instance variable when migrating this project to a class-based implementation
'
Public Function FindNumberOfCompanies() As Integer
    Dim tgtWs           As Worksheet
    Dim selectedRange   As Range
    Dim nc              As Integer
    
    Set tgtWs = Worksheets(Osiris_Review_Constant.MASTER_SHEET)
    Set selectedRange = tgtWs.Range(Osiris_Review_Constant.SCREENING_WORKSHEET_BASE_RANGE)
    nc = Osiris_Review_Gadgets.FindMaximumRow(selectedRange) - 2
    
    FindNumberOfCompanies = nc
    
End Function

'
' Description: Locate the first unscreened company record
' Coding Date: 2024/6/13
'   2024/6/13: fixed check out of bound issue when all records are all screened
'
Public Function findFirstUnscreenRecord() As Long
    Dim tgtWs           As Worksheet
    Dim lRow, r         As Long
    Dim compStat        As Range
    Dim wsName          As String
    Const firstDataRow  As Long = 3
    Dim firstRecordFound As Boolean
    
    r = firstDataRow
    Set tgtWs = Worksheets(Osiris_Review_Constant.SCREENING_SHEET)
    ' Set compStat = tgtWs.Range(Osiris_Review_Constant.SCREENING_WORKSHEET_STATUS_COLUMN & CStr(r))
    Set compStat = tgtWs.Range(Osiris_Review_Constant.SCREENING_WORKSHEET_COMPANY_NAME_COLUMN & CStr(r))
    firstRecordFound = False
    
    lRow = Osiris_Review_Gadgets.FindMaximumRow(compStat)
    wsName = compStat.Worksheet.Name
    
    For r = firstDataRow To lRow
        Set compStat = tgtWs.Range(Osiris_Review_Constant.SCREENING_WORKSHEET_STATUS_COLUMN & CStr(r))
'        If AscW(compStat.Value) = Osiris_Review_Constant.UNICODE_CHECK Then
        If recordYetReviewed(compStat.Value) Then
            firstRecordFound = True
            Debug.Print "Unscreened record: " & CStr(r)
            Exit For
        End If
    Next r
    If firstRecordFound = False Then
        r = 0
    End If
    Set compStat = Nothing
    Set tgtWs = Nothing
    findFirstUnscreenRecord = r
End Function

'
' Description: Determine if SCREENING_WORKSHEET_STATUS_COLUMN of the record marks a reviewed flag
'              i.e. CONST_COMPARABLE_STATE_TBD, CONST_COMPARABLE_STATE_NG, CONST_COMPARABLE_STATE_OK
'                   or CONST_COMPARABLE_STATE_CONDITION
' Coding Date: 2024/6/16
'
Public Function recordYetReviewed(stateCode As String)
    Dim yetr As Boolean
    
    If StrComp(stateCode, Osiris_Review_Constant.CONST_COMPARABLE_STATE_CONDITION, vbTextCompare) = 0 Then
        yetr = False
    ElseIf StrComp(stateCode, Osiris_Review_Constant.CONST_COMPARABLE_STATE_OK, vbTextCompare) = 0 Then
        yetr = False
    ElseIf StrComp(stateCode, Osiris_Review_Constant.CONST_COMPARABLE_STATE_NG, vbTextCompare) = 0 Then
        yetr = False
    ElseIf StrComp(stateCode, Osiris_Review_Constant.CONST_COMPARABLE_STATE_TBD, vbTextCompare) = 0 Then
        yetr = False
    Else
        yetr = True
    End If
    recordYetReviewed = yetr
End Function
