Attribute VB_Name = "Osiris_Review_Gadgets"
'
'   Description: A module containing Osiris data review associated gadgets
'
'   Date: 2024/4/17
'   Author: maoyi.fan@yapro.com.tw
'   Ver.: 0.1c
'   Revision History:
'       - 2024/4/16, 0.1c: Added function DoComparableQuartile() to calculate quartile numbers
'                          of comparable companies and other minor fixes
'       - 2024/4/15, 0.1b: First added
'
'   ToDo's:
'       1)

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
    Dim comparableCount                                                         As Long
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

        If comparableFlag = "Yes" Then
            ReDim Preserve PLI(comparableCount)
            PLI(comparableCount) = CDbl(PLIValue)
            comparableCount = comparableCount + 1
        End If
    Next r
    
    'Debug.Print "Number of comparable company: " & CStr(comparableCount)
    'Debug.Print "Comparable List: "
    'For r = 0 To comparableCount - 1
    '    Debug.Print "PLI(" & CStr(r) & "): " & PLI(r)
    'Next r
    ' Update customized variable of data type Quartile_Data_Type
    With q
        .minQuartile = WorksheetFunction.Quartile(PLI, 0)
        .lowerQuartile = WorksheetFunction.Quartile(PLI, 1)
        .medianQuartiile = WorksheetFunction.Quartile(PLI, 2)
        .upperQuartile = WorksheetFunction.Quartile(PLI, 3)
        .maxQuartile = WorksheetFunction.Quartile(PLI, 4)
    End With
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
    
    Set selectedRange = screenWorksheet.Range(Osiris_Review_Constant.CONST_BASE_RANGE)
    lRow = FindMaximumRow(selectedRange)
    Debug.Print "Maximum number of rows I found: " & lRow
    
    With ss
        .conditionCount = 0
        .okCount = 0
        .rejectCount = 0
        .unscreenedCount = 0
        .totalCount = 0
    End With
    For r = 3 To lRow
        Set selectedRange = screenWorksheet.Cells(r, Osiris_Review_Constant.CONST_STATUS_COLUMN)
        If selectedRange.Value = Osiris_Review_Constant.CONST_COMPARABLE_STATE_CONDITION Then
            ss.conditionCount = ss.conditionCount + 1
        ElseIf selectedRange.Value = Osiris_Review_Constant.CONST_COMPARABLE_STATE_NG Then
            ss.rejectCount = ss.rejectCount + 1
        ElseIf selectedRange.Value = Osiris_Review_Constant.CONST_COMPARABLE_STATE_OK Then
            ss.okCount = ss.okCount + 1
        ElseIf selectedRange.Value = Osiris_Review_Constant.CONST_COMPARABLE_STATE_TBD Then
            ss.unscreenedCount = ss.unscreenedCount + 1
        ElseIf AscW(selectedRange.Value) = Osiris_Review_Constant.UNICODE_CHECK Then
            ss.unscreenedCount = ss.unscreenedCount + 1
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
' Description: Returns comparable state label, especially handles those non-ASCII code included in Osiris database search results
' Code Date: 2024/04/14
'
Function ReturnStateLabel(ByVal comparableState As String) As String
    Debug.Print "Comparable state string: " & AscW(comparableState)
    If AscW(comparableState) = Osiris_Review_Constant.UNICODE_CHECK Then
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
