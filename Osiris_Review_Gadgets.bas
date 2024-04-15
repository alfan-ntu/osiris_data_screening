Attribute VB_Name = "Osiris_Review_Gadgets"
'
'   Description: A module containing Osiris data review associated gadgets
'
'   Date: 2024/4/15
'   Author: maoyi.fan@yapro.com.tw
'   Ver.: 0.1b
'   Revision History:
'       - 2024/4/15, 0.1b: First added
'
'   ToDo's:
'       1)

'
' Description: Check the column "N" of Screening_Worksheet and determine number of OK companies
'
Function NumberOfOK()

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
    
    lRow = targetRange.End(xlDown).Row
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


