Attribute VB_Name = "Common_Utilities"
'
'   Description: A group of funtion utilities or gadgets common to all Excel workbooks
'
'   Date: 2024/4/11
'   Author: maoyi.fan@yapro.com.tw
'   Ver.: 0.1a
'   Revision History:
'       - 2024/4/11, 0.1a: initial version
'
'   Function List:
'       1) ListSheets (Ctrl-Shift-L): Creates a worksheet 'Worksheet List' at the beginnning of all other worksheets,
'                                     listing all worksheets within the Active workbook
'
'       2) RedCrossout (Ctrl-Shift-C): Cross-outs the selected cells and changing the font color to red
'
'       3) WrapCell (Ctrl-Shift-W) : Wrap or Dewrap the selected cell
'
'       4) GotoFirstSheet (Ctrl-Shift-B): Goes to the very first worksheet of the active workbook
'
'       5) CompanyOMDetails (Ctrl-Shift-O): This is more an Osiris data screening utilities. It looks up the
'                                           Operating Margin financial data of the company in the selected cell
'       6) CompanyNCPDetails (Ctrl-Shift-N): This is more an Osiris data screening utilities. It looks up the
'                                           Net Cost Plus financial data of the company in the selected cell
'
Option Explicit
Public Const WORKSHEET_LIST As String = "Worksheet List"
Public Const WORKSHEET_INDEX_COLUMN     As String = "A"
Public Const WORKSHEET_NAME_COLUMN      As String = "B"
Public Const WORKSHEET_VISIBLE_COLUMN   As String = "C"
'
' Constant definitions associated with data screening of Osiris search/rejection results
' Note: The following two constants might be different from company to company
'
Public Const OM_DETAILS_SHEET           As String = "Benchmark 1"
Public Const NCP_DETAILS_SHEET          As String = "Benchmark 4"


Public newlyCreated As Boolean

'
' Description: Display the Net Cost Plus financial data of the selected company
' Keyboard Shortcut: (Ctrl-Shift-N)
' Code Date: 2024/4/12
'
Sub CompanyNCPDetails()
    Dim companyName, displayMessage, Response As String
    Dim targetWorksheet As Worksheet
    Dim targetWorksheetName As String
    Dim NCP_Title, NCPMinus1_Title, NCPMinus2_Title As String
    Dim NCP_average, NCP, OM_minus_1, OM_minus_2 As String
    Dim selectedRange, tempRange As Range
    Dim lRow, targetRow, r As Long
    
    companyName = ActiveCell.Value
    targetWorksheetName = "Benchmark 1"
    Set targetWorksheet = Worksheets(targetWorksheetName)
    OM_minus_1 = targetWorksheet.Range("D13").Value
    Set selectedRange = targetWorksheet.Range("B15")
    
    OM_Title = CStr(targetWorksheet.Cells(4, "E").Value)
    OM_Title = CleanMessyString(OM_Title)
    OMMinus1_Title = targetWorksheet.Cells(4, "F").Value
    OMMinus1_Title = CleanMessyString(OMMinus1_Title)
    OMMinus2_Title = targetWorksheet.Cells(4, "H").Value
    OMMinus2_Title = CleanMessyString(OMMinus2_Title)
    lRow = selectedRange.End(xlDown).Row
    For r = 1 To lRow
        Set tempRange = targetWorksheet.Cells(r, "B")
        If companyName = tempRange.Value Then
            OM_average = Format(targetWorksheet.Cells(r, "D").Value, "##0.00")
            OM = Format(targetWorksheet.Cells(r, "E").Value, "##0.00")
            OM_minus_1 = Format(targetWorksheet.Cells(r, "F").Value, "##0.00")
            OM_minus_2 = Format(targetWorksheet.Cells(r, "H").Value, "##0.00")
            Exit For
        End If
    Next r
    '
    ' Populate data to the UserForm
    '
    PLIDetailsForm.tbCompanyName.Value = companyName
    PLIDetailsForm.fyLabel.Value = OM_Title
    PLIDetailsForm.fyminus1Label.Value = OMMinus1_Title
    PLIDetailsForm.fyMinus2Label.Value = OMMinus2_Title
    PLIDetailsForm.tbOMAverage.Value = OM_average
    PLIDetailsForm.tbOM.Value = OM
    PLIDetailsForm.tbOMMinus1.Value = OM_minus_1
    PLIDetailsForm.tbOMMinus2.Value = OM_minus_2
    ' Display the UserForm
    PLIDetailsForm.Show
    
End Sub

'
' Description: Display the Operating Margin financial data of the selected company
' Keyboard Shortcut: (Ctrl-Shift-O)
' Code Date: 2024/4/11
'
Sub CompanyOMDetails()
Attribute CompanyOMDetails.VB_ProcData.VB_Invoke_Func = "O\n14"
    Dim companyName, PLIString As String
    Dim targetWorksheet As Worksheet
    Dim targetWorksheetName As String
    Dim OM_Title, OMMinus1_Title, OMMinus2_Title As String
    Dim OM_average, OM, OM_minus_1, OM_minus_2 As String
    Dim selectedRange, tempRange As Range
    Dim lRow, targetRow, r As Long
    
    companyName = ActiveCell.Value
    targetWorksheetName = "Benchmark 1"
    PLIString = "營業淨利率"
    
    Set targetWorksheet = Worksheets(targetWorksheetName)
    Set selectedRange = targetWorksheet.Range("B15")
    
    OM_Title = CStr(targetWorksheet.Cells(4, "E").Value)
    OM_Title = CleanMessyString(OM_Title)
    OMMinus1_Title = targetWorksheet.Cells(4, "F").Value
    OMMinus1_Title = CleanMessyString(OMMinus1_Title)
    OMMinus2_Title = targetWorksheet.Cells(4, "H").Value
    OMMinus2_Title = CleanMessyString(OMMinus2_Title)
    lRow = selectedRange.End(xlDown).Row
    For r = 1 To lRow
        Set tempRange = targetWorksheet.Cells(r, "B")
        If companyName = tempRange.Value Then
            OM_average = Format(targetWorksheet.Cells(r, "D").Value, "##0.00")
            OM = Format(targetWorksheet.Cells(r, "E").Value, "##0.00")
            OM_minus_1 = Format(targetWorksheet.Cells(r, "F").Value, "##0.00")
            OM_minus_2 = Format(targetWorksheet.Cells(r, "H").Value, "##0.00")
            Exit For
        End If
    Next r
    '
    ' Populate data to the UserForm
    '
    PLIDetailsForm.tbCompanyName.Value = companyName
    PLIDetailsForm.lblPLI.Caption = ""
    PLIDetailsForm.fyLabel.Value = OM_Title
    PLIDetailsForm.fyminus1Label.Value = OMMinus1_Title
    PLIDetailsForm.fyMinus2Label.Value = OMMinus2_Title
    PLIDetailsForm.tbOMAverage.Value = OM_average
    PLIDetailsForm.tbOM.Value = OM
    PLIDetailsForm.tbOMMinus1.Value = OM_minus_1
    PLIDetailsForm.tbOMMinus2.Value = OM_minus_2
    ' Display the UserForm
    PLIDetailsForm.Show
    
End Sub


'
' Description: Go to the first worksheet of the active workbook
' Keyboard Shortcut: (Ctrl-Shift-B)
' Code Date: 2024/4/11
'
'
Sub GotoFirstSheet()
Attribute GotoFirstSheet.VB_ProcData.VB_Invoke_Func = " \n14"
    Debug.Print "GotoFirstSheet macro is executed.."
    ActiveWorkbook.Worksheets(1).Activate
End Sub


'
' Description: listing all the worksheet names in
' Shortcut: Ctrl-Shift-L
' Code Date: 2024/4/11
'
Sub ListSheets()
Attribute ListSheets.VB_Description = "List names of all worksheets"
Attribute ListSheets.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim targetWs, listWs As Worksheet
    Dim x As Integer
    Dim targetRange As Range
    Dim wsExist As Boolean
    Dim targetWorksheetName As String
    Dim currDir As String
    
    targetWorksheetName = WORKSHEET_LIST
    ' Debug information
    Debug.Print "I am " & ThisWorkbook.Name
    Debug.Print "I am at " & ActiveSheet.Name
    ' Do sanity check of the working worksheet
    wsExist = worksheetExists(targetWorksheetName)
    If wsExist Then
        ' Debug.Print targetWorksheetName & " exists!"
        newlyCreated = False
    Else
        Debug.Print targetWorksheetName & " does not exist and a new one is created!"
        Sheets.Add(Before:=Sheets(1)).Name = targetWorksheetName
        newlyCreated = True
    End If
    
    Set listWs = Sheets(targetWorksheetName)
    listWs.Range("A:C").Clear
    
    x = 1
    ' Column A: Worksheet index column
    listWs.Cells(x, WORKSHEET_INDEX_COLUMN) = "Index"
    listWs.Cells(x, WORKSHEET_INDEX_COLUMN).HorizontalAlignment = xlCenter
    ' Column B: Worksheet name column
    listWs.Cells(x, WORKSHEET_NAME_COLUMN) = "Worksheet Name"
    listWs.Cells(x, WORKSHEET_NAME_COLUMN).HorizontalAlignment = xlCenter
    ' Column C: Worksheet hide or not column
    listWs.Cells(x, WORKSHEET_VISIBLE_COLUMN) = "Visble"
    listWs.Cells(x, WORKSHEET_VISIBLE_COLUMN).HorizontalAlignment = xlCenter
    '
    ' Traverse all the worksheets and get worksheet names of them all
    '
    x = x + 1
    For Each targetWs In Worksheets
        listWs.Cells(x, WORKSHEET_INDEX_COLUMN) = x - 1
        listWs.Cells(x, WORKSHEET_NAME_COLUMN) = targetWs.Name
        If targetWs.Visible = True Then
            listWs.Cells(x, WORKSHEET_VISIBLE_COLUMN) = "Yes"
        Else
            listWs.Cells(x, WORKSHEET_VISIBLE_COLUMN) = "No"
        End If
        
        'If newlyCreated Then
        '    listWs.Cells(x, Constants_Module.WORKSHEET_HIDE_COLUMN) = "No"
        'End If
        Set targetRange = listWs.Cells(x, WORKSHEET_NAME_COLUMN)
        listWs.Hyperlinks.Add Anchor:=targetRange, _
                              Address:="", _
                              SubAddress:="'" & targetWs.Name & "'!A1", _
                              TextToDisplay:=targetWs.Name
        ' Hide unwanted worksheets
        'If listWs.Cells(x, Constants_Module.WORKSHEET_HIDE_COLUMN) = "Yes" Then
        '    Sheets(targetWs.Name).Visible = False
        '    Debug.Print "Hide the worksheet: " & targetWs.Name
        'Else
        '    Sheets(targetWs.Name).Visible = True
        'End If
        x = x + 1
    Next targetWs
    
    Sheets(targetWorksheetName).Activate
    '
    ' Adjust the worksheet display
    '
    Call SetColumnWidth(WORKSHEET_NAME_COLUMN, 0)
    Call SetColumnWidth(WORKSHEET_VISIBLE_COLUMN, 10)

End Sub


'
' Description: RedCrossout Macro, crossing out the selected cells using red color
' Keyboard Shortcut: Ctrl-Shift-C
' Code Date: 2024/4/11
'
Sub RedCrossout()
Attribute RedCrossout.VB_ProcData.VB_Invoke_Func = " \n14"
    With Selection.Font
        .Name = "????"
        .FontStyle = "Regular"
        .Size = 12
        .Strikethrough = True
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .Color = 255
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
End Sub


'
' Description: Wrap or Dewrap the selected cell
' Shortcut: Ctrl-Shift-W
' Code Date: 2024/4/11
'
Sub WrapCell()
Attribute WrapCell.VB_ProcData.VB_Invoke_Func = " \n14"
    ActiveCell.WrapText = Not ActiveCell.WrapText
End Sub


'
' Description: changing the width of the specified column
' Arguments:
'      - colIndex: integer, specifies the column to adjust width
'      - w : integer, 0 to autofit; otherwise, number of character width
'
Sub SetColumnWidth(colIndex As String, w As Integer)
    If w = 0 Then
        Columns(colIndex).AutoFit
    Else
        Columns(colIndex).ColumnWidth = w
    End If
End Sub


'
' Description: check if a specified worksheet exists or not
'
Function worksheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    
    worksheetExists = False
    
    ' For Each ws In ThisWorkbook.Worksheets
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name = sheetName Then
            Debug.Print "ws.Name:" & ws.Name & " and sheetName:" & sheetName
            worksheetExists = True
            Exit Function
        End If
    Next ws
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
