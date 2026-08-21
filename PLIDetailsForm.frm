VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PLIDetailsForm 
   Caption         =   "Osiris 可比較公司篩選工具 (雅博會計師事務所)"
   ClientHeight    =   10530
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   12810
   OleObjectBlob   =   "PLIDetailsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PLIDetailsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'   Description: A UserForm supporting Osiris result screening; Primary progam dealing with Osiris data screening
'
'   Date: 2026/8/21
'   Author: maoyi.fan@yapro.com.tw
'   Ver.: 0.1n
'   Revision History:
'       - 2026/8/21, 0.1n: Supported generation of '可比較公司篩選過程' worksheet
'       - 2026/8/13, 0.1m: Supported generation of '可比較公司財務資料' worksheet
'       - 2024/9/5,  0.1k: Added column header of extra columns in worksheets "Screening_Worksheet", "PLI_Screening"
'       - 2024/8/10, 0.1j: Navigation among different companies according to the selected comparable status
'       - 2024/6/15, 0.1h: Adjusted constant arrangement to accommodate dual operation conditions
'       - 2024/6/13, 0.1g: Fixed the issue jumping to the first unscreened record when all records have been screened
'       - 2024/5/14, 0.1f: Created Screening_Worksheet and populate comparable state formula, country code... in
'                          PLI Screening Worksheet
'       - 2024/4/23, 0.1e: Presets Comparable worksheet to support PLI calculation based on comparable flag
'                          vlookup formula;
'                          Update UserFrame with PLI updates
'                          Supports reload of original company records from "列表 (2)"
'       - 2024/4/18, 0.1d: add ensurePLIWorksheetExists
'       - 2024/4/17, 0.1c: Add comment text box; save edited and review status back to Screening_Worksheet
'       - 2024/4/12, 0.1b: Add NCP support by abstracting the data search and display by PLI
'       - 2024/4/11, 0.1a: initial version
'
'   ToDo's:
'       1) Generate additional worksheets, including 'Comparable_Worksheet' and 'Rejection_Worksheet',
'          after the worksheet NCP_Screening/OM_Screening
'          Comparable_Worksheet: includes a table containing comparable company and country of the comparable company,
'                                and sorted alphabetically by the company column
'          Rejection_Worksheet: includes a table containing rejected company and the reason why the companies are
'                               rejected
'       2) Support company name search function'
'
'   NOTE: Before the Screening_Worksheet is created automatically, assuming 'Screening_Worksheet' has been created and is
'         being used as the working sheet for data screening
'
'
Option Explicit


'
' Description: new version of ComparableReview main program which accepts PLI indicator and currow row
'              as the argument so that traversing comparable compnay list is easier
' Date: 2024/6/13
' ToDo's:
'       1) Sanity check if the last column of this 'Screening_Worksheet' is of column 'N', i.e. R&D expense rejection
'          and Advertisement rejection are enabled in the database query criteria
'
Sub comparableReview(PLI_Switch As String)
    Dim currentRow          As Long
    Dim unscreenedRow       As Long
    Dim currentSheetName    As String
    Dim userChoice          As VbMsgBoxResult
    '
    ' Ensure the Screening_Worksheet exists
    '
    Call ensureScreeningWorksheetExists
    '
    ' Ensure the operation screen is 'Osiris_Review_Constant.SCREENING_SHEET', i.e. Screening_Worksheet
    '
    currentSheetName = ActiveSheet.Name
    If currentSheetName <> Osiris_Review_Constant.SCREENING_SHEET Then
        userChoice = MsgBox("請確定在Screeing_Worksheet工作頁上操作!", vbOKOnly)
        End
    End If
    
    userChoice = MsgBox("從第一筆未過濾資料開始?", vbYesNo + vbQuestion, "選取過濾資料")
    If userChoice = vbYes Then
        unscreenedRow = Osiris_Review_Gadgets.findFirstUnscreenRecord()
        If unscreenedRow <> 0 Then ' Skip jumping to the first unscreened row
            Cells(unscreenedRow, 1).Select
        End If
    End If
    currentRow = ActiveCell.row
    Call ensurePLIWorksheetExists(PLI_Switch)
    '
    ' Get the company name of the current row and pass it to comparableReviewByRow
    '
    Call populateComboBoxList
    currentRow = ActiveCell.row
    Call comparableReviewByRow(PLI_Switch, currentRow)
    ' Experimental modification, added vbModeless so that Showing UserForm and operating worksheet contents
    ' can be done at the same time
    ' NOTE: This statement is risky actually. It breaks the operation logic when users move or select other cells in
    '       worksheets or workbooks
    Me.Show vbModeless
    
End Sub
'
' Description: Preset the columns additional to 列表(2)
'
'
Sub presetScreeningWorksheet(ByVal pathToWorkbookLastYear As String, ByVal companyColLastYear As String, _
                             ByVal comparableColLastYear As String, ByVal reviewCommentColLastYear As String)
    Dim tgtWs                   As Worksheet
    Dim lRow, r                 As Long
    Dim targetRange             As Range
    Dim tmpInt                  As Integer
    Dim msgResponse
    Dim directoryStr            As String
    Dim fileNameStr             As String

    Dim previousWorkbook        As Workbook
    Dim workbookLastYear        As Workbook
    Dim worksheetLastYear       As String       ' fullpath name for VLOOKUP queries, i.e. specific syntax for VLOOKUP
    Dim screeningSheetLastYear  As Range
    Dim tmpStr                  As String
    Dim vlookupString           As String
    
    ' Set targetRange of Screening_Worksheet
    Set targetRange = Sheets(Osiris_Review_Constant.SCREENING_SHEET).Range(Osiris_Review_Constant.CONST_COMPANY_NAME_COLUMN & _
                             CStr(Osiris_Review_Constant.CONST_SCREENING_FIRST_DATA_ROW))
    lRow = Osiris_Review_Gadgets.FindMaximumRow(targetRange)
'    Debug.Print "Number of last row: " & CStr(lRow)
'    Debug.Print "Path to workbook last year: " & pathToWorkbookLastYear
    
    ' Set column headers
    tmpInt = CInt(Osiris_Review_Constant.CONST_SCREENING_FIRST_DATA_ROW) - 1
    Set targetRange = Sheets(Osiris_Review_Constant.SCREENING_SHEET).Range(Osiris_Review_Constant.CONST_COMMENT_COLUMN & CStr(tmpInt))
    With targetRange
        .Value = "Comment"
        .ColumnWidth = 30
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    Set targetRange = targetRange.Offset(0, 1)
    With targetRange
        .Value = "Comparable Last Year"
        .ColumnWidth = 12
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    Set targetRange = targetRange.Offset(0, 1)
    With targetRange
        .Value = "Comment Last Year"
        .ColumnWidth = 30
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    If pathToWorkbookLastYear = "" Or _
        companyColLastYear = "" Or _
        comparableColLastYear = "" Or _
        reviewCommentColLastYear = "" Then
        msgResponse = MsgBox("去年可比較公司篩選資料尚未設定!", vbOKOnly, "去年可比較公司資料")
        Set targetRange = Nothing
        Exit Sub
    Else
'
'       Populate VLOOKUP formula if screening sheet of last year is specified
'
        directoryStr = Common_Utilities.getDirectory(pathToWorkbookLastYear)
        fileNameStr = Common_Utilities.getFileName(pathToWorkbookLastYear)
        worksheetLastYear = directoryStr & "[" & fileNameStr & "]" & Osiris_Review_Constant.SCREENING_SHEET
        
        Set previousWorkbook = ActiveWorkbook
        Set workbookLastYear = Workbooks.Open(pathToWorkbookLastYear)
        Set screeningSheetLastYear = workbookLastYear.Sheets(Osiris_Review_Constant.SCREENING_SHEET).Range( _
            Osiris_Review_Constant.CONST_COMPANY_NAME_COLUMN & CStr(Osiris_Review_Constant.CONST_SCREENING_FIRST_DATA_ROW))
        tmpInt = Osiris_Review_Gadgets.FindMaximumRow(screeningSheetLastYear)
        tmpStr = ",'" & worksheetLastYear & "'!$B$3:$Z$" & CStr(tmpInt) & ","
        ' Debug.Print "rangeLastYear: " & tmpStr & " of " & CStr(tmpInt) & " potential comparable companies"
        ' resume back to the workbook of current year
        workbookLastYear.Close
        previousWorkbook.Activate
        Set tgtWs = Worksheets(Osiris_Review_Constant.SCREENING_SHEET)
        For r = Int(Osiris_Review_Constant.CONST_SCREENING_FIRST_DATA_ROW) To lRow
            ' populate comparable state from last year worksheet
            Set targetRange = tgtWs.Cells(r, Osiris_Review_Constant.CONST_COMMENT_COLUMN).Offset(0, 1)
            tmpInt = Asc(comparableColLastYear) - Asc(companyColLastYear) + 1
            vlookupString = "=IF(ISNA(VLOOKUP(B" & CStr(r) & tmpStr & CStr(tmpInt) & ", FALSE)), " & Chr(34) & "N/A" & Chr(34) & _
                    ",VLOOKUP(B" & CStr(r) & tmpStr & CStr(tmpInt) & ", FALSE))"
            targetRange.Formula = vlookupString
'            Debug.Print "Comparable vlookup: " & vlookupString
            ' populate review comment of last year
            Set targetRange = tgtWs.Cells(r, Osiris_Review_Constant.CONST_COMMENT_COLUMN).Offset(0, 2)
            tmpInt = Asc(reviewCommentColLastYear) - Asc(companyColLastYear) + 1
            vlookupString = "=IF(ISNA(VLOOKUP(B" & CStr(r) & tmpStr & CStr(tmpInt) & ", FALSE)), " & Chr(34) & "N/A" & Chr(34) & _
                    ",VLOOKUP(B" & CStr(r) & tmpStr & CStr(tmpInt) & ", FALSE))"
            targetRange.Formula = vlookupString
'            Debug.Print "Review comment vlookup: " & vlookupString
        Next r
        Set screeningSheetLastYear = Nothing
        Set workbookLastYear = Nothing

    End If
    Set targetRange = Nothing
End Sub
'
' Description: Ensure the Screening_Worksheet exists by copying Osiris_Review_Constant.MASTER_SHEET, 列表 (2),
'              if it doesn't and set the the first record as the selected target;
'
' Coding Date: 2024/9/5
'
' ToDo: Prompt to get worksheets of last year to populate formula to retrieve comparable states of last year
'
Sub ensureScreeningWorksheetExists()
    Dim worksheetIndex              As Integer
    Dim targetRange                 As Range
    Dim tmpInt                      As Integer
    Dim lRow                        As Long
    
    Dim pathToWorkbookLastYear      As String
    Dim companyColLastYear          As String
    Dim comparableColLastYear       As String
    Dim reviewCommentColLastYear    As String
    
    If Common_Utilities.worksheetExists(Osiris_Review_Constant.SCREENING_SHEET) Then
        Debug.Print "Screening worksheet, " & Osiris_Review_Constant.SCREENING_SHEET & " exists!"
    Else
        ' Create the Screening worksheet by copying 列表 (2) and placing it right after 列表 (2)
        ' if Screening_Worksheet does not exist
        Sheets(Osiris_Review_Constant.MASTER_SHEET).Copy After:=Sheets(Osiris_Review_Constant.MASTER_SHEET)
        worksheetIndex = Sheets(Osiris_Review_Constant.MASTER_SHEET).Index
        Sheets(worksheetIndex + 1).Name = Osiris_Review_Constant.SCREENING_SHEET
        Debug.Print "Screening worksheet, " & Osiris_Review_Constant.SCREENING_SHEET & " created!"
        '
        ' Prompt user to select screening workbook of last year @ 2025/5/2
        '
        Screening_Worksheet_Last_Year.Show vbModal
        pathToWorkbookLastYear = Screening_Worksheet_Last_Year.pathToWorkbookLastYear
        companyColLastYear = Screening_Worksheet_Last_Year.companyNameColumn
        comparableColLastYear = Screening_Worksheet_Last_Year.comparableStateColumn
        reviewCommentColLastYear = Screening_Worksheet_Last_Year.reviewCommentColumn
        Unload Screening_Worksheet_Last_Year
        
        tmpInt = CInt(Osiris_Review_Constant.CONST_SCREENING_FIRST_DATA_ROW) - 1
        
        ' preset additional columns in Screening_Worksheet
        Call presetScreeningWorksheet(pathToWorkbookLastYear, companyColLastYear, _
                                 comparableColLastYear, reviewCommentColLastYear)
                                 
                                 
        ' set the highlighted range
        Call Common_Utilities.SetColumnWidth(Osiris_Review_Constant.CONST_COMMENT_COLUMN, 30)
        Set targetRange = Sheets(Osiris_Review_Constant.SCREENING_SHEET).Range(Osiris_Review_Constant.SCREENING_WORKSHEET_BASE_RANGE)
        targetRange.Select
    End If
End Sub

'
' Description: Ensure the target worksheet, OM_Comparables or NCP_Comparables, exists based on the
'              selection of PLI_Switch
' Coding Date: 2024/4/18
'
Sub ensurePLIWorksheetExists(PLI_Switch As String)
    Dim originalCell            As Range
    Dim originalSheet           As String
    Dim newWorksheetName        As String
    Dim worksheetIndex          As Integer
    
    ' Store original ActiveCell
    Set originalCell = ActiveCell
    originalSheet = originalCell.Worksheet.Name
    
    If PLI_Switch = Osiris_Review_Constant.CONST_OM_PLI Then
        If Common_Utilities.worksheetExists(Osiris_Review_Constant.OM_COMPARABLE_SHEET) Then
            Debug.Print "Target worksheet: " & Osiris_Review_Constant.OM_COMPARABLE_SHEET & " for Operating Margin review exists!"
        Else
            ' Create the missing Operation Margin review worksheet
            Sheets(Osiris_Review_Constant.OM_DETAILS_SHEET).Copy After:=Sheets(Osiris_Review_Constant.OM_DETAILS_SHEET)
            worksheetIndex = Sheets(Osiris_Review_Constant.OM_DETAILS_SHEET).Index
            Sheets(worksheetIndex + 1).Name = Osiris_Review_Constant.OM_COMPARABLE_SHEET
            'Debug.Print "Newly created worksheet name: " & Sheets(worksheetIndex + 1).Name
            presetPLIWorksheet (Osiris_Review_Constant.CONST_OM_PLI)
        End If
    ElseIf PLI_Switch = Osiris_Review_Constant.CONST_NCP_PLI Then
        If Common_Utilities.worksheetExists(Osiris_Review_Constant.NCP_COMPARABLE_SHEET) Then
            Debug.Print "Target worksheet: " & Osiris_Review_Constant.NCP_COMPARABLE_SHEET & " for Net Cost Plus review exists!"
        Else
            ' Create the missing Net Cost Plus review worksheet
            Sheets(Osiris_Review_Constant.NCP_DETAILS_SHEET).Copy After:=Sheets(Osiris_Review_Constant.NCP_DETAILS_SHEET)
            worksheetIndex = Sheets(Osiris_Review_Constant.NCP_DETAILS_SHEET).Index
            Sheets(worksheetIndex + 1).Name = Osiris_Review_Constant.NCP_COMPARABLE_SHEET
            'Debug.Print "Newly created worksheet name: " & Sheets(worksheetIndex + 1).Name
            presetPLIWorksheet (Osiris_Review_Constant.CONST_NCP_PLI)
        End If
    End If
    ' Restore original selected cell
    Sheets(originalSheet).Select
    originalCell.Select

End Sub
'
' Description: preset headers of extra columns in the PLI screening worksheet
' Coding Date: 20204/9/5
'
Sub presetPLIWorkshee_Header(tgtWs As Worksheet)
    Dim tmpString                           As String
    Dim tmpRange                            As Range
    '
    ' prepare the header of extra columns in the PLI screening worksheet
    ' Comparable State
    tmpString = Osiris_Review_Constant.CONST_PLI_COMPARABLE_COLUMN & CStr(CInt(Osiris_Review_Constant.CONST_PLI_FIRST_DATA_ROW) - 1)
    Set tmpRange = tgtWs.Range(tmpString)
    tmpRange.Value = "Comparable State"
    tmpRange.HorizontalAlignment = xlHAlignCenter
    tmpRange.WrapText = True
    Call Common_Utilities.SetColumnWidth(Osiris_Review_Constant.CONST_PLI_COMPARABLE_COLUMN, 10)
    '
    ' Country/Region Code
    tmpString = Osiris_Review_Constant.CONST_PLI_COUNTRY_COLUMN & CStr(CInt(Osiris_Review_Constant.CONST_PLI_FIRST_DATA_ROW) - 1)
    Set tmpRange = tgtWs.Range(tmpString)
    tmpRange.Value = "Country/Region"
    tmpRange.HorizontalAlignment = xlHAlignCenter
    tmpRange.WrapText = True
    Call Common_Utilities.SetColumnWidth(Osiris_Review_Constant.CONST_PLI_COMPARABLE_COLUMN, 10)
    '
    ' Company Name in Proper Form
    tmpString = Osiris_Review_Constant.CONST_PLI_COMPANY_PROPER_COLUMN & CStr(CInt(Osiris_Review_Constant.CONST_PLI_FIRST_DATA_ROW) - 1)
    Set tmpRange = tgtWs.Range(tmpString)
    tmpRange.Value = "Company Name"
    tmpRange.HorizontalAlignment = xlHAlignCenter
    tmpRange.WrapText = True
    Call Common_Utilities.SetColumnWidth(Osiris_Review_Constant.CONST_PLI_COMPANY_PROPER_COLUMN, 15)
    '
    ' Rejection Reason
    tmpString = Osiris_Review_Constant.CONST_PLI_REJECTION_REASON_COLUMN & CStr(CInt(Osiris_Review_Constant.CONST_PLI_FIRST_DATA_ROW) - 1)
    Set tmpRange = tgtWs.Range(tmpString)
    tmpRange.Value = "Rejection Reason"
    tmpRange.HorizontalAlignment = xlHAlignCenter
    tmpRange.WrapText = True
    Call Common_Utilities.SetColumnWidth(Osiris_Review_Constant.CONST_PLI_REJECTION_REASON_COLUMN, 30)
End Sub
'
' Description: presetPLIWorksheet() presets PLI comparable column, CONST_PLI_COMPARABLE_COLUMN, to synchronize
'              screening results per Screening_Worksheet when the PLI comparable sheet is created
' Coding Date: 2024/9/5
' ToDo's:
'
Sub presetPLIWorksheet(ByVal PLI_Switch As String)
    Dim targetWorksheetName                 As String
    Dim tgtWs                               As Worksheet
    Dim selectedRange, tmpRange             As Range
    Dim lRow, r                             As Long
    Dim screeningSheet                      As String
    Dim screeningRangeString, tmpString     As String
    Dim nc, rowBase, rowEnd, colIndex       As Integer
    Dim countryINChinese                    As String
    
    ' setup country code dictionary
    Call Osiris_Review_Gadgets.setupCountryCodeDictionary
    
    screeningSheet = Osiris_Review_Constant.SCREENING_SHEET
    nc = Osiris_Review_Gadgets.FindNumberOfCompanies()
    rowBase = 3
    rowEnd = rowBase + nc - 1

    tmpString = "!$" & Osiris_Review_Constant.SCREENING_WORKSHEET_COMPANY_NAME_COLUMN & "$" & CStr(rowBase) & ":$" _
                & Osiris_Review_Constant.SCREENING_WORKSHEET_STATUS_COLUMN & "$" & CStr(rowEnd)
    Debug.Print "Screening source range: " & tmpString
    
    If PLI_Switch = Osiris_Review_Constant.CONST_OM_PLI Then
        ' OM Review
        targetWorksheetName = Osiris_Review_Constant.OM_COMPARABLE_SHEET
    Else
        ' NCP Review
        targetWorksheetName = Osiris_Review_Constant.NCP_COMPARABLE_SHEET
    End If
    Set tgtWs = Worksheets(targetWorksheetName)
        
    Call presetPLIWorkshee_Header(tgtWs)
    Set selectedRange = tgtWs.Range(Osiris_Review_Constant.PLI_SHEET_BASE_RANGE)
    lRow = Osiris_Review_Gadgets.FindMaximumRow(selectedRange)
    For r = Int(Osiris_Review_Constant.CONST_PLI_FIRST_DATA_ROW) To lRow
        ' Set comparable column vlookup formula
        screeningRangeString = screeningSheet & tmpString
        Set tmpRange = tgtWs.Cells(r, Osiris_Review_Constant.PLI_SHEET_COMPARABLE_COLUMN)
        colIndex = Asc(Osiris_Review_Constant.SCREENING_WORKSHEET_STATUS_COLUMN) - Asc(Osiris_Review_Constant.SCREENING_WORKSHEET_COMPANY_NAME_COLUMN) + 1
        tmpRange.Formula = "= VLOOKUP(B" & CStr(r) & ", " & screeningRangeString & ", " & CStr(colIndex) & ", FALSE)"
        
        ' Set country code column vlookup formula
        Set tmpRange = Nothing
        Set tmpRange = tgtWs.Cells(r, Osiris_Review_Constant.PLI_SHEET_COUNTRY_COLUMN)
        colIndex = Asc(Osiris_Review_Constant.SCREENING_WORKSHEET_COUNTRY_CODE_COLUMN) - Asc(Osiris_Review_Constant.SCREENING_WORKSHEET_COMPANY_NAME_COLUMN) + 1
        tmpRange.Formula = "= VLOOKUP(B" & CStr(r) & ", " & screeningRangeString & ", " & CStr(colIndex) & ", FALSE)"
        countryINChinese = Osiris_Review_Gadgets.countryCodeDict(tmpRange.Value)
        tmpRange.ClearContents
        tmpRange.Value = countryINChinese
        
        ' Convert company names in all capital letter case to proper case
        Set tmpRange = Nothing
        Set tmpRange = tgtWs.Cells(r, Osiris_Review_Constant.PLI_SHEET_COMPANY_PROPER_COLUMN)
        tmpRange.Formula = "= PROPER(B" & CStr(r) & ")"
        
        ' Add rejection reason vlookup formula
        Set tmpRange = Nothing
        Set tmpRange = tgtWs.Cells(r, Osiris_Review_Constant.PLI_SHEET_REJECTION_REASON_COLUMN)
        colIndex = Asc(Osiris_Review_Constant.SCREENING_WORKSHEET_REVIEW_COLUMN) - Asc(Osiris_Review_Constant.SCREENING_WORKSHEET_COMPANY_NAME_COLUMN) + 1
        tmpRange.Formula = "= VLOOKUP(B" & CStr(r) & ", " & screeningRangeString & ", " & CStr(colIndex) & ", FALSE)"
        
    Next r

End Sub

'
' Description: Break down the company information display and review
' Code date: 2024/4/15
'
Sub comparableReviewByRow(ByVal PLI_Switch As String, ByVal currentRow As Long)
    Dim targetWorksheetName                                     As String
    Dim tgtWs                                                   As Worksheet
    Dim selectedRange, tempRange                                As Range
    Dim companyName, companyIdx, PLIString                      As String
    Dim primaryBusiness, businessDescription, productAndService As String
    Dim PLI_Title, PLIMinus1_Title, PLIMinus2_Title             As String
    Dim PLI_average, PLI, PLI_minus_1, PLI_minus_2              As String
    Dim comparableStateLabel, rejectionReason                   As String
    Dim comparableLastYear                                      As String
    Dim rejectReasonLastYear                                    As String
    Dim commentText                                             As String
    Dim lRow, r                                                 As Long
    Dim screenStat                                              As Screening_Statistics
    Dim nc                                                      As Integer
    
    companyIdx = ActiveSheet.Cells(currentRow, Osiris_Review_Constant.SCREENING_WORKSHEET_IDX_COLUMN).Value
    companyName = ActiveSheet.Cells(currentRow, Osiris_Review_Constant.SCREENING_WORKSHEET_COMPANY_NAME_COLUMN).Value
    Set selectedRange = ActiveSheet.Cells(currentRow, Osiris_Review_Constant.SCREENING_WORKSHEET_COMPANY_NAME_COLUMN)
    
    nc = Osiris_Review_Gadgets.FindNumberOfCompanies()
    Debug.Print "Review company name: " & companyName & " ; Total number of companies: " & CStr(nc)
    
    If PLI_Switch = Osiris_Review_Constant.CONST_OM_PLI Then
        targetWorksheetName = Osiris_Review_Constant.OM_DETAILS_SHEET
        PLIString = Osiris_Review_Constant.CONST_OM_PLI_LABEL
    ElseIf PLI_Switch = Osiris_Review_Constant.CONST_NCP_PLI Then
        targetWorksheetName = Osiris_Review_Constant.NCP_DETAILS_SHEET
        PLIString = Osiris_Review_Constant.CONST_NCP_PLI_LABEL
    End If
    
    Set tgtWs = Worksheets(targetWorksheetName)
    Set selectedRange = tgtWs.Range(Osiris_Review_Constant.PLI_SHEET_BASE_RANGE)
    ' Locate the final row of the company list
    lRow = Osiris_Review_Gadgets.FindMaximumRow(selectedRange)
    ' Retrieve PLI numbers of the company under review
    PLI_Title = CStr(tgtWs.Cells(4, Osiris_Review_Constant.PLI_SHEET_CY_COLUMN).Value)
    PLI_Title = Osiris_Review_Gadgets.CleanMessyString(PLI_Title)
    PLIMinus1_Title = tgtWs.Cells(4, Osiris_Review_Constant.PLI_SHEET_LY_COLUMN).Value
    PLIMinus1_Title = CleanMessyString(PLIMinus1_Title)
    PLIMinus2_Title = tgtWs.Cells(4, Osiris_Review_Constant.PLI_SHEET_LLY_COLUMN).Value
    PLIMinus2_Title = CleanMessyString(PLIMinus2_Title)
    For r = 1 To lRow
        Set tempRange = tgtWs.Cells(r, Osiris_Review_Constant.PLI_SHEET_COMPANY_COLUMN)
        If companyName = tempRange.Value Then
            ' update PLI indices with proper display format
            PLI_average = Format(tgtWs.Cells(r, Osiris_Review_Constant.PLI_SHEET_AVERAGE_COLUMN).Value, "##0.00")
            PLI = Format(tgtWs.Cells(r, Osiris_Review_Constant.PLI_SHEET_CY_COLUMN).Value, "##0.00")
            PLI_minus_1 = Format(tgtWs.Cells(r, Osiris_Review_Constant.PLI_SHEET_LY_COLUMN).Value, "##0.00")
            PLI_minus_2 = Format(tgtWs.Cells(r, Osiris_Review_Constant.PLI_SHEET_LLY_COLUMN).Value, "##0.00")
            Exit For
        End If
    Next r
    
    '
    ' Retrieve company information from Screening_Worksheet and populate data to the UserForm
    '
    primaryBusiness = ActiveSheet.Cells(currentRow, Osiris_Review_Constant.SCREENING_WORKSHEET_TRADE_COLUMN).Value
    businessDescription = ActiveSheet.Cells(currentRow, Osiris_Review_Constant.SCREENING_WORKSHEET_COMPANY_DESCRIPTION_COLUMN).Value
    productAndService = ActiveSheet.Cells(currentRow, Osiris_Review_Constant.SCREENING_WORKSHEET_PNS_COLUMN).Value
    commentText = ActiveSheet.Cells(currentRow, Osiris_Review_Constant.SCREENING_WORKSHEET_COMMENT_COLUMN).Value
    '
    ' Determine comparable state label
    '
    comparableStateLabel = ActiveSheet.Cells(currentRow, Osiris_Review_Constant.SCREENING_WORKSHEET_STATUS_COLUMN).Value
    comparableStateLabel = Osiris_Review_Gadgets.ReturnStateLabel(comparableStateLabel)
    rejectionReason = ActiveSheet.Cells(currentRow, Osiris_Review_Constant.SCREENING_WORKSHEET_REVIEW_COLUMN).Value
'    comparableLastYear = ActiveSheet.Cells(currentRow, Osiris_Review_Constant.CONST_COMPARABLE_LY_COLUMN).Value
    comparableLastYear = ActiveSheet.Cells(currentRow, Osiris_Review_Constant.SCREENING_WORKSHEET_COMMENT_COLUMN).Offset(0, 1).Value
    rejectReasonLastYear = ActiveSheet.Cells(currentRow, Osiris_Review_Constant.SCREENING_WORKSHEET_COMMENT_COLUMN).Offset(0, 2).Value
    ' sanity check before calling AscW(Char) function
    If rejectionReason = "" Then
        rejectionReason = " "
    End If
    If AscW(rejectionReason) = Osiris_Review_Constant.UNICODE_CHECK Then
        rejectionReason = Osiris_Review_Constant.CONST_COMPARABLE_STATE_TBD
    End If
    '
    ' Determine screening statistics
    '
    ' Debug.Print "ActiveSheet name: " & ActiveSheet.Name
    '
    screenStat = Osiris_Review_Gadgets.ScreenStatistics(ActiveSheet)
    
    Me.cboxRejectionReason.Value = rejectionReason
    Me.tbCompanyIdx.Value = companyIdx & "/" & CStr(nc)
    Me.tbCompanyName.Value = companyName
    Me.tbPrimaryBusiness.Value = primaryBusiness
    Me.tbBusinessDescription.Value = businessDescription
    Me.tbProductAndService.Value = productAndService
    Me.tbComment.Value = commentText
    Me.tbComparableLastYear.Value = comparableLastYear
    Me.txtBoxCommentLastYear.Value = rejectReasonLastYear
    Me.cboxComparableState.Value = comparableStateLabel
    Me.lblPLI.Caption = PLIString
    Me.tbPLIAverage.Value = PLI_average
    Me.tbPLI.Value = PLI
    Me.tbPLIMinus1.Value = PLI_minus_1
    Me.tbPLIMinus2.Value = PLI_minus_2
    Me.tbComparableCount.Value = screenStat.okCount
    Me.tbConditionCount.Value = screenStat.conditionCount
    Me.tbRejectCount.Value = screenStat.rejectCount
    Me.tbUnscreenCount.Value = screenStat.unscreenedCount
End Sub

'
' Description: List the reviewed results and comparable classification of the reviewed company
' Coding Date: 2024/4/18
' ToDo's: Actually update the Screening_Worksheet with the reviewed classification and reason or business
'         description
'         - Stuffs acctually updated to Screening_Worksheet include: cboxComparableState, cboxRejectionReason,
'           tbBusinessDescription and tbComment
'         - Widget contents to be updated include, tbComparableCount, tbConditionCount, tbRejectCount and
'           tbUnscreenCount
'
Private Sub cbConfirm_Click()
    Dim companyUnderReview, comparableCategory, rejectionReason, comparableBusinessDescription As String
    Dim msgbox_prompt As String
    Dim msgbox_result As Integer
    Dim lRow, r As Long
    Dim updateScreeningWorksheet As Boolean
    Dim activeCellRow, activeCellColumn As Long
    
    updateScreeningWorksheet = False
    activeCellRow = ActiveCell.row
    
    companyUnderReview = Me.tbCompanyName.Value
    comparableCategory = Me.cboxComparableState.Value
    rejectionReason = Me.cboxRejectionReason.Value
    comparableBusinessDescription = Me.tbBusinessDescription
    msgbox_prompt = companyUnderReview & " 分類為 " & comparableCategory
    If comparableCategory = Osiris_Review_Constant.CONST_COMPARABLE_STATE_NG Then
    ' 可比較公司判定: 不適用
        If rejectionReason = "" Then
            msgbox_prompt = msgbox_prompt & vbNewLine & "拒絕理由不得為空白!"
            MsgBox msgbox_prompt, vbCritical
        Else
            msgbox_prompt = msgbox_prompt & vbNewLine & "拒絕理由: " & rejectionReason
            msgbox_result = MsgBox(msgbox_prompt, vbYesNo Or vbInformation)
            If msgbox_result = vbYes Then
                updateScreeningWorksheet = True
            End If
        End If
    ElseIf comparableCategory = Osiris_Review_Constant.CONST_COMPARABLE_STATE_OK Then
    ' 可比較公司判定: 可比較公司
        msgbox_prompt = msgbox_prompt & vbNewLine & "公司描述: " & comparableBusinessDescription
        msgbox_result = MsgBox(msgbox_prompt, vbYesNo Or vbInformation)
        If msgbox_result = vbYes Then
            updateScreeningWorksheet = True
            If Me.cboxRejectionReason = Osiris_Review_Constant.CONST_COMPARABLE_STATE_TBD Then
                Me.cboxRejectionReason.Value = Osiris_Review_Constant.RR_BLANK
            End If
        End If
    ElseIf comparableCategory = Osiris_Review_Constant.CONST_COMPARABLE_STATE_CONDITION Then
    ' 可比較公司判定: 條件性、需再次判定
        If rejectionReason = "" Then
            msgbox_prompt = msgbox_prompt & vbNewLine & "條件理由不得為空白!"
            MsgBox msgbox_prompt, vbCritical
        Else
            msgbox_prompt = msgbox_prompt & vbNewLine & "條件性接受理由: " & rejectionReason
            msgbox_result = MsgBox(msgbox_prompt, vbYesNo Or vbInformation)
            If msgbox_result = vbYes Then
                updateScreeningWorksheet = True
            End If
        End If
    ElseIf comparableCategory = Osiris_Review_Constant.CONST_COMPARABLE_STATE_TBD Then
    ' 可比較公司判定: 暫時略過
        msgbox_prompt = msgbox_prompt & vbNewLine & "點擊上一筆或下一筆按鈕，繼續過濾!"
        MsgBox msgbox_prompt, vbInformation
    End If
    
    If updateScreeningWorksheet Then
        Call updateWorksheets
    End If
End Sub

'
' Description: Update Screening_Worksheet based on review results
' Coding Date: 2024/4/17
' ToDo's:
'   - Update worksheet Benchmark 1 or Benchmark 4
'
Private Sub updateWorksheets()
    Dim comparableCategory, rejectConditionReason, comparableBusinessDescription, reviewComment As String
    Dim companyName     As String
    Dim currentRow      As Long
    Dim screenStat      As Screening_Statistics
    Dim q               As Quartile_Data_Type
    
    companyName = Me.tbCompanyName.Value
    comparableCategory = Me.cboxComparableState.Value
    rejectConditionReason = Me.cboxRejectionReason.Value
    comparableBusinessDescription = Me.tbBusinessDescription.Value
    reviewComment = Me.tbComment.Value
    currentRow = ActiveCell.row
    With q
        .minQuartile = CDbl(Me.tbMin.Value)
        .lowerQuartile = CDbl(Me.tbLowerQuartile.Value)
        .medianQuartiile = CDbl(Me.tbMedian.Value)
        .upperQuartile = CDbl(Me.tbMedian.Value)
        .maxQuartile = CDbl(Me.tbMax.Value)
    End With

    ' update Screening_Worksheet
    ActiveSheet.Cells(currentRow, Osiris_Review_Constant.SCREENING_WORKSHEET_STATUS_COLUMN).Value = comparableCategory
    ActiveSheet.Cells(currentRow, Osiris_Review_Constant.SCREENING_WORKSHEET_REVIEW_COLUMN).Value = rejectConditionReason
    ActiveSheet.Cells(currentRow, Osiris_Review_Constant.SCREENING_WORKSHEET_COMMENT_COLUMN).Value = reviewComment
    ' update PLI comparable worksheet
    If Me.lblPLI = Osiris_Review_Constant.CONST_OM_PLI_LABEL Then
        Debug.Print "Get quartile information from " & Osiris_Review_Constant.OM_COMPARABLE_SHEET
        q = getQuartileUpdate(Osiris_Review_Constant.OM_COMPARABLE_SHEET)
    Else
        Debug.Print "Get quartile information from  " & Osiris_Review_Constant.NCP_COMPARABLE_SHEET
        q = getQuartileUpdate(Osiris_Review_Constant.NCP_COMPARABLE_SHEET)
    End If
    
    ' update screen statistics on UserForm PLIDetailsForm
    screenStat = Osiris_Review_Gadgets.ScreenStatistics(ActiveSheet)
    Me.tbComparableCount.Value = screenStat.okCount
    Me.tbConditionCount.Value = screenStat.conditionCount
    Me.tbRejectCount.Value = screenStat.rejectCount
    Me.tbUnscreenCount.Value = screenStat.unscreenedCount
    ' update quartile information on UserForm PLIDetailsForm
    Me.tbMin.Value = Format(q.minQuartile, "#.00")
    Me.tbLowerQuartile.Value = Format(q.lowerQuartile, "#.00")
    Me.tbMedian.Value = Format(q.medianQuartiile, "#.00")
    Me.tbUpperQuartile.Value = Format(q.upperQuartile, "#.00")
    Me.tbMax.Value = Format(q.maxQuartile, "#.00")

End Sub

'
' Description: This function updates the PLI comparable worksheet and flags comparable state to the reviewed company.
'              It also calculates quartile data in case of Osiris_Review_Constant.CONST_COMPARABLE_STATE_OK
'              and returns the quartile data to update UserForm 'PLIDetailsForm'
' Coding Date: 2024/4/19
'
Function getQuartileUpdate(ByVal comparableSheet As String) As Quartile_Data_Type
    Dim tgtWs                       As Worksheet
    Dim selectedRange, tempRange    As Range
    Dim fRow, lRow, r               As Long
    Dim q                           As Quartile_Data_Type
    Dim PLIRangeString              As String
    
    Set tgtWs = Worksheets(comparableSheet)
    Set selectedRange = tgtWs.Range(Osiris_Review_Constant.PLI_SHEET_BASE_RANGE)

    '
    ' ToDo's: allocate this tempRange according to actual situation
    '
    lRow = Osiris_Review_Gadgets.FindMaximumRow(selectedRange)
    fRow = selectedRange.row
    PLIRangeString = Osiris_Review_Constant.PLI_SHEET_AVERAGE_COLUMN & CStr(fRow) & ":" & _
                     Osiris_Review_Constant.PLI_SHEET_COMPARABLE_COLUMN & CStr(lRow)
    ' Debug.Print "<getQuartileUpdate>PLIRangeString: " & PLIRangeString
    
    Set tempRange = tgtWs.Range(PLIRangeString)
    q = Osiris_Review_Gadgets.DoComparableQuartile(tempRange, Osiris_Review_Constant.BMK_AVG_YEAR)
    getQuartileUpdate = q
    
    Set selectedRange = Nothing
    Set tgtWs = Nothing
    Set tempRange = Nothing
End Function

'
' Description: The command button OK is clicked to close the UserForm
'
Private Sub cbExit_Click()
    Call Common_Utilities.saveWorkbook
    Unload Me
End Sub

'
' Description: Move to the next record for new review. Extended 'Next' to various criteria
' ToDo's: Add a new functionality to create 可比較公司財務資料 and 可比較公司篩選過程 worksheets
' Code Date: 2024/8/10
'
Private Sub cbNext_Click()
    Dim currRow, minRow, maxRow             As Long
    Dim activeCellRow, activeCellColumn     As Long
    Dim PLISwitch                           As String
    Dim nextRow                             As Long
    Dim answer                              As VbMsgBoxResult
    Dim originalSheet                       As Worksheet
    
    minRow = Osiris_Review_Gadgets.FindMinimumRow(ActiveSheet.Range(Osiris_Review_Constant.SCREENING_WORKSHEET_BASE_RANGE))
    maxRow = Osiris_Review_Gadgets.FindMaximumRow(ActiveSheet.Range(Osiris_Review_Constant.SCREENING_WORKSHEET_BASE_RANGE))
    currRow = ActiveCell.row
    nextRow = find_next_row(ActiveCell, cboxJump.Text, minRow, maxRow)
    Debug.Print "<Debug> Next row goes to : " & nextRow
    PLISwitch = Osiris_Review_Gadgets.PLILabelToSwitch(Me.lblPLI.Caption)
    Debug.Print "New Current row: " & nextRow & " PLI Switch: " & PLISwitch
    If nextRow > maxRow Then
        answer = MsgBox( _
            "已到達最後一筆可比較公司。" & vbCrLf & _
            "是否建立「可比較公司財務資料」及「可比較公司篩選過程」工作表？", _
            vbQuestion + vbYesNo, _
            "可比較公司篩選完成")
        If answer = vbYes Then
            Set originalSheet = ActiveSheet
            CreateResultWorksheets (PLISwitch)
            originalSheet.Activate
        End If
    Else
        With ActiveCell
            .Offset(nextRow - currRow, 0).Select
        End With
        Call comparableReviewByRow(PLISwitch, nextRow)
    End If
End Sub
'
' Description: Create two results worksheets, 可比較公司財務資料 and 可比較公司篩選過程, after the review of
'              potential comparables is done
' Code Date: 2026/8/11
' Status: worked with ChatGPT
'
Private Sub CreateResultWorksheets(ByVal PLI_Switch As String)
    Dim targetWb As Workbook
    Dim anchorWs As Worksheet
    Dim ws As Worksheet
    Dim answer As VbMsgBoxResult
    Dim financialExists As Boolean
    Dim screeningExists As Boolean

    Set targetWb = ActiveWorkbook
    
    Debug.Print ActiveSheet.Name
    
    ' Determine the worksheet after which the result worksheets will be added
    Select Case PLI_Switch
        Case Osiris_Review_Constant.CONST_OM_PLI
            Set anchorWs = targetWb.Worksheets( _
                Osiris_Review_Constant.OM_COMPARABLE_SHEET)
        Case Osiris_Review_Constant.CONST_NCP_PLI
            Set anchorWs = targetWb.Worksheets( _
                Osiris_Review_Constant.NCP_COMPARABLE_SHEET)
        Case Else
            MsgBox "無法判斷 PLI 類型，無法建立結果工作表。", _
                   vbExclamation, _
                   "錯誤"
            Exit Sub
    End Select
    ' Check whether the result worksheets already exist
    financialExists = Common_Utilities.worksheetExists( _
                        Osiris_Review_Constant.FINANCIAL_DATA_SHEET)
    screeningExists = Common_Utilities.worksheetExists( _
                        Osiris_Review_Constant.SCREENING_PROCESS_SHEET)

    ' If either worksheet already exists, ask whether to replace them
    If financialExists Or screeningExists Then
        answer = MsgBox( _
            "「可比較公司財務資料」或「可比較公司篩選過程」工作表已存在。" & _
            vbCrLf & vbCrLf & _
            "是否刪除既有工作表並重新建立？", _
            vbQuestion + vbYesNo, _
            "工作表已存在")
        If answer = vbNo Then
            Exit Sub
        End If
        ' Delete existing result worksheets
        Application.DisplayAlerts = False
        If financialExists Then
            targetWb.Worksheets( _
                Osiris_Review_Constant.FINANCIAL_DATA_SHEET).Delete
        End If
        If screeningExists Then
            targetWb.Worksheets( _
                Osiris_Review_Constant.SCREENING_PROCESS_SHEET).Delete
        End If
        Application.DisplayAlerts = True
    End If
    ' Create "可比較公司財務資料" immediately after the anchor sheet
    Set ws = targetWb.Worksheets.Add(After:=anchorWs)
    ws.Name = Osiris_Review_Constant.FINANCIAL_DATA_SHEET
    ' populate contents to Osiris_Review_Constant.FINANCIAL_DATA_SHEET
    Call populateComparableFinancialData(anchorWs, ws)
    
    ' Create "可比較公司篩選過程" immediately after the financial data sheet
    Set ws = targetWb.Worksheets.Add(After:=ws)
    ws.Name = Osiris_Review_Constant.SCREENING_PROCESS_SHEET
    ' populate contents to Osiris_Review_Constant.SCREENING_PROCESS_SHEET
    Call populateScreeningProcess(targetWb.Worksheets(Osiris_Review_Constant.SCREENING_SHEET), ws)
End Sub
'
' Description: Fill screening process contents to '可比較公司篩選過程'
' Parameters:
'       sourceWs: Osiris_Review_Constant.SCREENING_SHEET, i.e. Screening_Worksheet
'       targetWs: The newly created SCREENING_PROCESS_SHEET, i.e. 可比較公司篩選過程'
' Code Date: 2026/8/20
' Status: In-progress; co-working with ChatGPT
'
Private Sub populateScreeningProcess( _
    ByVal sourceWs As Worksheet, _
    ByVal targetWs As Worksheet)

    Dim lastRow, lastCol                        As Long
    Dim dataRange, deleteRange, companyRange    As Range
    Dim sortKey, cell                           As Range
    Dim companyCol, manualReviewCol             As Long
    Dim statusCol, removedColumns               As Long

    '----------------------------------------------------------
    ' 1. Copy Screening_Worksheet to the target worksheet
    '----------------------------------------------------------
    sourceWs.Cells.Copy Destination:=targetWs.Cells
    '----------------------------------------------------------
    ' 2. Determine the current table boundaries
    '----------------------------------------------------------
    lastRow = targetWs.Cells( _
                    targetWs.Rows.Count, 1).End(xlUp).row
    lastCol = targetWs.Cells( _
                    2, targetWs.Columns.Count).End(xlToLeft).Column
    Set dataRange = targetWs.Range( _
                        targetWs.Cells(2, 1), _
                        targetWs.Cells(lastRow, lastCol))
    statusCol = targetWs.Range( _
                    Osiris_Review_Constant.CONST_STATUS_COLUMN & "1").Column
    '----------------------------------------------------------
    ' 3. Filter Status = "OK"
    '    We display OK rows temporarily, then delete them.
    '----------------------------------------------------------
    dataRange.AutoFilter _
        Field:=statusCol - dataRange.Column + 1, _
        Criteria1:="OK"
    On Error Resume Next
    Set deleteRange = dataRange.Offset(1, 0) _
                              .Resize(dataRange.Rows.Count - 1) _
                              .SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    '----------------------------------------------------------
    ' 4. Delete all Status = "OK" rows
    '----------------------------------------------------------
    If Not deleteRange Is Nothing Then
        deleteRange.EntireRow.Delete
    End If
    ' Remove AutoFilter immediately after it has served its purpose
    If targetWs.AutoFilterMode Then
        targetWs.AutoFilterMode = False
    End If
    '----------------------------------------------------------
    ' 5. Recalculate the remaining table boundaries
    '----------------------------------------------------------
    lastRow = targetWs.Cells( _
                    targetWs.Rows.Count, 1).End(xlUp).row
    lastCol = targetWs.Cells( _
                    2, targetWs.Columns.Count).End(xlToLeft).Column
    '----------------------------------------------------------
    ' 6. Convert Company Name to Proper Case
    '----------------------------------------------------------
    companyCol = targetWs.Range( _
                    Osiris_Review_Constant.CONST_COMPANY_NAME_COLUMN & _
                    "1").Column
    Set companyRange = targetWs.Range( _
                            targetWs.Cells(3, companyCol), _
                            targetWs.Cells(lastRow, companyCol))
    For Each cell In companyRange
        If Not IsError(cell.Value) Then
            If Len(Trim$(CStr(cell.Value))) > 0 Then
                cell.Value = StrConv( _
                                CStr(cell.Value), _
                                vbProperCase)
            End If
        End If
    Next cell
    '----------------------------------------------------------
    ' 7. Sort remaining NG records by Company Name
    '----------------------------------------------------------
    Set dataRange = targetWs.Range( _
                        targetWs.Cells(2, 1), _
                        targetWs.Cells(lastRow, lastCol))
    Set sortKey = targetWs.Range( _
                    targetWs.Cells(3, companyCol), _
                    targetWs.Cells(lastRow, companyCol))
    With targetWs.Sort
        .SortFields.Clear
        .SortFields.Add _
            key:=sortKey, _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal
        .SetRange dataRange
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    '----------------------------------------------------------
    ' 8. Remove columns between Company Name and Manual Review
    '----------------------------------------------------------
    companyCol = targetWs.Range( _
                    Osiris_Review_Constant.CONST_COMPANY_NAME_COLUMN & _
                    "1").Column
    manualReviewCol = targetWs.Range( _
                        Osiris_Review_Constant.CONST_MANUAL_REVIEW_COLUMN & _
                        "1").Column
    ' hardcode the column widths, alignment of Company Name and Rejection Reason
    targetWs.Columns(companyCol).ColumnWidth = 40
    targetWs.Columns(companyCol).VerticalAlignment = xlCenter
    targetWs.Columns(manualReviewCol).ColumnWidth = 60
    targetWs.Columns(manualReviewCol + 1).VerticalAlignment = xlCenter  ' Status column
    targetWs.Columns(manualReviewCol + 2).ColumnWidth = 60              ' review comment column
    removedColumns = manualReviewCol - companyCol - 1
    If removedColumns > 0 Then
        targetWs.Range( _
            targetWs.Columns(companyCol + 1), _
            targetWs.Columns(manualReviewCol - 1) _
        ).Delete
    End If
    targetWs.Rows("1:" & lastRow).AutoFit
    '----------------------------------------------------------
    ' 9. Re-index the table
    '    Row 3 -> "1."
    '    Row 4 -> "2."
    '    ...
    '----------------------------------------------------------
    Set companyRange = targetWs.Range( _
                            targetWs.Cells(3, _
                                targetWs.Range( _
                                    Osiris_Review_Constant.CONST_IDX_COLUMN & _
                                    "1").Column), _
                            targetWs.Cells(lastRow, _
                                targetWs.Range( _
                                    Osiris_Review_Constant.CONST_IDX_COLUMN & _
                                    "1").Column))
    companyRange.NumberFormat = "@"
    For Each cell In companyRange
        cell.Value = CStr(cell.row - 2) & "."
        cell.HorizontalAlignment = xlCenter
        cell.VerticalAlignment = xlCenter
    Next cell
    
    Call PopulateManualReviewSummary(targetWs, manualReviewCol - removedColumns, lastRow)
    
End Sub
'
' Description: Compose rejection reason statistics table
' Code Date: 2026/8/21
' Note: in-progress; coworking with ChatGPT
'
Private Sub PopulateManualReviewSummary( _
    ByVal targetWs As Worksheet, _
    ByVal manualReviewCol As Long, _
    ByVal lastRow As Long)

    Dim dict                As Object
    Dim cell                As Range
    Dim key                 As Variant
    Dim summaryRow, row     As Long
    Dim row                 As Long

    Set dict = CreateObject("Scripting.Dictionary")
    '----------------------------------------------------------
    ' Collect the different Manual Review categories
    '----------------------------------------------------------
    For Each cell In targetWs.Range( _
                        targetWs.Cells(3, manualReviewCol), _
                        targetWs.Cells(lastRow, manualReviewCol))

        If Len(Trim$(CStr(cell.Value))) > 0 Then
            If dict.Exists(CStr(cell.Value)) Then
                ' increment the counter in case of duplicacy
                dict(CStr(cell.Value)) = dict(CStr(cell.Value)) + 1
            Else
                ' add new element to the dictionay if it is a new one
                dict.Add CStr(cell.Value), 1
            End If
        End If
    Next cell
    '----------------------------------------------------------
    ' Summary table starts two rows after the last data row
    '----------------------------------------------------------
    summaryRow = lastRow + 3
    ' Header of the summary table
    targetWs.Cells(summaryRow, manualReviewCol).Value = "Manual Review Summary"
    targetWs.Cells(summaryRow, manualReviewCol).HorizontalAlignment = xlCenter
    ' Listing categories and counts
    row = summaryRow + 1
    For Each key In dict.Keys
        targetWs.Cells(row, manualReviewCol).Value = key
        targetWs.Cells(row, manualReviewCol + 1).Value = dict(key)
        row = row + 1
     Next key

    '----------------------------------------------------------
    ' Format the summary
    '----------------------------------------------------------
    With targetWs.Range( _
            targetWs.Cells(summaryRow + 1, manualReviewCol), _
            targetWs.Cells(row, manualReviewCol))
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    With targetWs.Range( _
            targetWs.Cells(summaryRow + 1, manualReviewCol + 1), _
            targetWs.Cells(row, manualReviewCol + 1))
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    ' Make the count row numeric
    targetWs.Range( _
        targetWs.Cells(summaryRow + 1, manualReviewCol + 1), _
        targetWs.Cells(row - 1, manualReviewCol + 1)).NumberFormat = "0"

    ' AutoFit the summary rows
    targetWs.Rows(summaryRow & ":" & summaryRow + 1).AutoFit

End Sub


'
' Description: Fill financial data contents to '可比較公司財務資料' worksheet
' Code Date: 2026/8/12
' Note: co-worked with ChatGPT
'
Private Sub populateComparableFinancialData( _
    ByVal anchorWs As Worksheet, _
    ByVal targetWs As Worksheet)
    Dim lastRow, lastCol            As Long
    Dim dataRange, deleteRange      As Range
    Dim idxRange, cell, sortKey     As Range
    Dim lyCol                       As Long
    Dim llyCol                      As Long
    Dim companyProperCol            As Long
    Dim removedColumns              As Long
    '----------------------------------------------------------
    ' 1. Copy the content of the anchor worksheet
    '----------------------------------------------------------
    anchorWs.Cells.Copy Destination:=targetWs.Cells
    '----------------------------------------------------------
    ' 2. Copy PLI Ave-CY-LY-LLY header information
    '    from row 4 to row 14
    '----------------------------------------------------------
    targetWs.Range( _
        Osiris_Review_Constant.CONST_PLI_AVERAGE_COLUMN & "4:" & _
        Osiris_Review_Constant.CONST_PLI_LLY_COLUMN & "4" _
    ).Copy Destination:= _
        targetWs.Range( _
            Osiris_Review_Constant.CONST_PLI_AVERAGE_COLUMN & "14")
    '----------------------------------------------------------
    ' 3. Remove rows 4 ~ 13 where quartile information of all
    '    companies lists
    '----------------------------------------------------------
    targetWs.Rows("4:13").Delete
    '----------------------------------------------------------
    ' 4. Remove the Rejection Reason column
    '    IMPORTANT:
    '    Do this BEFORE inserting the four new PLI columns.
    '----------------------------------------------------------
    targetWs.Columns( _
        Osiris_Review_Constant.CONST_PLI_REJECTION_REASON_COLUMN).Delete
    '----------------------------------------------------------
    ' 5. Add four PLI percentage columns after Company Name
    '----------------------------------------------------------
    Call AddPLIResultColumns(targetWs)
    '----------------------------------------------------------
    ' 6. Determine the data range
    '----------------------------------------------------------
    lastRow = targetWs.Cells( _
                    targetWs.Rows.Count, 1).End(xlUp).row
    lastCol = targetWs.Cells( _
                    4, targetWs.Columns.Count).End(xlToLeft).Column
    Set dataRange = targetWs.Range( _
                        targetWs.Cells(4, 1), _
                        targetWs.Cells(lastRow, lastCol))
    '----------------------------------------------------------
    ' 7. Filter PLI Comparable column
    '    Keep only records with "OK"
    '----------------------------------------------------------
    dataRange.AutoFilter _
        Field:=targetWs.Range( _
            Osiris_Review_Constant.CONST_PLI_COMPARABLE_COLUMN & "4").Column _
            - dataRange.Column + 1, _
        Criteria1:="<>OK"
    '----------------------------------------------------------
    ' 8. Get visible non-OK data rows
    '----------------------------------------------------------
    On Error Resume Next
    Set deleteRange = dataRange.Offset(1, 0) _
                              .Resize(dataRange.Rows.Count - 1) _
                              .SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    '----------------------------------------------------------
    ' 9. Delete all non-OK records and remove AutoFilter
    '----------------------------------------------------------
    If Not deleteRange Is Nothing Then
        deleteRange.EntireRow.Delete
    End If
    If targetWs.AutoFilterMode Then
        targetWs.AutoFilterMode = False
    End If
    '----------------------------------------------------------
    ' 10. Recalculate the remaining data range
    '----------------------------------------------------------
    lastRow = targetWs.Cells( _
                    targetWs.Rows.Count, 1).End(xlUp).row
    lastCol = targetWs.Cells( _
                    4, targetWs.Columns.Count).End(xlToLeft).Column
    Set dataRange = targetWs.Range( _
                        targetWs.Cells(4, 1), _
                        targetWs.Cells(lastRow, lastCol))
    '----------------------------------------------------------
    ' 11. Remove columns between LY and LLY.
    '     WATCH: These redundant columns might be different from
    '            report to report
    '----------------------------------------------------------
    lyCol = targetWs.Range( _
                Osiris_Review_Constant.CONST_PLI_LY_COLUMN & "1").Column
    llyCol = targetWs.Range( _
                 Osiris_Review_Constant.CONST_PLI_LLY_COLUMN & "1").Column
    companyProperCol = targetWs.Range( _
                        Osiris_Review_Constant.CONST_PLI_COMPANY_PROPER_COLUMN & _
                        "1").Column
    removedColumns = llyCol - lyCol - 1
    If removedColumns > 0 Then
        targetWs.Range( _
            targetWs.Columns(lyCol + 1), _
            targetWs.Columns(llyCol - 1) _
        ).Delete
    End If
    '----------------------------------------------------------
    ' 12. Sort by Company Proper Name
    '----------------------------------------------------------
    companyProperCol = companyProperCol - removedColumns
    Set sortKey = targetWs.Range( _
                    targetWs.Cells(4, companyProperCol), _
                    targetWs.Cells(lastRow, companyProperCol))
    With targetWs.Sort
        .SortFields.Clear
        .SortFields.Add _
            key:=sortKey, _
            SortOn:=xlSortOnValues, _
            Order:=xlAscending, _
            DataOption:=xlSortNormal
        .SetRange dataRange
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ' re-index the table
    Set idxRange = targetWs.Range( _
                    targetWs.Cells(5, 1), _
                    targetWs.Cells(lastRow, 1))
    idxRange.NumberFormat = "@"
    For Each cell In idxRange
        cell.Value = CStr((cell.row - 4)) & "."
        cell.HorizontalAlignment = xlCenter
    Next cell
    '----------------------------------------------------------
    ' 13. Build quartile table
    '----------------------------------------------------------
    Call PopulateComparableQuartile( _
        targetWs, _
        companyProperCol, _
        5, _
        lastRow)
End Sub
'
' Description: Add PLI columns in precentaged format to do quartile calculation
' Code Date: 2026/8/12
' Note: Co-worked with ChatGPT
'
Private Sub AddPLIResultColumns(ByVal ws As Worksheet)
    Dim companyCol As Long
    Dim lastRow As Long
    Dim sourceColumns As Variant
    Dim i As Long
    Dim sourceCol As Long
    Dim targetCol As Long
    Dim r As Long
    Dim sourceValue As Variant

    companyCol = ws.Range( _
                    Osiris_Review_Constant.CONST_PLI_COMPANY_PROPER_COLUMN & _
                    "1").Column
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    ' Columns from which the four new percentage columns are derived
    sourceColumns = Array( _
        Osiris_Review_Constant.CONST_PLI_CY_COLUMN, _
        Osiris_Review_Constant.CONST_PLI_LY_COLUMN, _
        Osiris_Review_Constant.CONST_PLI_LLY_COLUMN, _
        Osiris_Review_Constant.CONST_PLI_AVERAGE_COLUMN _
    )
    ' Insert four columns immediately after Company Name
    ws.Columns(companyCol + 1).Resize(, 4).Insert Shift:=xlToRight
    ' Populate the four new columns
    For i = LBound(sourceColumns) To UBound(sourceColumns)
        sourceCol = ws.Range(sourceColumns(i) & "1").Column
        targetCol = companyCol + 1 + i
        ' Copy the header
        ws.Cells(4, targetCol).Value = ws.Cells(4, sourceCol).Value
        ' Copy values / 100
        For r = 5 To lastRow
            sourceValue = ws.Cells(r, sourceCol).Value
            If IsError(sourceValue) Then
                ws.Cells(r, targetCol).Value = sourceValue
            ElseIf IsNumeric(sourceValue) Then
                ws.Cells(r, targetCol).Value = CDbl(sourceValue) / 100
            Else
                ws.Cells(r, targetCol).Value = sourceValue
            End If
        Next r
        ' Percentage format
        ws.Range( _
            ws.Cells(5, targetCol), _
            ws.Cells(lastRow, targetCol) _
        ).NumberFormat = "0.00%"
    Next i
End Sub
'
' Description: Calculate quartile per year
' Code Date: 2026/8/12
' Note: Co-worked with ChatGPT
'
Private Sub PopulateComparableQuartile( _
    ByVal targetWs As Worksheet, _
    ByVal companyProperCol As Long, _
    ByVal firstDataRow As Long, _
    ByVal lastDataRow As Long)

    Dim i, resultRow            As Long
    Dim pliRange                As Range
    Dim q                       As Quartile_Data_Type
    '----------------------------------------------------------
    ' Quartile result starts two rows below the last company
    '----------------------------------------------------------
    resultRow = lastDataRow + 2
    '----------------------------------------------------------
    ' Title / first column
    '----------------------------------------------------------
    targetWs.Cells(resultRow, companyProperCol).Value = "PLI Quartile"
    '----------------------------------------------------------
    ' Four PLI columns are immediately after Company Name
    '----------------------------------------------------------
    For i = 1 To 4
        Set pliRange = targetWs.Range( _
            targetWs.Cells(firstDataRow, companyProperCol + i), _
            targetWs.Cells(lastDataRow, companyProperCol + i))
        q = Osiris_Review_Gadgets.DoFinancialDataQuartile(pliRange)
        ' Header
        targetWs.Cells(resultRow, companyProperCol + i).Value = _
            targetWs.Cells(4, companyProperCol + i).Value
        targetWs.Cells(resultRow, companyProperCol + i).HorizontalAlignment = xlCenter
        ' Quartile results
        targetWs.Cells(resultRow + 1, companyProperCol + i).Value = _
            q.minQuartile
        targetWs.Cells(resultRow + 2, companyProperCol + i).Value = _
            q.lowerQuartile
        targetWs.Cells(resultRow + 3, companyProperCol + i).Value = _
            q.medianQuartiile
        targetWs.Cells(resultRow + 4, companyProperCol + i).Value = _
            q.upperQuartile
        targetWs.Cells(resultRow + 5, companyProperCol + i).Value = _
            q.maxQuartile
        ' Percentage format
        targetWs.Range( _
            targetWs.Cells(resultRow + 1, companyProperCol + i), _
            targetWs.Cells(resultRow + 5, companyProperCol + i) _
        ).NumberFormat = "0.00%"
    Next i
    '----------------------------------------------------------
    ' Quartile labels
    '----------------------------------------------------------
    targetWs.Cells(resultRow + 1, companyProperCol).Value = "Minimum"
    targetWs.Cells(resultRow + 2, companyProperCol).Value = "Lower Quartile"
    targetWs.Cells(resultRow + 3, companyProperCol).Value = "Median"
    targetWs.Cells(resultRow + 4, companyProperCol).Value = "Upper Quartile"
    targetWs.Cells(resultRow + 5, companyProperCol).Value = "Maximum"
End Sub
'
' Description: Locate next row according to jump condition selected
' Code Date: 2024/8/9
'
Private Function find_next_row(currRange As Range, jump_criteria As String, minRow As Variant, maxRow As Variant) As Long
    Dim nextRow, r      As Long
    Dim comparableState As String
    Dim ws              As Worksheet
        
    Set ws = ActiveSheet
    nextRow = currRange.row
    ' Debug.Print "<Debug> Jump start row number: " & currRange.Row & ";jumpType: " & jump_criteria & "; minRow: " & CStr(minRow) & "; maxRow: " & CStr(maxRow)
    Select Case jump_criteria
    Case Osiris_Review_Constant.CONST_COMPARABLE_STATE_NEXT
        nextRow = currRange.row + 1
    Case Else
        For r = (currRange.row + 1) To maxRow
            comparableState = ws.Cells(r, Osiris_Review_Constant.SCREENING_WORKSHEET_STATUS_COLUMN).Value
            If comparableState = jump_criteria Then
                nextRow = r
                GoTo ReturnLine
            End If
        Next r
        nextRow = r
    End Select
ReturnLine:
    find_next_row = nextRow
End Function
'
' Description: Automatically update tbComment for specific Reject Resaons,
'   - RR_BIG_RD_EXPENSE => 'Huge R&D expenses'
'   - RR_BIG_MARKETING_EXPENSE => 'Huge advertising expenses'
'   - RR_MISSING_DATA   => 'Missing financial data'
'   - RR_THREE_YEAR_LOSS => '3-consecutive-years loss'
' Date: 2026/7/6
'
Private Sub cboxRejectionReason_Change()
    Dim appendText As String
    Select Case Me.cboxRejectionReason.Value
        Case Osiris_Review_Constant.RR_BIG_RD_EXPENSE
            appendText = "Huge R&D expenses"

        Case Osiris_Review_Constant.RR_BIG_MARKETING_EXPENSE
            appendText = "Huge advertising expenses"

        Case Osiris_Review_Constant.RR_MISSING_DATA
            appendText = "Missing financial data"

        Case Osiris_Review_Constant.RR_THREE_YEAR_LOSS
            appendText = "3-consecutive-year loss"

        Case Else
            Exit Sub
    End Select
    ' Append the text
    If Len(Trim(Me.tbComment.Text)) > 0 Then
        Me.tbComment.Text = Me.tbComment.Text & vbCrLf & appendText
    Else
        Me.tbComment.Text = appendText
    End If
End Sub

'
' Description: Move to the previous record for new review
' Code Date: 2024/4/15
'
Private Sub cbPrev_Click()
    Dim currRow, minRow, maxRow             As Long
    Dim activeCellRow, activeCellColumn     As Long
    Dim PLISwitch                           As String
    Dim prevRow                             As Long
    
    minRow = Osiris_Review_Gadgets.FindMinimumRow(ActiveSheet.Range(Osiris_Review_Constant.SCREENING_WORKSHEET_BASE_RANGE))
    maxRow = Osiris_Review_Gadgets.FindMaximumRow(ActiveSheet.Range(Osiris_Review_Constant.SCREENING_WORKSHEET_BASE_RANGE))
    
    currRow = ActiveCell.row
    Debug.Print "Current row: " & currRow
    prevRow = find_previous_row(ActiveCell, cboxJump.Text, minRow, maxRow)
    Debug.Print "<Debug> Previous row goes to : " & prevRow
    
    If prevRow < minRow Then
        MsgBox "已到達第一筆可比較公司資料", vbExclamation
    Else
        With ActiveCell
            .Offset(prevRow - currRow, 0).Select
        End With
        PLISwitch = Osiris_Review_Gadgets.PLILabelToSwitch(Me.lblPLI.Caption)
        Debug.Print "New Current row: " & prevRow & " PLI Switch: " & PLISwitch
        Call comparableReviewByRow(PLISwitch, prevRow)
    End If

End Sub

'
' Description: Locate previous row according to jump condition selected
' Code Date: 2024/8/9
'
Private Function find_previous_row(currRange As Range, jump_criteria As String, minRow As Variant, maxRow As Variant) As Long
    Dim prevRow, r As Long
    Dim comparableState As String
    Dim ws As Worksheet
        
    Set ws = ActiveSheet
    prevRow = currRange.row
    Debug.Print "<Debug> Jump start row number: " & currRange.row & ";jumpType: " & jump_criteria & "; minRow: " & CStr(minRow) & "; maxRow: " & CStr(maxRow)
    Select Case jump_criteria
    Case Osiris_Review_Constant.CONST_COMPARABLE_STATE_NEXT
        prevRow = currRange.row - 1
    Case Else
        For r = (currRange.row - 1) To minRow Step -1
            comparableState = ws.Cells(r, Osiris_Review_Constant.SCREENING_WORKSHEET_STATUS_COLUMN).Value
            If comparableState = jump_criteria Then
                prevRow = r
                GoTo ReturnLine
            End If
        Next r
        prevRow = r
    End Select
ReturnLine:
    find_previous_row = prevRow
End Function

'
' Description: Reload the original Osiris record of the Active row for restart a new review
' Coding Date: 2024/4/23
'
Private Sub cbReload_Click()
    Dim currentRow                          As Long
    Dim companyName, tempStr                As String
    Dim originalVr                          As Validated_Range
    Dim originalRecord                      As Range
    Dim srcWorksheet                        As Worksheet
    Dim updateScreeningWorksheet            As Boolean
    Dim PLISwitch                           As String
    
    currentRow = ActiveCell.row
    companyName = ActiveSheet.Cells(currentRow, Osiris_Review_Constant.SCREENING_WORKSHEET_COMPANY_NAME_COLUMN).Value
    updateScreeningWorksheet = False

    Set srcWorksheet = Sheets(Osiris_Review_Constant.MASTER_SHEET)
    Debug.Print "Resetting company: " & companyName & " at current row: " & CStr(currentRow)
    originalVr = retrieveOriginalRecord(srcWorksheet, companyName)
    
    If originalVr.valid Then
        tempStr = Osiris_Review_Constant.SCREENING_WORKSHEET_COMPANY_NAME_COLUMN & CStr(currentRow)
        Debug.Print "Screening record at: " & tempStr
        Debug.Print "Original info. range: " & originalVr.srcRangeStr
        '
        ' execute copy operation to reset the screening content
        '
        srcWorksheet.Range(originalVr.srcRangeStr).Copy _
              Destination:=ActiveSheet.Range(tempStr)
        updateScreeningWorksheet = True
    Else
        Debug.Print "Original record not found..."
    End If

    If updateScreeningWorksheet Then
        PLISwitch = Osiris_Review_Gadgets.PLILabelToSwitch(Me.lblPLI.Caption)
        Call comparableReviewByRow(PLISwitch, currentRow)
    End If

End Sub


'
' Description: Find the original Osiris record based on the input Company name and return to row range
'              correspondent to the company
' Coding Date: 2024/4/23
'
Function retrieveOriginalRecord(ByVal srcWorksheet As Worksheet, ByVal companyName As String) As Validated_Range
    Dim valid                               As Boolean
    Dim upperLeftCell, tempRange            As Range
    Dim vr                                  As Validated_Range
    Dim lRow, r                             As Long
    Dim srcRangeStr                         As String
    
    valid = False
    
    ' Visit the master worksheet and get the original record associated with the input companyName
    Set upperLeftCell = srcWorksheet.Range(Osiris_Review_Constant.SCREENING_WORKSHEET_BASE_RANGE)
    vr.srcRangeStr = upperLeftCell.Address
    
    ' Loop the original records row-by-row to find the associated company record
    lRow = Osiris_Review_Gadgets.FindMaximumRow(upperLeftCell)
    Debug.Print "Traverse company info. to row: " & CStr(lRow)
    
    For r = 1 To lRow
        Set tempRange = srcWorksheet.Cells(r, Osiris_Review_Constant.SCREENING_WORKSHEET_COMPANY_NAME_COLUMN)
        If companyName = tempRange.Value Then
            srcRangeStr = Osiris_Review_Constant.SCREENING_WORKSHEET_COMPANY_NAME_COLUMN & CStr(r) & ":" & Osiris_Review_Constant.SCREENING_WORKSHEET_STATUS_COLUMN & CStr(r)
            Debug.Print "Source company found: " & companyName & " Original company infomation range: " & srcRangeStr
            vr.srcRangeStr = srcRangeStr
            valid = True
            Exit For
        End If
    Next r
    
    vr.valid = valid
    retrieveOriginalRecord = vr
    
    Set upperLeftCell = Nothing
    Set tempRange = Nothing
End Function


Private Sub Label9_Click()

End Sub

'
' Description: close the UserForm when ESC key is pressed when the focus is on Business Description
'              textbox
'
Private Sub tbBusinessDescription_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
       Unload Me
    End If
End Sub

'
' Description: close the UserForm when ESC key is pressed
'
Private Sub tbComment_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
       Unload Me
    End If
End Sub

'
' Description: close the UserForm when ESC key is pressed when the focus is on CompanyIdx label
'
Private Sub tbCompanyIdx_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
       Unload Me
    End If
End Sub

'
' Description: close the UserForm when ESC key is pressed. Note: textbox CompanyName is the default
'              focus of this UserForm.
'
Private Sub tbCompanyName_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 27 Then
       Unload Me
    End If
End Sub

'
' Description: Rejection reason can be anything not listed on Enter
'
Private Sub cboxRejectionReason_Enter()
    Me.cboxRejectionReason.Text = Me.cboxRejectionReason.Value
    Debug.Print "(Enter) Rejection reason is " & Me.cboxRejectionReason.Text
End Sub


'
' Description: Populates allowable items for list boxes
'
Private Sub populateComboBoxList()
    ' comparable state of each potentially comparable company
    With Me.cboxComparableState
        .AddItem Osiris_Review_Constant.CONST_COMPARABLE_STATE_NG
        .AddItem Osiris_Review_Constant.CONST_COMPARABLE_STATE_OK
        .AddItem Osiris_Review_Constant.CONST_COMPARABLE_STATE_CONDITION
        .AddItem Osiris_Review_Constant.CONST_COMPARABLE_STATE_TBD
    End With
    ' possible rejection reason list
    With Me.cboxRejectionReason
        .AddItem Osiris_Review_Constant.RR_SIG_DIFF
        .AddItem Osiris_Review_Constant.RR_BIG_MARKETING_EXPENSE
        .AddItem Osiris_Review_Constant.RR_BIG_RD_EXPENSE
        .AddItem Osiris_Review_Constant.RR_MISSING_DATA
        .AddItem Osiris_Review_Constant.RR_THREE_YEAR_LOSS
        .AddItem Osiris_Review_Constant.RR_OTHERS
    End With
    ' jump condtion for 'Next' and 'Previous' button
    With Me.cboxJump
        .AddItem Osiris_Review_Constant.CONST_COMPARABLE_STATE_NEXT
        .AddItem Osiris_Review_Constant.CONST_COMPARABLE_STATE_NG
        .AddItem Osiris_Review_Constant.CONST_COMPARABLE_STATE_OK
        .AddItem Osiris_Review_Constant.CONST_COMPARABLE_STATE_CONDITION
        .AddItem Osiris_Review_Constant.CONST_COMPARABLE_STATE_TBD
    End With
End Sub

