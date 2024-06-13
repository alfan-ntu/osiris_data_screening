VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PLIDetailsForm 
   Caption         =   "Osiris 可比較公司篩選工具"
   ClientHeight    =   10092
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   12732
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
'   Date: 2024/6/13
'   Author: maoyi.fan@yapro.com.tw
'   Ver.: 0.1g
'   Revision History:
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
'       1) Generate additional worksheets, including 'Comparable_Worksheet' and 'Rejection_Worksheet'
'          Comparable_Worksheet: includes a table containing comparable company and country of the comparable company,
'                                and sorted alphabetically by the company column
'          Rejection_Worksheet: includes a table containing rejected company and the reason why the companies are
'                               rejected
'       2) Go to the first unscreened company record when opening the user form, PLIDetailsForm
'       3) Support company name search function
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
    currentRow = ActiveCell.Row
    Call ensurePLIWorksheetExists(PLI_Switch)
    '
    ' Get the company name of the current row and pass it to comparableReviewByRow
    '
    Call populateComboBoxList
    currentRow = ActiveCell.Row
    Call comparableReviewByRow(PLI_Switch, currentRow)
    ' Experimental modification, added vbModeless so that Showing UserForm and operating worksheet contents
    ' can be done at the same time
    ' NOTE: This statement is risky actually. It breaks the operation logic when users move or select other cells in
    '       worksheets or workbooks
    Me.Show vbModeless
    
End Sub
'
' Description: Ensure the Screening_Worksheet exists by copying Osiris_Review_Constant.MASTER_SHEET, 列表 (2),
'              if it doesn't and set the the first record as the selected target
' Coding Date: 2024/5/13
'
Sub ensureScreeningWorksheetExists()
    Dim worksheetIndex      As Integer
    Dim baseRange           As Range
    
    If Common_Utilities.worksheetExists(Osiris_Review_Constant.SCREENING_SHEET) Then
        Debug.Print "Screening worksheet, " & Osiris_Review_Constant.SCREENING_SHEET & " exists!"
    Else
        ' Create the Screening worksheet by copying 列表 (2) and placing it right after 列表 (2)
        Sheets(Osiris_Review_Constant.MASTER_SHEET).Copy After:=Sheets(Osiris_Review_Constant.MASTER_SHEET)
        worksheetIndex = Sheets(Osiris_Review_Constant.MASTER_SHEET).Index
        Sheets(worksheetIndex + 1).Name = Osiris_Review_Constant.SCREENING_SHEET
        Debug.Print "Screening worksheet, " & Osiris_Review_Constant.SCREENING_SHEET & " created!"
        Set baseRange = Sheets(Osiris_Review_Constant.SCREENING_SHEET).Range(Osiris_Review_Constant.CONST_BASE_RANGE)
        baseRange.Select
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
' Description: presetPLIWorksheet() presets PLI comparable column, CONST_PLI_COMPARABLE_COLUMN, to synchronize
'              screening results per Screening_Worksheet when the PLI comparable sheet is created
' Coding Date: 2024/5/21
' ToDo's:
'       1. eliminate the use of hard-coded variables, e.g. rowBase...
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

    tmpString = "!$" & Osiris_Review_Constant.CONST_COMPANY_NAME_COLUMN & "$" & CStr(rowBase) & ":$" _
                & Osiris_Review_Constant.CONST_STATUS_COLUMN & "$" & CStr(rowEnd)
    Debug.Print "Screening source range: " & tmpString
    
    If PLI_Switch = Osiris_Review_Constant.CONST_OM_PLI Then
        ' OM Review
        targetWorksheetName = Osiris_Review_Constant.OM_COMPARABLE_SHEET
    Else
        'NCP Review
        targetWorksheetName = Osiris_Review_Constant.NCP_COMPARABLE_SHEET
    End If
    Set tgtWs = Worksheets(targetWorksheetName)
    Set selectedRange = tgtWs.Range(Osiris_Review_Constant.CONST_PLI_SHEET_BASE_RANGE)
    
    lRow = Osiris_Review_Gadgets.FindMaximumRow(selectedRange)
    For r = 15 To lRow
        ' Set comparable column vlookup formula
        screeningRangeString = screeningSheet & tmpString
        Set tmpRange = tgtWs.Cells(r, Osiris_Review_Constant.CONST_PLI_COMPARABLE_COLUMN)
        colIndex = Asc(Osiris_Review_Constant.CONST_STATUS_COLUMN) - Asc(Osiris_Review_Constant.CONST_COMPANY_NAME_COLUMN) + 1
        tmpRange.Formula = "= VLOOKUP(B" & CStr(r) & ", " & screeningRangeString & ", " & CStr(colIndex) & ", FALSE)"
        
        ' Set country code column vlookup formula
        Set tmpRange = Nothing
        Set tmpRange = tgtWs.Cells(r, Osiris_Review_Constant.CONST_PLI_COUNTRY_COLUMN)
        colIndex = Asc(Osiris_Review_Constant.CONST_COUNTRY_CODE_COLUMN) - Asc(Osiris_Review_Constant.CONST_COMPANY_NAME_COLUMN) + 1
        tmpRange.Formula = "= VLOOKUP(B" & CStr(r) & ", " & screeningRangeString & ", " & CStr(colIndex) & ", FALSE)"
        countryINChinese = Osiris_Review_Gadgets.countryCodeDict(tmpRange.Value)
        tmpRange.ClearContents
        tmpRange.Value = countryINChinese
        
        ' Convert company names in all capital letter case to proper case
        Set tmpRange = Nothing
        Set tmpRange = tgtWs.Cells(r, Osiris_Review_Constant.CONST_PLI_COMPANY_PROPER_COLUMN)
        tmpRange.Formula = "= PROPER(B" & CStr(r) & ")"
        
        ' Add rejection reason vlookup formula
        Set tmpRange = Nothing
        Set tmpRange = tgtWs.Cells(r, Osiris_Review_Constant.CONST_PLI_REJECTION_REASON_COLUMN)
        colIndex = Asc(Osiris_Review_Constant.CONST_MANUAL_REVIEW_COLUMN) - Asc(Osiris_Review_Constant.CONST_COMPANY_NAME_COLUMN) + 1
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
    Dim commentText                                             As String
    Dim lRow, r                                                 As Long
    Dim screenStat                                              As Screening_Statistics
    Dim nc                                                      As Integer
    
    companyIdx = ActiveSheet.Cells(currentRow, Osiris_Review_Constant.CONST_IDX_COLUMN).Value
    companyName = ActiveSheet.Cells(currentRow, Osiris_Review_Constant.CONST_COMPANY_NAME_COLUMN).Value
    Set selectedRange = ActiveSheet.Cells(currentRow, Osiris_Review_Constant.CONST_COMPANY_NAME_COLUMN)
    
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
    Set selectedRange = tgtWs.Range(Osiris_Review_Constant.CONST_PLI_SHEET_BASE_RANGE)
    ' Locate the final row of the company list
    lRow = Osiris_Review_Gadgets.FindMaximumRow(selectedRange)
    ' Retrieve PLI numbers of the company under review
    PLI_Title = CStr(tgtWs.Cells(4, Osiris_Review_Constant.CONST_PLI_CY_COLUMN).Value)
    PLI_Title = Osiris_Review_Gadgets.CleanMessyString(PLI_Title)
    PLIMinus1_Title = tgtWs.Cells(4, Osiris_Review_Constant.CONST_PLI_LY_COLUMN).Value
    PLIMinus1_Title = CleanMessyString(PLIMinus1_Title)
    PLIMinus2_Title = tgtWs.Cells(4, Osiris_Review_Constant.CONST_PLI_LLY_COLUMN).Value
    PLIMinus2_Title = CleanMessyString(PLIMinus2_Title)
    For r = 1 To lRow
        Set tempRange = tgtWs.Cells(r, Osiris_Review_Constant.CONST_PLI_COMPANY_COLUMN)
        If companyName = tempRange.Value Then
            PLI_average = Format(tgtWs.Cells(r, Osiris_Review_Constant.CONST_PLI_AVERAGE_COLUMN).Value, "##0.00")
            PLI = Format(tgtWs.Cells(r, Osiris_Review_Constant.CONST_PLI_CY_COLUMN).Value, "##0.00")
            PLI_minus_1 = Format(tgtWs.Cells(r, Osiris_Review_Constant.CONST_PLI_LY_COLUMN).Value, "##0.00")
            PLI_minus_2 = Format(tgtWs.Cells(r, Osiris_Review_Constant.CONST_PLI_LLY_COLUMN).Value, "##0.00")
            Exit For
        End If
    Next r
    
    '
    ' Retrieve company information from Screening_Worksheet and populate data to the UserForm
    '
    primaryBusiness = ActiveSheet.Cells(currentRow, Osiris_Review_Constant.CONST_TRADE_COLUMN).Value
    businessDescription = ActiveSheet.Cells(currentRow, Osiris_Review_Constant.CONST_COMPANY_DESCRIPTION_COLUMN).Value
    productAndService = ActiveSheet.Cells(currentRow, Osiris_Review_Constant.CONST_PNS_COLUMN).Value
    commentText = ActiveSheet.Cells(currentRow, Osiris_Review_Constant.CONST_COMMENT_COLUMN).Value
    '
    ' Determine comparable state label
    '
    comparableStateLabel = ActiveSheet.Cells(currentRow, Osiris_Review_Constant.CONST_STATUS_COLUMN).Value
    comparableStateLabel = Osiris_Review_Gadgets.ReturnStateLabel(comparableStateLabel)
    rejectionReason = ActiveSheet.Cells(currentRow, Osiris_Review_Constant.CONST_MANUAL_REVIEW_COLUMN).Value
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
    activeCellRow = ActiveCell.Row
    
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
    currentRow = ActiveCell.Row
    With q
        .minQuartile = CDbl(Me.tbMin.Value)
        .lowerQuartile = CDbl(Me.tbLowerQuartile.Value)
        .medianQuartiile = CDbl(Me.tbMedian.Value)
        .upperQuartile = CDbl(Me.tbMedian.Value)
        .maxQuartile = CDbl(Me.tbMax.Value)
    End With
    
    ' update Screening_Worksheet
    ActiveSheet.Cells(currentRow, Osiris_Review_Constant.CONST_STATUS_COLUMN).Value = comparableCategory
    ActiveSheet.Cells(currentRow, Osiris_Review_Constant.CONST_MANUAL_REVIEW_COLUMN).Value = rejectConditionReason
    ActiveSheet.Cells(currentRow, Osiris_Review_Constant.CONST_COMMENT_COLUMN).Value = reviewComment
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
    Me.tbMin.Value = q.minQuartile
    Me.tbLowerQuartile.Value = q.lowerQuartile
    Me.tbMedian.Value = q.medianQuartiile
    Me.tbUpperQuartile.Value = q.upperQuartile
    Me.tbMax.Value = q.maxQuartile

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
    Set selectedRange = tgtWs.Range(Osiris_Review_Constant.CONST_PLI_SHEET_BASE_RANGE)

    '
    ' ToDo's: allocate this tempRange according to actual situation
    '
    lRow = Osiris_Review_Gadgets.FindMaximumRow(selectedRange)
    fRow = selectedRange.Row
    PLIRangeString = Osiris_Review_Constant.CONST_PLI_AVERAGE_COLUMN & CStr(fRow) & ":" & _
                     Osiris_Review_Constant.CONST_PLI_COMPARABLE_COLUMN & CStr(lRow)
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
' Description: Move to the next record for new review
' Code Date: 2024/4/15
'
Private Sub cbNext_Click()
    Dim currRow, minRow, maxRow             As Long
    Dim activeCellRow, activeCellColumn     As Long
    Dim PLISwitch                           As String
    
    minRow = Osiris_Review_Gadgets.FindMinimumRow(ActiveSheet.Range(Osiris_Review_Constant.CONST_BASE_RANGE))
    maxRow = Osiris_Review_Gadgets.FindMaximumRow(ActiveSheet.Range(Osiris_Review_Constant.CONST_BASE_RANGE))
    
    currRow = ActiveCell.Row
    Debug.Print "Current row: " & currRow
    If currRow < maxRow Then
        With ActiveCell
            .Offset(1, 0).Select
        End With
        currRow = ActiveCell.Row
        PLISwitch = Osiris_Review_Gadgets.PLILabelToSwitch(Me.lblPLI.Caption)
        Debug.Print "New Current row: " & currRow & " PLI Switch: " & PLISwitch
        
        Call comparableReviewByRow(PLISwitch, currRow)
    Else
        MsgBox "已到達最後一筆可比較公司資料", vbExclamation
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
    
    minRow = Osiris_Review_Gadgets.FindMinimumRow(ActiveSheet.Range(Osiris_Review_Constant.CONST_BASE_RANGE))
    maxRow = Osiris_Review_Gadgets.FindMaximumRow(ActiveSheet.Range(Osiris_Review_Constant.CONST_BASE_RANGE))
    
    currRow = ActiveCell.Row
    Debug.Print "Current row: " & currRow
    If currRow > minRow Then
        With ActiveCell
            .Offset(-1, 0).Select
        End With
        currRow = ActiveCell.Row
        PLISwitch = Osiris_Review_Gadgets.PLILabelToSwitch(Me.lblPLI.Caption)
        Debug.Print "New Current row: " & currRow & " PLI Switch: " & PLISwitch
        Call comparableReviewByRow(PLISwitch, currRow)
    Else
        MsgBox "已到達第一筆可比較公司資料", vbExclamation
    End If
End Sub

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
    
    currentRow = ActiveCell.Row
    companyName = ActiveSheet.Cells(currentRow, Osiris_Review_Constant.CONST_COMPANY_NAME_COLUMN).Value
    updateScreeningWorksheet = False

    Set srcWorksheet = Sheets(Osiris_Review_Constant.MASTER_SHEET)
    Debug.Print "Resetting company: " & companyName & " at current row: " & CStr(currentRow)
    originalVr = retrieveOriginalRecord(srcWorksheet, companyName)
    
    If originalVr.valid Then
        tempStr = Osiris_Review_Constant.CONST_COMPANY_NAME_COLUMN & CStr(currentRow)
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
    Set upperLeftCell = srcWorksheet.Range(Osiris_Review_Constant.CONST_BASE_RANGE)
    vr.srcRangeStr = upperLeftCell.Address
    
    ' Loop the original records row-by-row to find the associated company record
    lRow = Osiris_Review_Gadgets.FindMaximumRow(upperLeftCell)
    Debug.Print "Traverse company info. to row: " & CStr(lRow)
    
    For r = 1 To lRow
        Set tempRange = srcWorksheet.Cells(r, Osiris_Review_Constant.CONST_COMPANY_NAME_COLUMN)
        If companyName = tempRange.Value Then
            srcRangeStr = Osiris_Review_Constant.CONST_COMPANY_NAME_COLUMN & CStr(r) & ":" & Osiris_Review_Constant.CONST_STATUS_COLUMN & CStr(r)
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
    With Me.cboxComparableState
        .AddItem Osiris_Review_Constant.CONST_COMPARABLE_STATE_NG
        .AddItem Osiris_Review_Constant.CONST_COMPARABLE_STATE_OK
        .AddItem Osiris_Review_Constant.CONST_COMPARABLE_STATE_CONDITION
        .AddItem Osiris_Review_Constant.CONST_COMPARABLE_STATE_TBD
    End With
    
    With Me.cboxRejectionReason
        .AddItem Osiris_Review_Constant.RR_SIG_DIFF
        .AddItem Osiris_Review_Constant.RR_BIG_MARKETING_EXPENSE
        .AddItem Osiris_Review_Constant.RR_BIG_RD_EXPENSE
        .AddItem Osiris_Review_Constant.RR_MISSING_DATA
        .AddItem Osiris_Review_Constant.RR_THREE_YEAR_LOSS
        .AddItem Osiris_Review_Constant.RR_OTHERS
    End With
End Sub

