VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PLIDetailsForm 
   Caption         =   "Osiris 可比較公司篩選工具"
   ClientHeight    =   10104
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   12816
   OleObjectBlob   =   "PLIDetailsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PLIDetailsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'   Description: A UserForm supporting Osiris result screening
'
'   Date: 2024/4/14
'   Author: maoyi.fan@yapro.com.tw
'   Ver.: 0.1b
'   Revision History:
'       -
'       - 2024/4/12, 0.1b: Add NCP support by abstracting the data search and display by PLI
'       - 2024/4/11, 0.1a: initial version
'
'   ToDo's:
'       1) Add screening results update in the UserForm
'       2) Confirm the comparable classification
'       3) Reload the original record
'       4) Store the classification to the PLI reviewed worksheet
'
'
Option Explicit

'
' Description: new version of ComparableReview main program which accepts PLI indicator and currow row
'              as the argument so that traversing comparable compnay list is easier
' Date: 2024/4/15
'
'
Sub comparableReview(PLI_Switch As String)
    Dim currentRow          As Long
    '
    ' Get the company name of the current row and pass it to comparableReviewByRow
    '
    Call populateComboBoxList
    currentRow = ActiveCell.Row
    Call comparableReviewByRow(PLI_Switch, currentRow)
    ' Experimental modification, added vbModeless so that Showing UserForm and operating worksheet contents
    ' can be done at the same time
    Me.Show vbModeless
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
    Dim lRow, r                                                 As Long
    Dim screenStat                                              As Screening_Statistics
    
    companyIdx = ActiveSheet.Cells(currentRow, Osiris_Review_Constant.CONST_IDX_COLUMN).Value
    companyName = ActiveSheet.Cells(currentRow, Osiris_Review_Constant.CONST_COMPANY_NAME_COLUMN).Value
    Debug.Print "Company name: " & companyName
    PLIString = ActiveSheet.Cells(currentRow, Osiris_Review_Constant.CONST_STATUS_COLUMN).Value
    Debug.Print "Comparable state: " & AscW(PLIString)
    
    If PLI_Switch = Osiris_Review_Constant.CONST_OM_PLI Then
        targetWorksheetName = Osiris_Review_Constant.OM_DETAILS_SHEET
        PLIString = Osiris_Review_Constant.CONST_OM_PLI_LABEL
    ElseIf PLI_Switch = Osiris_Review_Constant.CONST_NCP_PLI Then
        targetWorksheetName = Osiris_Review_Constant.NCP_DETAILS_SHEET
        PLIString = Osiris_Review_Constant.CONST_NCP_PLI_LABEL
    End If
    
    Set tgtWs = Worksheets(targetWorksheetName)
    Set selectedRange = tgtWs.Range("B15")
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
    '
    ' Determine comparable state label
    '
    comparableStateLabel = ActiveSheet.Cells(currentRow, Osiris_Review_Constant.CONST_STATUS_COLUMN).Value
    comparableStateLabel = Osiris_Review_Gadgets.ReturnStateLabel(comparableStateLabel)
    rejectionReason = ActiveSheet.Cells(currentRow, Osiris_Review_Constant.CONST_MANUAL_REVIEW_COLUMN).Value
    If AscW(rejectionReason) = Osiris_Review_Constant.UNICODE_CHECK Then
        rejectionReason = Osiris_Review_Constant.CONST_COMPARABLE_STATE_TBD
    End If
    '
    ' Determine screening statistics
    '
    screenStat = Osiris_Review_Gadgets.ScreenStatistics(ActiveSheet)
    
    Me.cboxRejectionReason.Value = rejectionReason
    Me.tbCompanyIdx.Value = companyIdx & "/" & CStr(lRow - 2)
    Me.tbCompanyName.Value = companyName
    Me.tbPrimaryBusiness.Value = primaryBusiness
    Me.tbBusinessDescription.Value = businessDescription
    Me.tbProductAndService.Value = productAndService
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
' ToDo's: Actually update the Screening_Worksheet with the reviewed classification and reason or business
'         description
'
Private Sub cbConfirm_Click()
    Dim companyUnderReview, comparableCategory, rejectionReason, comparableBusinessDescription As String
    Dim msgbox_prompt As String
    Dim lRow, r As Long
    
    Dim activeCellRow, activeCellColumn As Long
    activeCellRow = ActiveCell.Row
        
    companyUnderReview = Me.tbCompanyName.Value
    comparableCategory = Me.cboxComparableState.Value
    rejectionReason = Me.cboxRejectionReason.Value
    comparableBusinessDescription = Me.tbBusinessDescription
    msgbox_prompt = companyUnderReview & " is classified as " & comparableCategory
    If comparableCategory = Osiris_Review_Constant.CONST_COMPARABLE_STATE_NG Then
        If rejectionReason = "" Then
            msgbox_prompt = msgbox_prompt & vbNewLine & "Rejection reason cannot be blank!"
            MsgBox msgbox_prompt, vbExclamation
        Else
            msgbox_prompt = msgbox_prompt & vbNewLine & "Rejection: " & rejectionReason
            MsgBox msgbox_prompt, vbOKOnly Or vbInformation
        End If
    ElseIf comparableCategory = Osiris_Review_Constant.CONST_COMPARABLE_STATE_OK Then
        msgbox_prompt = msgbox_prompt & vbNewLine & "Business Description: " & comparableBusinessDescription
        MsgBox msgbox_prompt, vbOKOnly Or vbInformation
    End If
    
End Sub

'
' Description: The command button OK is clicked to close the UserForm
'
Private Sub cbExit_Click()
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
' Description: Reload the original Osiris record for review
'
Private Sub cbReload_Click()
    Debug.Print "Reload the record under review from the worksheet 列表 (2)"
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
        .AddItem Osiris_Review_Constant.RR_OTHERS
    End With
End Sub


