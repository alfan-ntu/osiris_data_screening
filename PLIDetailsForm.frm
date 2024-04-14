VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PLIDetailsForm 
   Caption         =   "可比較公司歷年利潤率"
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
'       - 2024/4/12, 0.1b: Add NCP support by abstracting the data search and display by PLI
'       - 2024/4/11, 0.1a: initial version
'
'   ToDo's:
'       1) Perform a source code re-organization to group Osiris data screening functionalities in this
'          UserForm code in order to leave Common_Utilities as simple and as genal as possible
'
'


'
' Description: List the reviewed results and comparable classification of the reviewed company
' ToDo's: Actually update the Screening_Worksheet with the reviewed classification and reason or business
'         description
'
Private Sub cbConfirm_Click()
    Dim companyUnderReview, comparableCategory, rejectionReason, comparableBusinessDescription As String
    Dim msgbox_prompt As String
        
    companyUnderReview = Me.tbCompanyName.Value
    comparableCategory = Me.cboxComparableState.Value
    rejectionReason = Me.cboxRejectionReason
    comparableBusinessDescription = Me.tbBusinessDescription
    
    msgbox_prompt = companyUnderReview & " is classified as " & comparableCategory
    If comparableCategory = Common_Utilities.CONST_COMPARABLE_STATE_NG Then
        msgbox_prompt = msgbox_prompt & vbNewLine & "Rejection: " & rejectionReason
    ElseIf comparableCategory = Common_Utilities.CONST_COMPARABLE_STATE_OK Then
        msgbox_prompt = msgbox_prompt & vbNewLine & "Business Description: " & comparableBusinessDescription
    End If
    
    MsgBox msgbox_prompt, vbOKOnly
    
End Sub

'
' Description: The command button OK is clicked to close the UserForm
'
Private Sub cbExit_Click()
    Unload Me
End Sub
'
' Description: Reload the original Osiris record for review
'
Private Sub cbReload_Click()
    Debug.Print "Reload the record under review from the worksheet 列表 (2)"
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
Private Sub UserForm_Activate()
    Debug.Print "UserForm_Activate() is called"
    With Me.cboxComparableState
        .AddItem Common_Utilities.CONST_COMPARABLE_STATE_TBD
        .AddItem Common_Utilities.CONST_COMPARABLE_STATE_NG
        .AddItem Common_Utilities.CONST_COMPARABLE_STATE_OK
        .AddItem Common_Utilities.CONST_COMPARABLE_STATE_CONDITION
    End With
    
    With Me.cboxRejectionReason
        .AddItem Common_Utilities.RR_SIG_DIFF
        .AddItem Common_Utilities.RR_BIG_MARKETING_EXPENSE
        .AddItem Common_Utilities.RR_BIG_RD_EXPENSE
        .AddItem Common_Utilities.RR_MISSING_DATA
        .AddItem Common_Utilities.RR_OTHERS
    End With
End Sub

