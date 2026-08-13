VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Screening_Worksheet_Last_Year 
   Caption         =   "去年度可比較公司篩選表"
   ClientHeight    =   2895
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   9645.001
   OleObjectBlob   =   "Screening_Worksheet_Last_Year.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Screening_Worksheet_Last_Year"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'
'   Description: A UserForm supporting retrieving data from the Screening Worksheet of last year
'
'   Date: 2025/5/2
'   Author: maoyi.fan@yapro.com.tw
'   Ver.: 0.1l
'   Revision History:
'       - 2025/5/2,  0.1l: New User Form and added column header of extra columns in worksheets "Screening_Worksheet", "PLI_Screening"

'
'   ToDo's:
'       1)
'
'
Option Explicit
'
' Public variables returned to the caller
'
Public pathToWorkbookLastYear   As String
Public companyNameColumn        As String
Public comparableStateColumn    As String
Public reviewCommentColumn      As String
'
' Description: Clear all the input text boxes
'
Private Sub cbCancel_Click()
    Me.tboxFileSelected.Value = ""
    Me.tbCompanyCol.Value = ""
    Me.tbComparableCol.Value = ""
    Me.tbReviewCol.Value = ""
End Sub

'
' Description:Capture values input to public variables
'
Private Sub cbOK_Click()
    pathToWorkbookLastYear = Me.tboxFileSelected.Value
    companyNameColumn = Me.tbCompanyCol.Value
    comparableStateColumn = Me.tbComparableCol.Value
    reviewCommentColumn = Me.tbReviewCol.Value
'    Unload Me
    Me.Hide
End Sub

'
' Description: Click to select last year's comparable screening workbook
'
' Coding Date: 2025/4/23
'
' ToDo's:
'
Private Sub cbSelect_Click()
    Dim fileSelected    As String
    Dim initPath        As String
    
    initPath = "C:\"
    fileSelected = GetLastYearWorkbook(initPath)
'    Debug.Print "File selected: " & fileSelected
    Me.tboxFileSelected.Value = fileSelected
End Sub


'
' Pop out directory select dialog and return selected folder
'
Function GetLastYearWorkbook(strPath As String) As String
    Dim fldr As FileDialog
    Dim sItem As String

    Set fldr = Application.FileDialog(msoFileDialogFilePicker)
    With fldr
        .Title = "Select an Excel File"
        .AllowMultiSelect = False
        .InitialFileName = strPath
        .Filters.Clear
        .Filters.Add "Excel files", "*.xlsx; *.xls; *.xlsm"
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With

NextCode:
    GetLastYearWorkbook = sItem
    Set fldr = Nothing
End Function
