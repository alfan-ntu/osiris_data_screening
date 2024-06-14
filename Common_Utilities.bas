Attribute VB_Name = "Common_Utilities"
'
'   Description: A group of funtion utilities or gadgets common to all Excel workbooks, primarily
'                to support data screening of Osiris comparable company results
'
'   Date: 2024/6/15
'   Author: maoyi.fan@yapro.com.tw
'   Ver.: 0.1h
'   Revision History:
'       - 2024/6/15, 0.1h: Adjusted constant arrangement to accommodate dual operation conditions
'       - 2024/4/15, 0.1c: Moved Osiris data screening stuffs to PLIDetailsForm to make this Common_Utilities
'                          as general as possible
'       - 2024/4/12, 0.1b: Add NCP support by abstracting the data search and display by PLI
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
'
'       6) CompanyNCPDetails (Ctrl-Shift-N): This is more an Osiris data screening utilities. It looks up the
'                                           Net Cost Plus financial data of the company in the selected cell
'
'   ToDo's:
'       1) Take user's input to determine if the data screening is of single rejection or dual rejection reasons
'
Option Explicit
'
' Constants and global variables supporting worksheet lists generation
'
Public Const WORKSHEET_LIST                     As String = "Worksheet List"
Public Const WORKSHEET_INDEX_COLUMN             As String = "A"
Public Const WORKSHEET_NAME_COLUMN              As String = "B"
Public Const WORKSHEET_VISIBLE_COLUMN           As String = "C"

Public newlyCreated As Boolean

'
' Description: Display the Net Cost Plus(NCP) financial data of the selected company
' Keyboard Shortcut: (Ctrl-Shift-N)
' Code Date: 2024/4/12
'
Sub CompanyNCPDetails()
    ' Test code: 2024/6/14
    Osiris_Review_Constant.configOpParam (Osiris_Review_Constant.SINGLE_EXCLUSION_CRITERIA)
    Debug.Print "Param1: " & Osiris_Review_Constant.OP_PARAM1 & " Param2: " & Str(Osiris_Review_Constant.OP_PARAM2)
    
    PLIDetailsForm.comparableReview (CONST_NCP_PLI)
End Sub

'
' Description: Display the Operating Margin(OM) financial data of the selected company
' Keyboard Shortcut: (Ctrl-Shift-O)
' Code Date: 2024/4/11
'
Sub CompanyOMDetails()
    ' Test code: 2024/6/14
    Osiris_Review_Constant.configOpParam (Osiris_Review_Constant.DUAL_EXCLUSION_CRITERIA)
    Debug.Print "Param1: " & Osiris_Review_Constant.OP_PARAM1 & " Param2: " & Str(Osiris_Review_Constant.OP_PARAM2)
    
    PLIDetailsForm.comparableReview (CONST_OM_PLI)
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
            ' Debug.Print "ws.Name:" & ws.Name & " and sheetName:" & sheetName
            worksheetExists = True
            Exit Function
        End If
    Next ws
End Function

'
' Description: save the current workbook
'              It saves the current workbook with the current file name, saves to the specified file
'              if the optional file name is given
' Date: 2024/5/8
'
Sub saveWorkbook(Optional wbName As String)
    Dim wb As Workbook
    
    ' Get the reference to the active workbook
    Set wb = ActiveWorkbook
    
    If wbName <> "" Then
        ' Save the workbook with the specified file name, wbName
        wb.SaveAs wbName
        Debug.Print "Active workbook saved as " & wbName
    Else
        ' Save the workbook using its current file name and format
        wb.Save
        Debug.Print "Active workbook " & ThisWorkbook.Name & " updated!"
    End If
End Sub
