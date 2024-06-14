VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'
'   Description: Initialization of anything common to all the Excel application instances
'
'   Date: 2024/4/11
'   Author: maoyi.fan@yapro.com.tw
'   Ver.: 0.1a
'   Revision History:
'       - 2024/4/11, 0.1a: initial version
'
'   Function List:
'       1) Workbook_Open(): Initializes those keyboard shortcuts in order to register macros common to all
'                           Excel workbooks. Comment out anything not necessary
'
'   ToDo's:
'       1) Use a general macro to support the selection among different marco profiles
'
Sub Workbook_Open()
    ' Assign the shortcut key Ctrl-Shift-B to the subroutine Sub GotoFirstSheet in Common_Utilities module
    Application.OnKey "+^{B}", "GotoFirstSheet"
    Debug.Print "Assigning shortcut Ctrl-Shift-B to 'GotoFirstSheet'"
    
    ' Assign the shortcut key Ctrl-Shift-L to the subroutine Sub ListSheets in Common_Utilities module
    Application.OnKey "+^{L}", "ListSheets"
    Debug.Print "Assigning shortcut Ctrl-Shift-L to 'ListSheets'"
    
    ' Assign the shortcut key Ctrl-Shift-W to the subroutine Sub WrapCell in Common_Utilities module
    Application.OnKey "+^{W}", "WrapCell"
    Debug.Print "Assigning shortcut Ctrl-Shift-W to 'WrapCell'"
    
    ' Assign the shortcut key Ctrl-Shift-C to the subroutine Sub RedCrossout in Common_Utilities module
    ' Application.OnKey "+^{C}", "RedCrossout"
    ' Debug.Print "Assigning shortcut Ctrl-Shift-D to 'RedCrossout'"
    
    ' Assign the shortcut key Ctrl-Shift-O to the subroutine Sub CompanyOMDetails in Common_Utitities module
    Application.OnKey "+^{O}", "CompanyOMDetails"
    
    ' Assign the shortcut key Ctrl-Shift-N to the subroutine Sub CompanyNCPDetails in Common_Utilities
    Application.OnKey "+^{N}", "CompanyNCPDetails"
    
    ' Assign the shortcut key Ctrl-Shift-T to the subroutine Sub configOpParam in Osisris_Review_Constant module
    ' Newly added test code
    ' Application.OnKey "+^{T}", "Osiris_Review_Constant.configOpParam"
    
End Sub
