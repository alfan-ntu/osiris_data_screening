Attribute VB_Name = "Osiris_Review_Constant"
'
'   Description: A module listing  Osiris data review associated constants
'
'   Date: 2024/4/23
'   Author: maoyi.fan@yapro.com.tw
'   Ver.: 0.1e
'   Revision History:
'       - 2024/4/23, 0.1e: Support reload original company records
'       - 2024/4/15, 0.1b: First added
'
'   ToDo's:
'       1)
'

'
' Constant definitions associated with data screening of Osiris search/rejection results
'
Public Const CONST_OM_PLI                       As String = "Operating Margin"
Public Const CONST_NCP_PLI                      As String = "Net Cost Plus"
Public Const OM_DETAILS_SHEET                   As String = "Benchmark 1"
Public Const OM_COMPARABLE_SHEET                As String = "OM_Screening"
Public Const NCP_DETAILS_SHEET                  As String = "Benchmark 4"
Public Const NCP_COMPARABLE_SHEET               As String = "NCP_Screening"
Public Const CONST_OM_PLI_LABEL                 As String = "營業淨利率"
Public Const CONST_NCP_PLI_LABEL                As String = "成本及營業費用淨利率"
' SCREENINT_SHEET is actually a worksheet replicating
Public Const SCREENING_SHEET                    As String = "Screening_Worksheet"
'
' Layout constants associated with 列表 (2) or Screening_Worksheet
'
Public Const MASTER_SHEET                       As String = "列表 (2)"
Public Const CONST_BASE_RANGE                   As String = "B3"
Public Const CONST_IDX_COLUMN                   As String = "A"
Public Const CONST_COMPANY_NAME_COLUMN          As String = "B"
Public Const CONST_TRADE_COLUMN                 As String = "C"
Public Const CONST_COMPANY_DESCRIPTION_COLUMN   As String = "D"
Public Const CONST_PNS_COLUMN                   As String = "E"
Public Const CONST_COUNTRY_CODE_COLUMN          As String = "F"
Public Const CONST_MANUAL_REVIEW_COLUMN         As String = "M"
Public Const CONST_STATUS_COLUMN                As String = "N"
Public Const CONST_COMMENT_COLUMN               As String = "O"

' STATUS_COLUMN_OFFSET stores the offset from CONST_COMPANY_NAME_COLUMN to the
Public Const STATUS_COLUMN_OFFSET               As Integer = 13
'
' Layout constants associated with PLI Benchmark worksheet
' PLI Benchmark worksheet means OM_COMPARABLE_SHEET in case of Operating Margin review,
'                               NCP_COMPARABLE_SHEET in case of Net Cost Plus review
'
Public Const CONST_PLI_COMPANY_COLUMN           As String = "B"
Public Const CONST_PLI_SHEET_BASE_RANGE         As String = "B15"
'Public Const CONST_PLI_AVERAGE_COLUMN           As String = "D"
'Public Const CONST_PLI_CY_COLUMN                As String = "E"
'Public Const CONST_PLI_LY_COLUMN                As String = "F"
'Public Const CONST_PLI_LLY_COLUMN               As String = "H"
'Public Const CONST_PLI_COMPARABLE_COLUMN        As String = "I"
'
' Notice: Report layout changed in FY2024
' Date: 2024/5/7
'
Public Const CONST_PLI_AVERAGE_COLUMN           As String = "D"
Public Const CONST_PLI_CY_COLUMN                As String = "F"
Public Const CONST_PLI_LY_COLUMN                As String = "G"
Public Const CONST_PLI_LLY_COLUMN               As String = "I"
Public Const CONST_PLI_COMPARABLE_COLUMN        As String = "J"
Public Const CONST_PLI_COUNTRY_COLUMN           As String = "K"
Public Const CONST_PLI_COMPANY_PROPER_COLUMN    As String = "L"
Public Const CONST_PLI_REJECTION_REASON_COLUMN  As String = "M"

'
' Benchmark worksheet related layout
'
Public Const BMK_AVG_YEAR                       As Integer = 0
'Public Const BMK_CURRENT_YEAR                   As Integer = 1
'Public Const BMK_LAST_YEAR                      As Integer = 2
'Public Const BMK_YEAR_BEFORE_LAST_YEAR          As Integer = 4
'Public Const BMK_COMPARABLE_OFFSET              As Integer = 5
'
' Notice: Report layout changed in FY2024; Very annoying
' Date: 2024/5/7
'
Public Const BMK_CURRENT_YEAR                   As Integer = 2
Public Const BMK_LAST_YEAR                      As Integer = 3
Public Const BMK_YEAR_BEFORE_LAST_YEAR          As Integer = 5
Public Const BMK_COMPARABLE_OFFSET              As Integer = 6

'
' UserForm related constants
'
Public Const CONST_COMPARABLE_STATE_TBD         As String = "TBD"
Public Const CONST_COMPARABLE_STATE_NG          As String = "NG"
Public Const CONST_COMPARABLE_STATE_OK          As String = "OK"
Public Const CONST_COMPARABLE_STATE_CONDITION   As String = "Condition"

Public Const UNICODE_CHECK                      As Integer = 10004
Public Const UNICODE_FORBIDDEN                  As Integer = 8856

Public Const RR_SIG_DIFF                        As String = "Significantly different activities or products"
Public Const RR_BIG_RD_EXPENSE                  As String = "Consolidated and Unconsolidated Research And Development Expense / Total Net Sales is greater than or equal to 0.1%"
Public Const RR_BIG_MARKETING_EXPENSE           As String = "Consolidated and Unconsolidated Advertising Expense / Total Net Sales is greater than or equal to 0.1%"
Public Const RR_MISSING_DATA                    As String = "Missing Financial Data"
Public Const RR_THREE_YEAR_LOSS                 As String = "Operating income loss for 3 years"
Public Const RR_OTHERS                          As String = "Others"
Public Const RR_BLANK                           As String = ""


