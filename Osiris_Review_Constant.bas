Attribute VB_Name = "Osiris_Review_Constant"
'
'   Description: A module listing  Osiris data review associated constants
'
'   Date: 2024/6/15
'   Author: maoyi.fan@yapro.com.tw
'   Ver.: 0.1h
'   Revision History:
'       - 2024/6/15, 0.1h: Adjusted constant arrangement to accommodate dual operation conditions
'       - 2024/4/23, 0.1e: Support reload original company records
'       - 2024/4/15, 0.1b: First added
'
'   ToDo's:
'       1)
'
Option Explicit
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

'
' STATUS_COLUMN_OFFSET stores the offset from CONST_COMPANY_NAME_COLUMN to the Status column
' Note: This might be 12 for the Osiris query of one Advertisement or R&D search criterion; 13
'       in case of both criteria are set
'
'Public Const STATUS_COLUMN_OFFSET               As Integer = 13
'
' Layout constants associated with PLI Benchmark worksheet
' PLI Benchmark worksheet means OM_COMPARABLE_SHEET in case of Operating Margin review,
'                               NCP_COMPARABLE_SHEET in case of Net Cost Plus review
'
Public Const CONST_PLI_COMPANY_COLUMN           As String = "B"
Public Const CONST_PLI_SHEET_BASE_RANGE         As String = "B15"
'
' Either Advertisement or R&D filter criterion is specified
'
'Public Const CONST_PLI_AVERAGE_COLUMN           As String = "D"
'Public Const CONST_PLI_CY_COLUMN                As String = "E"
'Public Const CONST_PLI_LY_COLUMN                As String = "F"
'Public Const CONST_PLI_LLY_COLUMN               As String = "H"
'Public Const CONST_PLI_COMPARABLE_COLUMN        As String = "I"
'
' Notice: Report layout changed in FY2024
' Date: 2024/5/7
' Both Advertisement and R&D filter criteria are specified
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
'
' Either Advertisement or R&D filter criterion is specified
'
'Public Const BMK_AVG_YEAR                       As Integer = 0
'Public Const BMK_CURRENT_YEAR                   As Integer = 1
'Public Const BMK_LAST_YEAR                      As Integer = 2
'Public Const BMK_YEAR_BEFORE_LAST_YEAR          As Integer = 4
'Public Const BMK_COMPARABLE_OFFSET              As Integer = 5
'
' Notice: Report layout changed in FY2024; Very annoying
' Date: 2024/5/7
' Both Advertisement and R&D filter criteria are specified
'
'Public Const BMK_AVG_YEAR                       As Integer = 0
'Public Const BMK_CURRENT_YEAR                   As Integer = 2
'Public Const BMK_LAST_YEAR                      As Integer = 3
'Public Const BMK_YEAR_BEFORE_LAST_YEAR          As Integer = 5
'Public Const BMK_COMPARABLE_OFFSET              As Integer = 6

'
' UserForm related constants
'
Public Const CONST_COMPARABLE_STATE_TBD         As String = "TBD"
Public Const CONST_COMPARABLE_STATE_NG          As String = "NG"
Public Const CONST_COMPARABLE_STATE_OK          As String = "OK"
Public Const CONST_COMPARABLE_STATE_CONDITION   As String = "Condition"
Public Const CONST_COMPARABLE_STATE_EMPTY       As String = ""

Public Const UNICODE_CHECK                      As Integer = 10004
Public Const UNICODE_FORBIDDEN                  As Integer = 8856

Public Const RR_SIG_DIFF                        As String = "Significantly different activities or products"
Public Const RR_BIG_RD_EXPENSE                  As String = "Consolidated and Unconsolidated Research and Development Expense / Total Net Sales is greater than or equal to 1%"
Public Const RR_BIG_MARKETING_EXPENSE           As String = "Consolidated and Unconsolidated Advertising Expense / Total Net Sales is greater than or equal to 1%"
Public Const RR_MISSING_DATA                    As String = "Missing Financial Data"
Public Const RR_THREE_YEAR_LOSS                 As String = "Operating income loss for 3 years"
Public Const RR_OTHERS                          As String = "Others"
Public Const RR_BLANK                           As String = ""

'
' Trial code to test switch between two sets of constants
'
Public Const SINGLE_EXCLUSION_CRITERIA          As Integer = 1
Public Const DUAL_EXCLUSION_CRITERIA            As Integer = 2
Public Const PARAM1_SINGLE                      As String = "PARAM1 Single Criteria"
Public Const PARAM1_DUAL                        As String = "PARAM1 Dual Criteria"
Public Const PARAM2_SINGLE                      As Integer = 1
Public Const PARAM2_DUAL                        As Integer = 2
Public OP_PARAM1                                As String
Public OP_PARAM2                                As Integer
'
' Screening_Worksheet, 列表 (2) related layout variables
'
Public STATUS_COLUMN_OFFSET                             As Integer
Public PARAM_STATUS_COLUMN                              As String
Public SCREENING_WORKSHEET_BASE_RANGE                   As String
Public SCREENING_WORKSHEET_IDX_COLUMN                   As String
Public SCREENING_WORKSHEET_COMPANY_NAME_COLUMN          As String
Public SCREENING_WORKSHEET_TRADE_COLUMN                 As String
Public SCREENING_WORKSHEET_COMPANY_DESCRIPTION_COLUMN   As String
Public SCREENING_WORKSHEET_PNS_COLUMN                   As String
Public SCREENING_WORKSHEET_COUNTRY_CODE_COLUMN          As String
Public SCREENING_WORKSHEET_REVIEW_COLUMN                As String
Public SCREENING_WORKSHEET_STATUS_COLUMN                As String
Public SCREENING_WORKSHEET_COMMENT_COLUMN               As String
'
' PLI, 利潤率表, related layout variables
'
Public PLI_SHEET_COMPANY_COLUMN                         As String
Public PLI_SHEET_BASE_RANGE                             As String
Public PLI_SHEET_AVERAGE_COLUMN                         As String
Public PLI_SHEET_CY_COLUMN                              As String
Public PLI_SHEET_LY_COLUMN                              As String
Public PLI_SHEET_LLY_COLUMN                             As String
Public PLI_SHEET_COMPARABLE_COLUMN                      As String
Public PLI_SHEET_COUNTRY_COLUMN                         As String
Public PLI_SHEET_COMPANY_PROPER_COLUMN                  As String
Public PLI_SHEET_REJECTION_REASON_COLUMN                As String
'
' Benchmark worksheet related variables
'
Public BMK_AVG_YEAR                                     As Integer
Public BMK_CURRENT_YEAR                                 As Integer
Public BMK_LAST_YEAR                                    As Integer
Public BMK_YEAR_BEFORE_LAST_YEAR                        As Integer
Public BMK_COMPARABLE_OFFSET                            As Integer
'
' Description: Configure operation/layout parameters according to the number of screening criteria
' Keyboard Shortcut:
' Code Date: 2024/6/13
'
Public Sub configOpParam(sw As Integer)
    '
    ' Configure common operation parameters
    '
    SCREENING_WORKSHEET_BASE_RANGE = "B3"
    SCREENING_WORKSHEET_IDX_COLUMN = "A"
    SCREENING_WORKSHEET_COMPANY_NAME_COLUMN = "B"
    SCREENING_WORKSHEET_TRADE_COLUMN = "C"
    SCREENING_WORKSHEET_COMPANY_DESCRIPTION_COLUMN = "D"
    SCREENING_WORKSHEET_PNS_COLUMN = "E"
    SCREENING_WORKSHEET_COUNTRY_CODE_COLUMN = "F"
    
    PLI_SHEET_COMPANY_COLUMN = "B"
    PLI_SHEET_BASE_RANGE = "B15"
    PLI_SHEET_AVERAGE_COLUMN = "D"
    
    BMK_AVG_YEAR = 0
   
    '
    ' Configure operation parameters dependent on
    '
    If sw = Osiris_Review_Constant.SINGLE_EXCLUSION_CRITERIA Then
        Debug.Print "Assigning golbal parameter OP_PARAM1, OP_PARAM2 to SINGLE mode..."
        STATUS_COLUMN_OFFSET = 12
        SCREENING_WORKSHEET_REVIEW_COLUMN = "L"
        SCREENING_WORKSHEET_STATUS_COLUMN = "M"
        SCREENING_WORKSHEET_COMMENT_COLUMN = "N"
        
        PLI_SHEET_CY_COLUMN = "E"
        PLI_SHEET_LY_COLUMN = "F"
        PLI_SHEET_LLY_COLUMN = "H"
        PLI_SHEET_COMPARABLE_COLUMN = "I"
        PLI_SHEET_COUNTRY_COLUMN = "J"
        PLI_SHEET_COMPANY_PROPER_COLUMN = "K"
        PLI_SHEET_REJECTION_REASON_COLUMN = "L"
        
        BMK_CURRENT_YEAR = 1
        BMK_LAST_YEAR = 2
        BMK_YEAR_BEFORE_LAST_YEAR = 4
        BMK_COMPARABLE_OFFSET = 5
        
        OP_PARAM1 = PARAM1_SINGLE
        OP_PARAM2 = PARAM2_SINGLE
    Else
        Debug.Print "Assigning golbal parameter OP_PARAM1, OP_PARAM2 to DUAL mode..."
        STATUS_COLUMN_OFFSET = 13
        SCREENING_WORKSHEET_REVIEW_COLUMN = "M"
        SCREENING_WORKSHEET_STATUS_COLUMN = "N"
        SCREENING_WORKSHEET_COMMENT_COLUMN = "O"
        
        PLI_SHEET_CY_COLUMN = "F"
        PLI_SHEET_LY_COLUMN = "G"
        PLI_SHEET_LLY_COLUMN = "I"
        PLI_SHEET_COMPARABLE_COLUMN = "J"
        PLI_SHEET_COUNTRY_COLUMN = "K"
        PLI_SHEET_COMPANY_PROPER_COLUMN = "L"
        PLI_SHEET_REJECTION_REASON_COLUMN = "M"
        
        BMK_CURRENT_YEAR = 2
        BMK_LAST_YEAR = 3
        BMK_YEAR_BEFORE_LAST_YEAR = 5
        BMK_COMPARABLE_OFFSET = 6
        
        OP_PARAM1 = PARAM1_DUAL
        OP_PARAM2 = PARAM2_DUAL
    End If
End Sub
