Attribute VB_Name = "Osiris_Review_Constant"
'
'   Description: A module listing  Osiris data review associated constants
'
'   Date: 2024/9/5
'   Author: maoyi.fan@yapro.com.tw
'   Ver.: 0.1k
'   Revision History:
'       - 2024/9/5,  0.1k: Unified the way to handle different report layout due to its dynamic behavior
'       - 2024/8/17, 0.1j: Column layout changes all the time, adjust data parsing column variables based on
'                          CONST_XXXX. Walk through all these CONST definitions before a new comparables
'                          screening
'       - 2024/6/15, 0.1h: Adjusted constant arrangement to accommodate dual operation conditions
'       - 2024/4/23, 0.1e: Support reload original company records
'       - 2024/4/15, 0.1b: First added
'
'   ToDo's:
'       1) Configure PLI_SHEET_xxx_xxxxx operation parameters using CONST_PLI_xxx_xxxxx so that
'          it'll be easier for gadget functions to parse the Profit Level Indicator related worksheet
'
Option Explicit
'
' Constant definitions associated with data screening of Osiris search/rejection results
'
Public Const CONST_OM_PLI                       As String = "Operating Margin"
Public Const CONST_NCP_PLI                      As String = "Net Cost Plus"
Public Const OM_DETAILS_SHEET                   As String = "Benchmark 1"           ' confirm this before starting any new comparables review
Public Const OM_COMPARABLE_SHEET                As String = "OM_Screening"
Public Const NCP_DETAILS_SHEET                  As String = "Benchmark 4"           ' confirm this before starting any new comparables review
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
Public Const CONST_SCREENING_COMPANY_COLUMN     As String = "B"
Public Const CONST_SCREENING_FIRST_DATA_ROW     As String = "3"
Public Const CONST_IDX_COLUMN                   As String = "A"
Public Const CONST_COMPANY_NAME_COLUMN          As String = "B"         ' NAME
Public Const CONST_TRADE_COLUMN                 As String = "G"         ' TRADE_DESCRIPTION_EN
Public Const CONST_COMPANY_DESCRIPTION_COLUMN   As String = "M"         ' DESCRIPTION_HISTORY
Public Const CONST_PNS_COLUMN                   As String = "L"         ' PRODUCTS_SERVICES
Public Const CONST_COUNTRY_CODE_COLUMN          As String = "C"         ' COUNTRY_ISO_CODE
Public Const CONST_MANUAL_REVIEW_COLUMN         As String = "K"         ' MANUAL_REVIEW
Public Const CONST_STATUS_COLUMN                As String = "Q"         ' STATUS
Public Const CONST_COMMENT_COLUMN               As String = "R"         ' review Comments column

'
' STATUS_COLUMN_OFFSET stores the offset from CONST_COMPANY_NAME_COLUMN to the Status column
' It changes from review to review due to different criteria selected and different column
' layout of the report! (2024/9/4)
'
' Layout constants associated with PLI Benchmark worksheet
' PLI Benchmark worksheet means OM_COMPARABLE_SHEET in case of Operating Margin review,
'                               NCP_COMPARABLE_SHEET in case of Net Cost Plus review
'
Public Const CONST_PLI_COMPANY_COLUMN           As String = "B"
Public Const CONST_PLI_FIRST_DATA_ROW           As String = "15"
'
' Notice: PLI(Benchmark) Report layout changed
' Date: 2024/5/7
'
Public Const CONST_PLI_AVERAGE_COLUMN           As String = "D"         ' PLI Average Column
Public Const CONST_PLI_CY_COLUMN                As String = "F"         ' PLI Current Year Column
Public Const CONST_PLI_LY_COLUMN                As String = "G"         ' PLI Last Year Column
Public Const CONST_PLI_LLY_COLUMN               As String = "I"         ' PLI Year Before Last Year Column
Public Const CONST_PLI_COMPARABLE_COLUMN        As String = "J"         ' Comparable Status Column
Public Const CONST_PLI_COUNTRY_COLUMN           As String = "K"         ' Country Code Column
Public Const CONST_PLI_COMPANY_PROPER_COLUMN    As String = "L"         ' Compnay Name in Proper Form Column
Public Const CONST_PLI_REJECTION_REASON_COLUMN  As String = "M"         ' Rejection Reason Column

'
' Benchmark(PLI) worksheet related layout constants
' Constants related to the column offsets to locate the PLI ratios for each year
'
Public Const CONST_BMK_AVG_YEAR_OFFSET                As Integer = 0

'
' UserForm related constants
'
Public Const CONST_COMPARABLE_STATE_TBD         As String = "TBD"
Public Const CONST_COMPARABLE_STATE_NG          As String = "NG"
Public Const CONST_COMPARABLE_STATE_OK          As String = "OK"
Public Const CONST_COMPARABLE_STATE_CONDITION   As String = "Condition"
Public Const CONST_COMPARABLE_STATE_EMPTY       As String = ""
Public Const CONST_COMPARABLE_STATE_NEXT        As String = "NEXT"

Public Const UNICODE_CHECK                      As Integer = 10004
Public Const UNICODE_FORBIDDEN                  As Integer = 8856
Public Const UNICODE_UNKNOWN                    As Integer = 8413     'Osiris displays Unicode 8413 initially for those records
                                                                      'not banned, not checked

Public Const RR_SIG_DIFF                        As String = "Significantly different business activities or products"
Public Const RR_BIG_RD_EXPENSE                  As String = "Consolidated and Unconsolidated Research and Development Expense / Total Net Sales is greater than or equal to 1%"
Public Const RR_BIG_MARKETING_EXPENSE           As String = "Consolidated and Unconsolidated Advertising Expense / Total Net Sales is greater than or equal to 1%"
Public Const RR_MISSING_DATA                    As String = "Missing Financial Data"
Public Const RR_THREE_YEAR_LOSS                 As String = "Operating income loss for 3 years"
Public Const RR_OTHERS                          As String = "Others"
Public Const RR_BLANK                           As String = ""

'
' Trial code to test switch between two sets of constants
'
Public Const DEFAULT_COLUMN_LAYOUT              As Integer = 0
Public Const SINGLE_EXCLUSION_CRITERIA          As Integer = 1
Public Const DUAL_EXCLUSION_CRITERIA            As Integer = 2
Public Const PARAM1_SINGLE                      As String = "PARAM1 Single Criteria"
Public Const PARAM1_DUAL                        As String = "PARAM1 Dual Criteria"
Public Const PARAM1_DEFAULT                     As String = "PARAM1 Default"
Public Const PARAM2_SINGLE                      As Integer = 1
Public Const PARAM2_DUAL                        As Integer = 2
Public Const PARAM2_DEFAULT                     As Integer = 0
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
' Arguments
'     sw: SINGLE_EXCLUSION_CRITERIA 單一排除條件報表格式配置
'         DUAL_EXCLUSION_CRITERIA 雙重排除條件報表格式配置
'         DEFAULT_COLUMN_LAYOUT or others 預設報表格式配置，根據上述 Screening_Worksheet 或是 Benchmark 格式配置
'
Public Sub configOpParam(sw As Integer)
    '
    ' Configure common operation parameters
    '
    ' Screening Worksheet portion
    SCREENING_WORKSHEET_BASE_RANGE = Osiris_Review_Constant.CONST_SCREENING_COMPANY_COLUMN & _
                                     Osiris_Review_Constant.CONST_SCREENING_FIRST_DATA_ROW
    SCREENING_WORKSHEET_IDX_COLUMN = Osiris_Review_Constant.CONST_IDX_COLUMN
    SCREENING_WORKSHEET_COMPANY_NAME_COLUMN = Osiris_Review_Constant.CONST_COMPANY_NAME_COLUMN
    SCREENING_WORKSHEET_TRADE_COLUMN = Osiris_Review_Constant.CONST_TRADE_COLUMN
    SCREENING_WORKSHEET_COMPANY_DESCRIPTION_COLUMN = Osiris_Review_Constant.CONST_COMPANY_DESCRIPTION_COLUMN
    SCREENING_WORKSHEET_PNS_COLUMN = Osiris_Review_Constant.CONST_PNS_COLUMN
    SCREENING_WORKSHEET_COUNTRY_CODE_COLUMN = Osiris_Review_Constant.CONST_COUNTRY_CODE_COLUMN
    '
    ' Benchmark (Profit Level Index) worksheet portion
    PLI_SHEET_COMPANY_COLUMN = Osiris_Review_Constant.CONST_PLI_COMPANY_COLUMN
    PLI_SHEET_BASE_RANGE = Osiris_Review_Constant.CONST_PLI_COMPANY_COLUMN & _
                           Osiris_Review_Constant.CONST_PLI_FIRST_DATA_ROW
    Debug.Print "PLI benchmark worksheet base: " & PLI_SHEET_BASE_RANGE
    PLI_SHEET_AVERAGE_COLUMN = Osiris_Review_Constant.CONST_PLI_AVERAGE_COLUMN
    '
    ' BMK_AVG_YEAR specifies the column offset for which quartile calculation is based
    ' on. See the implementation of the function Osiris_Review_Gadgets.DoComparableQuartile
    '
    BMK_AVG_YEAR = Osiris_Review_Constant.CONST_BMK_AVG_YEAR_OFFSET
       
    '
    ' Configure operation parameters dependent on number of exclusion criteria
    ' Note: Determining column layout according to number of exclusion criteria is not always
    '       correct based on those Osiris query results. Added a default setting based on
    '       CONST_XXXX definitions (2024/8/17)
    '
    If sw = Osiris_Review_Constant.SINGLE_EXCLUSION_CRITERIA Then
        Debug.Print "Assigning column parsing related constants to SINGLE_EXCLUSION_CRITERIA mode..."
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
    ElseIf sw = Osiris_Review_Constant.DUAL_EXCLUSION_CRITERIA Then
        Debug.Print "Assigning column parsing related constants to DUAL_EXCLUSION_CRITERIA mode..."
        STATUS_COLUMN_OFFSET = 13
        SCREENING_WORKSHEET_REVIEW_COLUMN = "M"
        SCREENING_WORKSHEET_STATUS_COLUMN = "N"
        SCREENING_WORKSHEET_COMMENT_COLUMN = "O"
        '
        ' Notice: The following column arrangement might be different from query to query
        '         Check this before continuing the company screening process!c
        '
        PLI_SHEET_CY_COLUMN = "F"
        PLI_SHEET_LY_COLUMN = "G"
        PLI_SHEET_LLY_COLUMN = "I"
        PLI_SHEET_COMPARABLE_COLUMN = "J"
        PLI_SHEET_COUNTRY_COLUMN = "K"
        PLI_SHEET_COMPANY_PROPER_COLUMN = "L"    ' proper(Company Name)
        PLI_SHEET_REJECTION_REASON_COLUMN = "M"

        BMK_CURRENT_YEAR = 2
        BMK_LAST_YEAR = 3
        BMK_YEAR_BEFORE_LAST_YEAR = 5
        BMK_COMPARABLE_OFFSET = 6

        OP_PARAM1 = PARAM1_DUAL
        OP_PARAM2 = PARAM2_DUAL
    Else
        Debug.Print "Assigning column parsing related constants to DEFAULT mode..."
        '
        ' Screening_Worksheet column layout related constants
        '
        STATUS_COLUMN_OFFSET = Asc(Osiris_Review_Constant.CONST_STATUS_COLUMN) - Asc(Osiris_Review_Constant.CONST_COMPANY_NAME_COLUMN)
        SCREENING_WORKSHEET_REVIEW_COLUMN = Osiris_Review_Constant.CONST_MANUAL_REVIEW_COLUMN
        SCREENING_WORKSHEET_STATUS_COLUMN = Osiris_Review_Constant.CONST_STATUS_COLUMN
        SCREENING_WORKSHEET_COMMENT_COLUMN = Osiris_Review_Constant.CONST_COMMENT_COLUMN
        '
        ' Benchmark (PLI) worksheet column layout related constants
        PLI_SHEET_CY_COLUMN = Osiris_Review_Constant.CONST_PLI_CY_COLUMN
        PLI_SHEET_LY_COLUMN = Osiris_Review_Constant.CONST_PLI_LY_COLUMN
        PLI_SHEET_LLY_COLUMN = Osiris_Review_Constant.CONST_PLI_LLY_COLUMN
        '
        ' Additional columns for comparables and rejected companies in TP reports
        PLI_SHEET_COMPARABLE_COLUMN = Osiris_Review_Constant.CONST_PLI_COMPARABLE_COLUMN
        PLI_SHEET_COUNTRY_COLUMN = Osiris_Review_Constant.CONST_PLI_COUNTRY_COLUMN
        PLI_SHEET_COMPANY_PROPER_COLUMN = Osiris_Review_Constant.CONST_PLI_COMPANY_PROPER_COLUMN
        PLI_SHEET_REJECTION_REASON_COLUMN = Osiris_Review_Constant.CONST_PLI_REJECTION_REASON_COLUMN
        '
        ' Column defitions to locate PLI ratios of each year in Benchmark worksheet
        BMK_CURRENT_YEAR = Asc(Osiris_Review_Constant.CONST_PLI_CY_COLUMN) - Asc(Osiris_Review_Constant.CONST_PLI_AVERAGE_COLUMN)
        BMK_LAST_YEAR = Asc(Osiris_Review_Constant.CONST_PLI_LY_COLUMN) - Asc(Osiris_Review_Constant.CONST_PLI_AVERAGE_COLUMN)
        BMK_YEAR_BEFORE_LAST_YEAR = Asc(Osiris_Review_Constant.CONST_PLI_LLY_COLUMN) - Asc(Osiris_Review_Constant.CONST_PLI_AVERAGE_COLUMN)
        BMK_COMPARABLE_OFFSET = Asc(Osiris_Review_Constant.CONST_PLI_COMPARABLE_COLUMN) - Asc(Osiris_Review_Constant.CONST_PLI_AVERAGE_COLUMN)
        '
        ' Global operation parameter string and integer for operation mode display only
        OP_PARAM1 = Osiris_Review_Constant.PARAM1_DEFAULT
        OP_PARAM2 = Osiris_Review_Constant.PARAM2_DEFAULT
    End If
End Sub
