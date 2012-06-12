Attribute VB_Name = "modPublic"
Option Explicit

Public ENV                          As New clsParameter
Public QUEUE                        As New clsCommand
Public RANK_OBJ                     As New clsRank
Public EQP                          As New clsEQ_Information
'Public FTP_CLIENT                   As New clsFTP

Type DEVICE_INFO                    'Device Information Structure
    PORT_NO                         As Integer
    DEVICE_NAME                     As String
    DEVICE_STATE                    As String
End Type

Type CST_INFO_ELEMENTS              'Cassette MES Data Structure
    CSTID                           As String
    PFCD                            As String
    OWNER                           As String
    PROCESS_NUM                     As String
    PORTID                          As String
    PORT_TYPE                       As String
    DESTINATION_FAB                 As String
    PANEL_COUNT                     As String
    RMANO                           As String
    OQCNO                           As String
    SOURCE_FAB                      As String
    CST_SPARE(1 To 5)               As String
End Type

Type PANEL_INFO_ELEMENTS            'Panel MES Data Structure
    SLOT_NUM                        As String
    PANELID                         As String
    LIGHT_ON_PANEL_GRADE            As String
    LIGHT_ON_REASON_CODE            As String
    CELL_LINE_RESCUE_FLAG           As String
    CELL_REPAIR_JUDGE_GRADE         As String
    TFT_REPAIR_GRADE                As String
    CF_PANELID                      As String
    CF_PANEL_OX_INFORMATION            As String
    PANEL_OWNER_TYPE                As String
    ABNORMAL_CF                     As String
    ABNORMAL_TFT                    As String
    ABNORMAL_LCD                    As String
    GROUP_ID                        As String
    REPAIR_REWORK_COUNT             As String
    CARBONIZATION_FLAG              As String
    CARBONIZATION_GRADE             As String
    CARBONIZATION_REWORK_COUNT      As String
    POLARIZER_REWORK_COUNT          As String
    X_TOTAL_PIXEL                   As String
    Y_TOTAL_PIXEL                   As String
    X_ONE_PIXEL_LENGTH              As String
    Y_ONE_PIXEL_LENGTH              As String
    LCD_Q_TAP_LOT_GROUPID           As String
    SK_FLAG                         As String
    ODK_AK_FLAG                     As String
    BPAM_REWORK_FLAG                As String
    LCD_BRIGHT_DOT_FLAG             As String
    CF_PS_HEIGHT_ERR_FLAG           As String
    CF_R_DEFECT_CODE                As String
    PI_INSPECTION_NG_FLAG           As String
    PI_OVER_BAKE_FLAG               As String
    PI_OVER_Q_TIME_FLAG             As String
    ODF_OVER_BAKE_FLAG              As String
    ODF_OVER_Q_TIME_FLAG            As String
    HVA_OVER_BAKE_FLAG              As String
    HVA_OVER_Q_TIME_FLAG            As String
    SEAL_INSPECTION_FLAG            As String
    ODF_CHECKER_FLAG                As String
    ODF_DOOR_OPEN_FLAG              As String
    LOT1_OPERATION_MODE             As String
    LOT2_OPERATION_MODE             As String
    PRODUCTID                       As String
    OWNERID                         As String
    PREPROCESSID                    As String
    SPARE(1 To 10)                  As String
'    AAAAAA                          As String
End Type

Type JOB_DATA_STRUCTURE             'JOB MES Data Structure
    CST_SEQUENCE                    As String
    JOB_SEQUENCE                    As String
    CIM_MODE                        As String
    JOB_JUDGE                       As String
    JOB_GRADE                       As String
    GLASSID                         As String
    BURR_CHECK_JUDGE                As String
    BEVELING_JUDGE                  As String
    CLEANER_M_PORT_JUDGE            As String
    TEST_CV_JUDGE                   As String
    SAMPLING_SLOT_FLAG              As String
    PROCESS_INPUT_FLAG              As String
    NEED_GRINDING_FLAG              As String
    MISALIGNMENT_FLAG               As String
    SMALL_MULTI_PANEL_FLAG          As String
    AK_FLAG                         As String
    SK_FLAG                         As String
    NO_MATCH_GLASS_IN_BC_FLAG       As String
    CASSETTE_SETTING_CODE           As String
    ABNORMAL_FLAG_CODE              As String
    LIGHT_ON_REASON_CODE            As String
    PANEL_NG_FLAG                   As String
    CUT_FLAG                        As String
    RESERVED                        As String
End Type

Type SHARE_DATA_STRUCTURE           'SHARE MES Data Structure
    PANELID                         As String
    GLASS_TYPE                      As String
    PRODUCTID                       As String
    PROCESSID                       As String
    RECIPEID                        As String
    SALE_ORDER                      As String
    CF_GLASSID                      As String
    ARRAY_LOTID                     As String
    ARRAY_GLASSID                   As String
    CF_GLASS_INFO                   As String
    TFT_PANEL_JUDGE                 As String
    PRE_PROCESSID1                  As String
    GROUPID                         As String
    TRANSFER_TIME                   As String
End Type

Type RANK_DATA_STRUCTURE            'Rank Data Information Structure
    RANK_DIVISION                   As String
    PANEL_IN_OR_OUTSIDE             As String
    DEFECT_CODE                     As String
    DEFECT_NAME                     As String
    DEFECT_DIVISION                 As String
    DEFECT_TYPE                     As String
    JUDGE_OR_NOT                    As String
    USE_XY                          As String
    DETAIL_DIVISION                 As String
    ACCUMULATION                    As String
    ADDRESS_COUNT                   As String
    ODF                             As String
    PRIORITY                        As Integer
    POP_UP                          As String
    Rank(30)                        As String
End Type

Type GRADE_DATA_STRUCTURE           'Grade/Rank Maching Data Information Structure
    RANK_DIVISION                   As String
    DEFECT_CODE                     As String
    GRADE                           As String
    RANK                            As String
End Type

Type RANK_PRIORITY_STRUCTURE        'Priority Data Structure in each Ranks
    RANK                            As String
    PRIORITY                        As Integer
    RANK_COUNT                      As Integer
End Type

Type ITEM_CONTROL                   'Item Enable/Disable Flag Structure
    ITEM_NAME                       As String
    ENABLE_DISABLE                  As String
End Type

Type DEFECT_DATA_STRUCTURE          'Defect Data Structure
    PANELID                         As String
    DEFECT_NO                       As Integer
    DEFECT_CODE                     As String
    DEFECT_NAME                     As String
    COLOR                           As String
    GRAY_LEVEL                      As Integer
    DETAIL_DIVISION                 As String
    DATA_ADDRESS(1 To 3)            As String
    GATE_ADDRESS(1 To 3)            As String
    RANK                            As String
    GRADE                           As String
    PRIORITY                        As Integer
    ACCUMULATION                    As Integer
End Type

Type DEFECT_PRIORITY_STRUCTURE      'Defect Priority Information Structure
    DEFECT_TYPE                     As String
    DEFECT_PRIORITY                 As Integer
    DEFECT_INDEX                    As Integer
    DEFECT_GRADE                    As String
    DEFECT_RANK                     As String
    DEFECT_CODE                     As String
End Type

Type DEFECT_FILE_HEADER             'Defect File Header Structure
    JPS_VERSION                     As String
    FILE_CREATE_TIME                As String
    EQUIP_TYPE                      As String
    EQ_ID                           As String
    SUBEQ_ID                        As String
End Type

Type DEFECT_FILE_PANEL_DATA         'Defect File Panel Data Structure
    PANELID                         As String
    GLASS_TYPE                      As String
    PRODUCT_ID                      As String
    PROCESS_ID                      As String
    RECIPE_ID                       As String
    SALEORDER                       As String
    CF_GLASS_ID                     As String
    ARRAY_LOT_ID                    As String
    ARRAY_GLASS_ID                  As String
    CF_GLASS_OX_INFO                As String
    TFT_PANEL_JUDGE                 As String
    GROUP_ID                        As String
End Type

Type DEFECT_FILE_EQP_PANEL_DATA     'Defect File EQP Panel Data Structure
    RECIPE_NO                       As String
    START_TIME                      As String
    END_TIME                        As String
    OPERATOR_ID                     As String
    TACT_TIME                       As String
    MAIN_DEFECT_CODE                As String
    TOTAL_POINT_DEFECT_COUNT        As String
    OPERATION_MODE                  As String
    RJS_NO                          As String
    INSPECTION_TIME                 As String
    TRANSFER_TIME                   As String
End Type

Type DEFECT_FILE_PANEL_SUMMARY      'Defect File Panel Summary Data Structure
    PANELID                         As String
    JUDGE_RANK                      As String
    MAIN_DEFECT_CODE                As String
    DATA_TOTAL_PIXEL                As String
    GATE_TOTAL_PIXEL                As String
    DRIVE_TYPE                      As String
    BACKLIGHT                       As String
    LIGHT_ON_PRE_GRADE     As String
    LIGHT_ON_TARGET_REASON_TYPE     As String
    SLOT_ID                         As String
End Type

Type DEFECT_FILE_DEFECT_DATA        'Defect File Defect Data Structure
    PANELID                         As String
    DEFECT_NO                       As String
    DEFECT_CODE                     As String
    COLOR                           As String
    PANEL_GRADE                     As String
    GRAY_LEVEL                      As String
    DEFECT_NAME                     As String
    DATA_X1                         As String
    GATE_Y1                         As String
    DATA_X2                         As String
    GATE_Y2                         As String
    DATA_X3                         As String
    GATE_Y3                         As String
    PANEL_COORDINATE_X1             As String
    PANEL_COORDINATE_Y1             As String
    PANEL_COORDINATE_X2             As String
    PANEL_COORDINATE_Y2             As String
    PANEL_COORDINATE_X3             As String
    PANEL_COORDINATE_Y3             As String
    GLASS_COORDINATE_X1             As String
    GLASS_COORDINATE_Y1             As String
    GLASS_COORDINATE_X2             As String
    GLASS_COORDINATE_Y2             As String
    GLASS_COORDINATE_X3             As String
    GLASS_COORDINATE_Y3             As String
    MARK_TYPE                       As String
    EACH_PTN_INSPECTION_TIME()      As String
    PATTERN_NAME()                  As String
    PATTERN_COUNT                   As Integer
End Type

Type DEFECT_FILE_LCD_DATA           'Defect File LCD Data Structure
    PANELID                         As String
    LIGHT_ON_SOURCE_GRADE           As String
    LIGHT_ON_SOURCE_REASON_CODE     As String
    TOTAL_LIGHT_ON_DEFECT_COUNT     As String
End Type

Type DEFECT_FILE_PANEL_PDS_SUMMARY  'Defect File Panel PDS Summary Data Structure
    PANELID                         As String
    PARAMETER_NAME                  As String       'Recipe Name
    AVG                             As String       'CST_MES_DATA's spare1 25bytes
    MIN                             As String       'CST_MES_DATA's spare2 25bytes
    MAX                             As String       'CST_MES_DATA's spare3 25bytes
    STD                             As String       'CST_MES_DATA's spare4 25bytes
    COUNT                           As String       'CST_MES_DATA's spare5 25bytes
End Type

Type PFCD_ADDRESS_STRUCTURE         'PFCD Address Data Structure
    PRODUCT_ID                      As String
    PANEL_NO                        As String
    W                               As Double
    L                               As Double
    B1                              As Double
    B2                              As Double
    XC                              As Double
    YC                              As Double
    XO                              As Double
    YO                              As Double
    ORIGIN_LOCATION                 As String
    SOURCE_DIRECTION                As String
End Type

Type USER_DATA                      'User Data Structure
    USER_ID                         As String
    USER_NAME                       As String
    ID_CARD_CODE                    As String
    USER_PW1                        As String
    USER_PW2                        As String
    USER_LEVEL                      As String
End Type

Type VERSION_DATA                   'Version Data Structure
    MACHINE_ID                      As String
    JPS_VERSION                     As String
    EQ_VERSION                      As String
    JPS_NAME                        As String
    INSTALL_DAY                     As String
    USER                            As String
    JPS_SETUP_PATH                  As String
    JPS_LOG_PATH                    As String
    JPS_SERVER_PATH                 As String
End Type

Type PFCD_DATA                      'PFCD Data Structure
    PFCD                            As String
    X_PIXEL_LENGTH                  As String
    Y_PIXEL_LENGTH                  As String
    DATA                            As String
    GATE                            As String
    CSTC                            As String
    MAX_PANEL                       As String
    PANEL_TYPE                      As String
End Type

Type EQTYPE_DATA                    'EQP Type Data Structure
    PC_NAME                         As String
    PC_IP                           As String
    MACHINE_NAME                    As String
    EQ_MODEL                        As String
    FS_DRIVE                        As String
    FS_IP                           As String
    FS_USER_NAME                    As String
    FS_PASSWORD                     As String
End Type

Type FS_PATH_DATA                   'FS Path Data Structure
    EQTYPE                          As String
    PFCD_PID                        As String
    RANK                            As String
    USER                            As String
    PTN_LIST                        As String
    VERSION                         As String
    TA_HISTORY                      As String
End Type

Type DEFECT_CODE_HIDE_DATA          'Defect Code Hide Data Structure
    PTN_LIST                        As String
    DEFECT_CODE(1 To 10)            As String
End Type

Type USER_LOGON_DATA                'User LogOn Data Structure
    LOGON_DATE                      As String
    LOGON_TIME                      As String
    USER_ID                         As String
    USER_NAME                       As String
End Type

Type PANEL_DATA                     'Panel Data Structure
    KEYID                           As String       'KEYID = PANELID + _ + YYYYMMDD + HHMMSS
    TIME                            As String
    PANELID                         As String
    PANEL_RANK                      As String
    PANEL_GRADE                     As String
    PANEL_LOSSCODE                  As String
    LOSSCODE_NAME                   As String
    USER_NAME                       As String
    PANEL_TYPE1                     As String
    PANEL_TYPE2                     As String
    PATH                            As String
    FILENAME                        As String
    RUN_DATE                        As Long
    RUN_TIME                        As Long
    TACT_TIME                       As Long
End Type

Type PUBLIC_NOTICE                  'Bulletin Board Message Structure
    DATE                            As String
    TIME                            As String
    MESSAGE()                       As String
    UPDATED                         As Boolean
    MESSAGE_COUNT                   As Integer
End Type

Public Type PATTERN_LIST_DATA              'Pattern List Data Structure
    FILENAME                        As String
    PATTERN_CODE                    As String
    PATTERN_NAME                    As String
    DELAY_TIME                      As Integer
    LEVEL                           As Integer
    DH                              As Integer
    DL                              As Integer
    VGH                             As Integer
    VGL                             As Integer
    RESCUE_HIGH                     As Integer
    RESCUE_LOW                      As Integer
    VCOM                            As Integer
    INSPECTION_START                As Double
    INSPECTION_END                  As Double
    INSPECTION_TIME                 As Double
End Type

Public Type AUTO_ALARM_DATA
    PROCESS_NUM                     As String
    PFCD                            As String
    DEFECT_CODE                     As String
    RANK                            As String
    COUNT_TIME                      As Integer
    COUNT                           As Integer
    ALARM_TEXT                      As String
    CURRENT_COUNT                   As Integer
    EXPIRY_DATE                     As Long
    EXPIRY_TIME                     As Long
End Type

Public Type COUNT_CHANGE_DATA
    FINAL_GRADE                     As String
    NEW_GRADE                       As String
    COUNT                           As Integer
    CURRENT_COUNT                   As Integer
End Type

Public pubDefect_Count              As Integer                          '2012.03.26 Added by K.H.KIM

Public pubCST_INFO                  As CST_INFO_ELEMENTS
Public pubPANEL_INFO                As PANEL_INFO_ELEMENTS
Public pubJOB_INFO                  As JOB_DATA_STRUCTURE
Public pubSHARE_INFO                As SHARE_DATA_STRUCTURE
Public pubDEFECT_DATA()             As DEFECT_DATA_STRUCTURE            '2012.03.26 Added by K.H.KIM

'leo
Public RankLevel()                  As String

