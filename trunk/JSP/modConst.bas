Attribute VB_Name = "modConst"
Option Explicit

'RS232 Command define
Public Const cSTX = "@"
Public Const cETX = ","

'Data type define
Public Const cDB_TEXT = 0
Public Const cDB_BOOLEAN = 1
Public Const cDB_INTEGER = 2
Public Const cDB_LONG = 3
Public Const cDB_DOUBLE = 4

'Device Status define
Public Const cDEVICE_ENABLE = "E"
Public Const cDEVICE_DISABLE = "D"
Public Const cDEVICE_ONLINE = "ON LINE"
Public Const cDEVICE_OFFLINE = "OFF LINE"

'Command Queue define
Public Const cQUEUE_MAX = 50
Public Const cCOMMAND_QUEUE_INVALID_PORT_NO = -1
Public Const cCOMMAND_QUEUE_NORMALCY = 0
Public Const cCOMMAND_QUEUE_EMPTY = 1
Public Const cCOMMAND_QUEUE_FULL = 2

'Data Size define
Public Const cSIZE_BACKLIGHT_VALUE = 5
Public Const cSIZE_ALARM_CODE = 4
Public Const cSIZE_ALARM_TEXT = 100
Public Const cSIZE_TIME = 14
Public Const cSIZE_PANELDRIVETYPE = 2
Public Const cSIZE_INFO_LENGTH = 3
Public Const cSIZE_JOB_DATA_LENGTH = 73
Public Const cSIZE_SHARE_DATA_LENGTH = 207
Public Const cSIZE_PFCD = 12
Public Const cSIZE_OWNER = 1
Public Const cSIZE_RUNMODE = 2
Public Const cSIZE_MACHINENAME = 8
Public Const cSIZE_USERNAME = 8
Public Const cSIZE_USERCODE = 10
Public Const cSIZE_WORKNO = 16
Public Const cSIZE_PANELTYPE = 16
Public Const cSIZE_GRADE = 2
Public Const cSIZE_LOSSCODE = 5
Public Const cSIZE_PNL_JUDGEMODE = 1
Public Const cSIZE_FLAG = 1
Public Const cSIZE_RUNNINGSTATUS = 3
Public Const cSIZE_DCRSTATUS = 3
Public Const cSIZE_PANELID = 12
Public Const cSIZE_CSTINFO = 3
Public Const cSIZE_PANELINFO = 3
Public Const cSIZE_CSTID_MES = 8
Public Const cSIZE_PRODUCTID_MES = 12
Public Const cSIZE_OWNER_MES = 4
Public Const cSIZE_OWNERID_MES = 6
Public Const cSIZE_PREPROCESSID = 4
Public Const cSIZE_PROCESSNUM_MES = 4
Public Const cSIZE_PORTID_MES = 2
Public Const cSIZE_PORTTYPE_MES = 2
Public Const cSIZE_DESTFAB_MES = 6
Public Const cSIZE_PANELCOUNT_MES = 5
Public Const cSIZE_RMANO_MES = 12
Public Const cSIZE_OQCNO_MES = 12
Public Const cSIZE_SOURCE_FAB_MES = 6
Public Const cSIZE_CST_SPARE_MES = 25
Public Const cSIZE_SLOTNO_MES = 4
Public Const cSIZE_LIGHTON_PNL_GRADE_MES = 2
Public Const cSIZE_LIGHTON_REASON_MES = 5
Public Const cSIZE_CELLREPAIR_GRADE_MES = 2
Public Const cSIZE_TFT_REPAIR_GRADE_MES = 2
Public Const cSIZE_CF_PNLINFO = 6
Public Const cSIZE_PANEL_OWNER_TYPE_MES = 1
Public Const cSIZE_ABNORMAL_MES = 4
Public Const cSIZE_ABNORMAL_LCD_MES = 12
Public Const cSIZE_GROUPID_MES = 12
Public Const cSIZE_WORKORDER_MES = 12
Public Const cSIZE_PANELANGLE_MES = 1
Public Const cSIZE_REWORKCOUNT_MES = 2
Public Const cSIZE_LINKKEY_MES = 4
Public Const cSIZE_TOTALPIXEL_MES = 5   'This data not received from EQ. Read PFCD.PID from Server. Change read module
Public Const cSIZE_ONEPIXEL_LENGTH_MES = 6
Public Const cSIZE_QTAPLOT_MES = 4
Public Const cSIZE_SK_FLAG_MES = 4
Public Const cSIZE_CF_R_DEFECT_CODE_MES = 4
Public Const cSIZE_ODK_AK_FLAG_MES = 4
Public Const cSIZE_BPAM_REWORK_FLAG = 2
Public Const cSIZE_LCD_BRIGHT_DOT_FLAG = 2
Public Const cSIZE_LOT_OPERATIONMODE = 2
Public Const cSIZE_CST_MES_DATA = "079"
Public Const cSIZE_PANEL_MES_DATA = "169"
Public Const cSIZE_CARBONIZATION_FLAG = 1
Public Const cSIZE_CARBONIZATION_GRADE = 2
Public Const cSIZE_CARBONIZATION_REWORK_COUNT = 2

'JOB Data define
Public Const cSIZE_CST_SEQUENCE_JOB = 5
Public Const cSIZE_JOB_SEQUENCE_JOB = 5
Public Const cSIZE_CIM_MODE_JOB = 1
Public Const cSIZE_JOB_JUDGE_JOB = 2
Public Const cSIZE_JOB_GRADE_JOB = 2
Public Const cSIZE_GLASSID_JOB = 12
Public Const cSIZE_BURR_CHECK_JUDGE_JOB = 2
Public Const cSIZE_BEVELING_JUDGE_JOB = 2
Public Const cSIZE_CLEANER_M_PORT_JUDGE_JOB = 2
Public Const cSIZE_TEST_CV_JUDGE_JOB = 2
Public Const cSIZE_FLAG_JOB = 1
Public Const cSIZE_CASSETTE_SETTING_CODE_JOB = 4
Public Const cSIZE_ABNORMAL_FLAG_CODE_JOB = 4
Public Const cSIZE_LIGHT_ON_REASON_CODE_JOB = 5
Public Const cSIZE_PANEL_NG_FLAG_JOB = 2
Public Const cSIZE_CUT_FLAG_JOB = 2
Public Const cSIZE_RESERVED_JOB = 10
Public Const cSIZE_JOB_DATA = "073"

'SHARED Data define
Public Const cSIZE_PANELID_SHARE = 12
Public Const cSIZE_GLASS_TYPE_SHARE = 14
Public Const cSIZE_PRODUCTID_SHARE = 14
Public Const cSIZE_PROCESSID_SHARE = 4
Public Const cSIZE_RECIPEID_SHARE = 40
Public Const cSIZE_SALE_ORDER_SHARE = 10
Public Const cSIZE_CF_GLASSID_SHARE = 10
Public Const cSIZE_ARRAY_LOTID_SHARE = 11
Public Const cSIZE_ARRAY_GLASSID_SHARE = 16
Public Const cSIZE_CF_GLASS_INFO_SHARE = 52
Public Const cSIZE_TFT_PANEL_JUDGE_SHARE = 2
Public Const cSIZE_PRE_PROCESSID1_SHARE = 4
Public Const cSIZE_GROUPID_SHARE = 15
Public Const cSIZE_TRANSFER_TIME_SHARE = 3

'Point Defect Type define
Public Const cDEFECT_TYPE_TB = 1
Public Const cDEFECT_TYPE_TD = 2
Public Const cDEFECT_TYPE_TT = 3

Public Const cFTP_HOST = 1
Public Const cFTP_DEFECT = 2

'EQP STEP
Public Const cSTEP_RONI = 1
Public Const cSTEP_RBBC = 10
Public Const cSTEP_RABC = 20
Public Const cSTEP_PJPG = 30
Public Const cSTEP_PSPO = 40
Public Const cSTEP_RRAL = 50

