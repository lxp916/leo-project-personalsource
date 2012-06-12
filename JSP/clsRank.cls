VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim DEFECT_DATA()                       As DEFECT_DATA_STRUCTURE
Dim DEFECT_PRIORITY(1 To 9)             As DEFECT_PRIORITY_STRUCTURE
Dim RANK_PRIORITY(1 To 30)               As RANK_PRIORITY_STRUCTURE

Dim LINE_LENGTH_TB()                    As Double
Dim LINE_LENGTH_TD()                    As Double
Dim LINE_LENGTH_TT()                    As Double

Dim POINT_DEFECT_COUNT(1 To 3)          As Integer
Dim LINE_LENGTH_COUNT(1 To 3)           As Integer

Dim TB_COUNT                            As Integer
Dim TD_COUNT                            As Integer
Dim CURRENT_PRIORITY                    As Integer
Dim GRADE_DEFECT_INDEX                  As Integer
Dim DEFECT_COUNT                        As Integer
Dim SK_PANEL_COUNT                      As Integer
Dim SK_PANEL_SETTING                    As Integer

Dim CURRENT_KEYID                       As String
Dim CURRENT_RANK                        As String
Dim CURRENT_GRADE                       As String
Dim HIGHEST_GRADE                       As String
Dim START_TIME                          As String
Dim END_TIME                            As String
Dim SELECT_DEFECTCODE                   As String

Dim TACT_TIME                           As Double
Dim TRANSFER_TIME                       As Double
Dim START_TIMER                         As Double
Dim END_TIMER                           As Double

Public Sub Init_Class()

    TB_COUNT = 0
    TD_COUNT = 0
    DEFECT_COUNT = 0
    
    CURRENT_RANK = ""
    CURRENT_GRADE = ""
'    START_TIME = ""
    END_TIME = ""
    
'==========================================================================================================
'
'  Modify Date : 2011. 12. 23
'  Modify by K.H. KIM
'  Content
'    - When rank object initialize time, rank priority also initialize too
'
'
'  Start of modify
'
'==========================================================================================================
    Call Init_RANK_Priority
    
End Sub

Public Sub Add_TB_Count(ByVal pDEFECT_COUNT As Integer)

    TB_COUNT = TB_COUNT + pDEFECT_COUNT
    
End Sub

Public Function Get_TB_Count() As Integer

    Get_TB_Count = TB_COUNT
    
End Function

Public Sub Add_TD_Count(ByVal pDEFECT_COUNT As Integer)

    TD_COUNT = TD_COUNT + pDEFECT_COUNT
    
End Sub

Public Function Get_TD_Count() As Integer

    Get_TD_Count = TD_COUNT
    
End Function

Public Sub Set_Current_Rank(ByVal pRank As String, ByVal pPriority As Integer)

    If pPriority < CURRENT_PRIORITY Then
        CURRENT_RANK = pRank
    End If
    
End Sub

Public Function Get_Current_Rank() As String

    Get_Current_Rank = CURRENT_RANK
    
End Function

Public Function Get_Current_Priority() As Integer

    Get_Current_Priority = CURRENT_PRIORITY
    
End Function

Public Sub Set_Current_Grade(ByVal pGrade As String)

    CURRENT_GRADE = pGrade
    
End Sub

Public Function Get_Current_Grade() As String

    Get_Current_Grade = CURRENT_GRADE
    
End Function

Public Sub Set_DEFECT_DATA_COUNT(ByVal pCount As Integer)

    If pCount > 0 Then
        DEFECT_COUNT = pCount
        ReDim DEFECT_DATA(DEFECT_COUNT)
    End If
    
End Sub

Public Sub Reset_DEFECT_COUNT()

    DEFECT_COUNT = 0
    
End Sub

Public Function Get_DEFECT_DATA_COUNT() As Integer

    Get_DEFECT_DATA_COUNT = DEFECT_COUNT
    
End Function

Public Sub Set_DEFECT_DATA(ByVal pIndex As Integer, ByVal pPanelID As String, ByVal pDEFECT_CODE As String, ByVal pDEFECT_NAME As String, _
                           ByVal pDETAIL_DIVISION As String, pDATA_ADDRESS() As String, pGATE_ADDRESS() As String, ByVal pCOLOR As String, _
                           ByVal pGRAY_LEVEL As Integer, ByVal pAccumulation As String)
                           
    Dim intIndex                            As Integer
    
    If (0 < pIndex) And (pIndex <= DEFECT_COUNT) Then
        With DEFECT_DATA(pIndex)
            .PANELID = pPanelID
            .DEFECT_CODE = pDEFECT_CODE
            .DEFECT_NAME = pDEFECT_NAME
            .DETAIL_DIVISION = pDETAIL_DIVISION
            .COLOR = pCOLOR
            .GRAY_LEVEL = pGRAY_LEVEL
            For intIndex = 1 To 3
                .DATA_ADDRESS(intIndex) = pDATA_ADDRESS(intIndex)
                .GATE_ADDRESS(intIndex) = pGATE_ADDRESS(intIndex)
            Next intIndex
            If IsNumeric(pAccumulation) = True Then
                .ACCUMULATION = CInt(pAccumulation)
            Else
                .ACCUMULATION = 0
            End If
        End With
    End If
    
End Sub

Public Function Get_DEFECT_DATA_by_Index(ByVal pIndex As Integer, pPanelID As String, pDEFECT_CODE As String, pDEFECT_NAME As String, pDETAIL_DIVISION As String, _
                                         pDATA_ADDRESS() As String, pGATE_ADDRESS() As String, pGrade As String, pRank As String, pCOLOR As String, _
                                         pGRAY_LEVEL As Integer, pAccumulation As Integer) As Boolean

    Dim intIndex                            As Integer
    
On Error GoTo ErrorHandler

    If (0 < pIndex) And (pIndex <= DEFECT_COUNT) Then
        With DEFECT_DATA(pIndex)
            pPanelID = .PANELID
            pDEFECT_CODE = .DEFECT_CODE
            pDEFECT_NAME = .DEFECT_NAME
            pCOLOR = .COLOR
            pGRAY_LEVEL = .GRAY_LEVEL
            For intIndex = 1 To 3
                pDATA_ADDRESS(intIndex) = .DATA_ADDRESS(intIndex)
                pGATE_ADDRESS(intIndex) = .GATE_ADDRESS(intIndex)
            Next intIndex
'            pGrade = .GRADE
'            pRANK = .RANK
            pDETAIL_DIVISION = .DETAIL_DIVISION
            pAccumulation = .ACCUMULATION
        End With
    End If
    
    Get_DEFECT_DATA_by_Index = True
    
    Exit Function
    
ErrorHandler:

    Get_DEFECT_DATA_by_Index = False
    
End Function

Public Function Get_DEFECT_DATA_by_Name(pPanelID As String, ByVal pDEFECT_CODE As String, pDEFECT_NAME As String, _
                                        pDATA_ADDRESS() As String, pGATE_ADDRESS() As String, pGrade As String, pRank As String, pCOLOR As String, _
                                        pGRAY_LEVEL As Integer, pAccumulation As Integer) As Boolean

    Dim intIndex                As Integer
    Dim intData_Index           As Integer
    
On Error GoTo ErrorHandler

    intData_Index = 0
    If DEFECT_COUNT > 0 Then
        For intIndex = 1 To DEFECT_COUNT
            With DEFECT_DATA(intIndex)
                If .DEFECT_CODE = pDEFECT_CODE Then
                    intData_Index = intIndex
                End If
            End With
        Next intIndex
        If intData_Index > 0 Then
            With DEFECT_DATA(intData_Index)
                pPanelID = .PANELID
                pDEFECT_NAME = .DEFECT_NAME
                pCOLOR = .COLOR
                pGRAY_LEVEL = .GRAY_LEVEL
                For intIndex = 1 To 3
                    pDATA_ADDRESS(intIndex) = .DATA_ADDRESS(intIndex)
                    pGATE_ADDRESS(intIndex) = .GATE_ADDRESS(intIndex)
                Next intIndex
                pGrade = .GRADE
                pRank = .Rank
                pAccumulation = .ACCUMULATION
            End With
        End If
    End If
    
    Get_DEFECT_DATA_by_Name = True
    
    Exit Function
    
ErrorHandler:

    Get_DEFECT_DATA_by_Name = False
    
End Function

Public Sub Set_DEFECT_GRADE(ByVal pDEFECT_CODE As String, pDATA_ADDRESS() As String, pGATE_ADDRESS() As String, ByVal pGrade As String)

    Dim intIndex                As Integer
    Dim intData_Index           As Integer
    
    intData_Index = 0
    If DEFECT_COUNT > 0 Then
        For intIndex = 1 To DEFECT_COUNT
            With DEFECT_DATA(intIndex)
                If .DEFECT_CODE = pDEFECT_CODE Then
                    If Mid(pDEFECT_CODE, 2, 1) = "M" Then
                        If (.DATA_ADDRESS(1) = pDATA_ADDRESS(1)) And (.DATA_ADDRESS(2) = pDATA_ADDRESS(2)) And (.DATA_ADDRESS(3) = pDATA_ADDRESS(3)) And _
                           (.GATE_ADDRESS(1) = pGATE_ADDRESS(1)) And (.GATE_ADDRESS(2) = pGATE_ADDRESS(2)) And (.GATE_ADDRESS(3) = pGATE_ADDRESS(3)) Then
                            intData_Index = intIndex
                        End If
                    Else
                        If (.DATA_ADDRESS(1) = pDATA_ADDRESS(1)) And (.GATE_ADDRESS(1) = pGATE_ADDRESS(1)) Then
                            intData_Index = intIndex
                        End If
                    End If
                End If
            End With
        Next intIndex
        If intData_Index > 0 Then
            DEFECT_DATA(intData_Index).GRADE = pGrade
        End If
    End If
    
End Sub

Public Sub Set_DEFECT_RANK(ByVal pDEFECT_CODE As String, pRank As String, pDATA_ADDRESS() As String, pGATE_ADDRESS() As String)

    Dim intIndex                As Integer
    Dim intData_Index           As Integer
    
    intData_Index = 0
    If DEFECT_COUNT > 0 Then
        For intIndex = 1 To DEFECT_COUNT
            With DEFECT_DATA(intIndex)
                If .DEFECT_CODE = pDEFECT_CODE Then
                    If Mid(pDEFECT_CODE, 2, 1) = "M" Then
                        If (.DATA_ADDRESS(1) = pDATA_ADDRESS(1)) And (.DATA_ADDRESS(2) = pDATA_ADDRESS(2)) And (.DATA_ADDRESS(3) = pDATA_ADDRESS(3)) And _
                           (.GATE_ADDRESS(1) = pGATE_ADDRESS(1)) And (.GATE_ADDRESS(2) = pGATE_ADDRESS(2)) And (.GATE_ADDRESS(3) = pGATE_ADDRESS(3)) Then
                            intData_Index = intIndex
                        End If
                    Else
                        If (.DATA_ADDRESS(1) = pDATA_ADDRESS(1)) And (.GATE_ADDRESS(1) = pGATE_ADDRESS(1)) Then
                            intData_Index = intIndex
                        End If
                    End If
                End If
            End With
        Next intIndex
        If intData_Index > 0 Then
            DEFECT_DATA(intData_Index).Rank = pRank
        End If
    End If
    
End Sub

Public Sub Set_Current_KEYID(ByVal pKEYID As String)

    CURRENT_KEYID = pKEYID
    
End Sub

Public Function Get_Current_KEYID() As String

    Get_Current_KEYID = CURRENT_KEYID
    
End Function

Public Sub Set_Highest_Grade(ByVal pGrade As String)

    HIGHEST_GRADE = pGrade
    
End Sub

Public Function Get_Highest_Grade() As String
    
    Get_Highest_Grade = HIGHEST_GRADE
    
End Function

Public Sub Init_DEFECT_PRIORITY()

    Dim intIndex                As Integer
    
    For intIndex = 1 To 9
        With DEFECT_PRIORITY(intIndex)
            .DEFECT_CODE = ""
            .DEFECT_GRADE = ""
            .DEFECT_RANK = ""
            .DEFECT_TYPE = ""
            .DEFECT_INDEX = 0
            .DEFECT_PRIORITY = 100
        End With
    Next intIndex
    
    GRADE_DEFECT_INDEX = 0
    
    DEFECT_PRIORITY(1).DEFECT_TYPE = "D"
    DEFECT_PRIORITY(2).DEFECT_TYPE = "L"
    DEFECT_PRIORITY(3).DEFECT_TYPE = "G"
    DEFECT_PRIORITY(4).DEFECT_TYPE = "M"
    DEFECT_PRIORITY(5).DEFECT_TYPE = "F"
    DEFECT_PRIORITY(6).DEFECT_TYPE = "P"
    DEFECT_PRIORITY(7).DEFECT_TYPE = "A"
    DEFECT_PRIORITY(8).DEFECT_TYPE = "C"
    DEFECT_PRIORITY(9).DEFECT_TYPE = "O"
    
End Sub

Public Function Check_DEFECT_TYPE_PRIORITY(pDefect_Rank As String) As String

    Dim intIndex                As Integer
    Dim intDefect_Index         As Integer
    Dim intSourceRankPriority   As Integer
    Dim intTargetRankPriority   As Integer
    
    intDefect_Index = 1
    For intIndex = 2 To 9
        With DEFECT_PRIORITY(intIndex)
            If DEFECT_PRIORITY(intDefect_Index).DEFECT_GRADE = "" Then
                If .DEFECT_GRADE <> "" Then
                    intDefect_Index = intIndex
                End If
            ElseIf .DEFECT_GRADE <> "" Then
                intSourceRankPriority = Get_Rank_Priority_by_Rank(DEFECT_PRIORITY(intDefect_Index).DEFECT_RANK)
                intTargetRankPriority = Get_Rank_Priority_by_Rank(.DEFECT_RANK)
                If intSourceRankPriority = intTargetRankPriority Then
                    If DEFECT_PRIORITY(intDefect_Index).DEFECT_PRIORITY > .DEFECT_PRIORITY Then
                        intDefect_Index = intIndex
                    End If
                ElseIf intSourceRankPriority < intTargetRankPriority Then
                    intDefect_Index = intIndex
                End If
            
            End If
        End With
    Next intIndex
    Check_DEFECT_TYPE_PRIORITY = DEFECT_PRIORITY(intDefect_Index).DEFECT_GRADE
    pDefect_Rank = DEFECT_PRIORITY(intDefect_Index).DEFECT_RANK
    GRADE_DEFECT_INDEX = DEFECT_PRIORITY(intDefect_Index).DEFECT_INDEX
    
End Function

Public Sub Set_DEFECT_PRIOITY(ByVal pDEFECT_TYPE As String, ByVal pPriority As Integer)

    Dim intIndex                As Integer
    
    For intIndex = 1 To 9
        If DEFECT_PRIORITY(intIndex).DEFECT_TYPE = pDEFECT_TYPE Then
            DEFECT_PRIORITY(intIndex).DEFECT_PRIORITY = pPriority
        End If
    Next intIndex
    
End Sub

Public Function Get_DEFECT_PRIORITY_by_DEFECT_TYPE(ByVal pDEFECT_TYPE As String) As Integer

    Dim intIndex                As Integer
    
    For intIndex = 1 To 9
        If DEFECT_PRIORITY(intIndex).DEFECT_TYPE = pDEFECT_TYPE Then
            Get_DEFECT_PRIORITY_by_DEFECT_TYPE = DEFECT_PRIORITY(intIndex).DEFECT_PRIORITY
        End If
    Next intIndex
    
End Function

Public Function Get_DEFECT_INDEX_by_DEFECT_TYPE(ByVal pDEFECT_TYPE As String) As Integer

    Dim intIndex                As Integer
    
    For intIndex = 1 To 9
        If DEFECT_PRIORITY(intIndex).DEFECT_TYPE = pDEFECT_TYPE Then
            Get_DEFECT_INDEX_by_DEFECT_TYPE = DEFECT_PRIORITY(intIndex).DEFECT_INDEX
        End If
    Next intIndex
    
End Function

Public Sub Get_DEFECT_PRIORITY_by_Index(ByVal pIndex As Integer, pDEFECT_TYPE As String, pDEFECT_PRIORITY As Integer)

    pDEFECT_TYPE = DEFECT_PRIORITY(pIndex).DEFECT_TYPE
    pDEFECT_PRIORITY = DEFECT_PRIORITY(pIndex).DEFECT_PRIORITY
    
End Sub

Public Sub Set_DEFECT_GRADE_by_PRIORITY(ByVal pDEFECT_TYPE As String, ByVal pGrade As String, ByVal pPriority As Integer, _
                                        ByVal pDefect_Index As Integer, ByVal pDEFECT_CODE As String, ByVal pRank As String)

    Dim intIndex                As Integer
    
    For intIndex = 1 To 9
        If DEFECT_PRIORITY(intIndex).DEFECT_TYPE = pDEFECT_TYPE Then
            DEFECT_PRIORITY(intIndex).DEFECT_GRADE = pGrade
            DEFECT_PRIORITY(intIndex).DEFECT_PRIORITY = pPriority
            DEFECT_PRIORITY(intIndex).DEFECT_INDEX = pDefect_Index
            DEFECT_PRIORITY(intIndex).DEFECT_CODE = pDEFECT_CODE
            DEFECT_PRIORITY(intIndex).DEFECT_RANK = pRank
        End If
    Next intIndex
    
End Sub

Public Function Get_DEFECT_PRIORITY_GRADE_by_Index(ByVal pIndex As Integer) As String

    Get_DEFECT_PRIORITY_GRADE_by_Index = DEFECT_PRIORITY(pIndex).DEFECT_GRADE
    
End Function

Public Function Get_DEFECT_PRIORITY_GRADE_by_DEFECT_TYPE(ByVal pDEFECT_TYPE As String) As String

    Dim intIndex                        As Integer
    
    For intIndex = 1 To 9
        If DEFECT_PRIORITY(intIndex).DEFECT_TYPE = pDEFECT_TYPE Then
            Get_DEFECT_PRIORITY_GRADE_by_DEFECT_TYPE = DEFECT_PRIORITY(intIndex).DEFECT_GRADE
        End If
    Next intIndex
    
End Function

Public Function Get_DEFECT_PRIORITY_RANK_by_Index(ByVal pIndex As Integer) As String

    Get_DEFECT_PRIORITY_RANK_by_Index = DEFECT_PRIORITY(pIndex).DEFECT_RANK
    
End Function

Public Function Get_DEFECT_PRIORITY_RANK_by_DEFECT_TYPE(ByVal pDEFECT_TYPE As String) As String

    Dim intIndex                        As Integer
    
    For intIndex = 1 To 9
        If DEFECT_PRIORITY(intIndex).DEFECT_TYPE = pDEFECT_TYPE Then
            Get_DEFECT_PRIORITY_RANK_by_DEFECT_TYPE = DEFECT_PRIORITY(intIndex).DEFECT_RANK
        End If
    Next intIndex
    
End Function

Public Function Get_GRADE_DEFECT_INDEX() As String

    Get_GRADE_DEFECT_INDEX = GRADE_DEFECT_INDEX
    
End Function

Public Function Get_GRADE_by_DEFECT_TYPE(ByVal pDEFECT_TYPE As String) As String

    Dim intIndex                As Integer
    
    For intIndex = 1 To 9
        If DEFECT_PRIORITY(intIndex).DEFECT_TYPE = pDEFECT_TYPE Then
            Get_GRADE_by_DEFECT_TYPE = DEFECT_PRIORITY(intIndex).DEFECT_GRADE
        End If
    Next intIndex
    
End Function

Public Function Get_GRADE_DEFECT_CODE() As String

    If DEFECT_COUNT > 0 Then
        Get_GRADE_DEFECT_CODE = DEFECT_DATA(GRADE_DEFECT_INDEX).DEFECT_CODE
    Else
        Get_GRADE_DEFECT_CODE = ""
    End If
    
End Function

Public Function Get_GRADE_DEFECT_NAME() As String

    If DEFECT_COUNT > 0 Then
        Get_GRADE_DEFECT_NAME = DEFECT_DATA(GRADE_DEFECT_INDEX).DEFECT_NAME
    Else
        Get_GRADE_DEFECT_NAME = ""
    End If
    
End Function

Public Function Get_GRADE_RANK() As String

    If DEFECT_COUNT > 0 Then
        Get_GRADE_RANK = DEFECT_DATA(GRADE_DEFECT_INDEX).Rank
    Else
        Get_GRADE_RANK = ""
    End If
    
End Function

Public Sub Init_RANK_Priority()
'============Leo 2012.05.22 Add Rank Level Start
Dim intLoop As Integer

For intLoop = 1 To UBound(RankLevel)
    With RANK_PRIORITY(intLoop)
        .Rank = RankLevel(intLoop - 1)
        .PRIORITY = intLoop
    End With
Next intLoop
'    With RANK_PRIORITY(1)
'        .Rank = "Y"
'        .PRIORITY = 1
'    End With
'    With RANK_PRIORITY(2)
'        .Rank = "L"
'        .PRIORITY = 2
'    End With
'    With RANK_PRIORITY(3)
'        .Rank = "K"
'        .PRIORITY = 3
'    End With
'    With RANK_PRIORITY(4)
'        .Rank = "C"
'        .PRIORITY = 4
'    End With
'    With RANK_PRIORITY(5)
'        .Rank = "S"
'        .PRIORITY = 5
'    End With
    
    '============Leo 2012.05.22 Add Rank Level End
End Sub

Public Function Get_Rank_Priority_by_Rank(ByVal pRank As String) As Integer

    Dim intIndex                As Integer
    
    Get_Rank_Priority_by_Rank = 0
    '============Leo 2012.05.22 Add Rank Level Start
'    For intIndex = 1 To 5
    For intIndex = 1 To UBound(RankLevel) + 1
    '============Leo 2012.05.22 Add Rank Level end
        If RANK_PRIORITY(intIndex).Rank = pRank Then
            Get_Rank_Priority_by_Rank = RANK_PRIORITY(intIndex).PRIORITY
        End If
    Next intIndex

End Function

Public Function Get_Rank_Priority_by_Index(ByVal pIndex As Index) As Integer

    If pIndex > 0 Then
        Get_Rank_Priority_by_Index = RANK_PRIORITY(pIndex).PRIORITY
    End If
    
End Function

Public Sub Init_Line_Length(ByVal pDEFECT_TYPE As Integer, ByVal pDEFECT_COUNT As Integer)

    Dim intIndex                        As Integer
    Dim intLine_Count                   As Integer
    
    POINT_DEFECT_COUNT(pDEFECT_TYPE) = pDEFECT_COUNT
    intLine_Count = 0
    For intIndex = 1 To pDEFECT_COUNT
        intLine_Count = intLine_Count + (intIndex - 1)
    Next intIndex
    LINE_LENGTH_COUNT(pDEFECT_TYPE) = intLine_Count
    
    Select Case pDEFECT_TYPE
    Case cDEFECT_TYPE_TB:
        ReDim LINE_LENGTH_TB(intLine_Count)
    Case cDEFECT_TYPE_TD:
        ReDim LINE_LENGTH_TD(intLine_Count)
    Case cDEFECT_TYPE_TT:
        ReDim LINE_LENGTH_TT(intLine_Count)
    End Select
    
End Sub

Public Sub Set_Line_Length(ByVal pDEFECT_TYPE As Integer, ByVal pLINE_LENGTH_COUNT As Integer, pDEFECT_LENGTH() As Double)

    Dim intIndex                        As Integer
    
    Select Case pDEFECT_TYPE
    Case cDEFECT_TYPE_TB:
        For intIndex = 1 To pLINE_LENGTH_COUNT
            LINE_LENGTH_TB(intIndex) = pDEFECT_LENGTH(intIndex)
        Next intIndex
    Case cDEFECT_TYPE_TD:
        For intIndex = 1 To pLINE_LENGTH_COUNT
            LINE_LENGTH_TD(intIndex) = pDEFECT_LENGTH(intIndex)
        Next intIndex
    Case cDEFECT_TYPE_TT:
        For intIndex = 1 To pLINE_LENGTH_COUNT
            LINE_LENGTH_TT(intIndex) = pDEFECT_LENGTH(intIndex)
        Next intIndex
    End Select
    
End Sub

Public Sub Get_Line_Length(ByVal pDEFECT_TYPE As Integer, pLINE_LENGTH() As Double)

    Dim intIndex                        As Integer
    
    ReDim pLINE_LENGTH(LINE_LENGTH_COUNT(pDEFECT_TYPE))
    Select Case pDEFECT_TYPE
    Case cDEFECT_TYPE_TB:
        For intIndex = 1 To LINE_LENGTH_COUNT(pDEFECT_TYPE)
            pLINE_LENGTH(intIndex) = LINE_LENGTH_TB(intIndex)
        Next intIndex
    Case cDEFECT_TYPE_TD:
        For intIndex = 1 To LINE_LENGTH_COUNT(pDEFECT_TYPE)
            pLINE_LENGTH(intIndex) = LINE_LENGTH_TD(intIndex)
        Next intIndex
    Case cDEFECT_TYPE_TT:
        For intIndex = 1 To LINE_LENGTH_COUNT(pDEFECT_TYPE)
            pLINE_LENGTH(intIndex) = LINE_LENGTH_TT(intIndex)
        Next intIndex
    End Select
    
End Sub

Public Sub Calculate_Point_Distance(ByVal pIndex As Integer, ByVal pW As Double, ByVal pL As Double, ByVal pB1 As Double, ByVal pB2 As Double)

    Dim arrDEFECT_TB()                  As Integer
    Dim arrDEFECT_TD()                  As Integer
    Dim arrDEFECT_TT()                  As Integer
    
    Dim intIndex                        As Integer
    Dim intSubIndex                     As Integer
    Dim intTB_Index                     As Integer
    Dim intTD_Index                     As Integer
    Dim intTT_Index                     As Integer
    Dim intTT_Count                     As Integer
    
    Dim lngX1                           As Long
    Dim lngX2                           As Long
    Dim lngY1                           As Long
    Dim lngY2                           As Long
    
    Dim dblDistance                     As Double
    
    intTB_Index = 0
    intTD_Index = 0
    intTT_Index = 0
    intTT_Count = 0
    ReDim arrDEFECT_TB(TB_COUNT)
    ReDim arrDEFECT_TD(TD_COUNT)
    ReDim arrDEFECT_TT(TB_COUNT + TD_COUNT)
    Call Init_Line_Length(cDEFECT_TYPE_TB, TB_COUNT)
    Call Init_Line_Length(cDEFECT_TYPE_TD, TD_COUNT)
    Call Init_Line_Length(cDEFECT_TYPE_TT, TB_COUNT + TD_COUNT)
    
    For intIndex = 1 To DEFECT_COUNT - 6
        With DEFECT_DATA(intIndex)
            If .DETAIL_DIVISION = "B" Then
                intTB_Index = intTB_Index + 1
                arrDEFECT_TB(intTB_Index) = intIndex
                
                intTT_Index = intTT_Index + 1
                arrDEFECT_TT(intTT_Index) = intIndex
            ElseIf .DETAIL_DIVISION = "D" Then
                intTD_Index = intTD_Index + 1
                arrDEFECT_TD(intTD_Index) = intIndex
                
                intTT_Index = intTT_Index + 1
                arrDEFECT_TT(intTT_Index) = intIndex
            End If
        End With
    Next intIndex
    
    intTB_Index = 0
    For intIndex = 1 To TB_COUNT - 1
        For intSubIndex = intIndex + 1 To TB_COUNT
            If Trim(DEFECT_DATA(arrDEFECT_TB(intIndex)).DATA_ADDRESS(1)) <> "" And Trim(DEFECT_DATA(arrDEFECT_TB(intSubIndex)).DATA_ADDRESS(1)) <> "" And _
               Trim(DEFECT_DATA(arrDEFECT_TB(intIndex)).GATE_ADDRESS(1)) <> "" And Trim(DEFECT_DATA(arrDEFECT_TB(intSubIndex)).GATE_ADDRESS(1)) <> "" Then
                lngX1 = CLng(DEFECT_DATA(arrDEFECT_TB(intIndex)).DATA_ADDRESS(pIndex))
                lngX2 = CLng(DEFECT_DATA(arrDEFECT_TB(intSubIndex)).DATA_ADDRESS(pIndex))
                lngY1 = CLng(DEFECT_DATA(arrDEFECT_TB(intIndex)).GATE_ADDRESS(pIndex))
                lngY2 = CLng(DEFECT_DATA(arrDEFECT_TB(intSubIndex)).GATE_ADDRESS(pIndex))
                dblDistance = Sqr(((lngX1 - lngX2) * (pW + pB1)) ^ 2 + ((lngY1 - lngY2) * (pL + pB2)) ^ 2)
                intTB_Index = intTB_Index + 1
                LINE_LENGTH_TB(intTB_Index) = dblDistance
                dblDistance = 0
            End If
        Next intSubIndex
    Next intIndex
    
    intTD_Index = 0
    For intIndex = 1 To TD_COUNT - 1
        For intSubIndex = intIndex + 1 To TD_COUNT
            If Trim(DEFECT_DATA(arrDEFECT_TD(intIndex)).DATA_ADDRESS(1)) <> "" And Trim(DEFECT_DATA(arrDEFECT_TD(intSubIndex)).DATA_ADDRESS(1)) <> "" And _
               Trim(DEFECT_DATA(arrDEFECT_TD(intIndex)).GATE_ADDRESS(1)) <> "" And Trim(DEFECT_DATA(arrDEFECT_TD(intSubIndex)).GATE_ADDRESS(1)) <> "" Then
                lngX1 = CLng(DEFECT_DATA(arrDEFECT_TD(intIndex)).DATA_ADDRESS(pIndex))
                lngX2 = CLng(DEFECT_DATA(arrDEFECT_TD(intSubIndex)).DATA_ADDRESS(pIndex))
                lngY1 = CLng(DEFECT_DATA(arrDEFECT_TD(intIndex)).GATE_ADDRESS(pIndex))
                lngY2 = CLng(DEFECT_DATA(arrDEFECT_TD(intSubIndex)).GATE_ADDRESS(pIndex))
                dblDistance = Sqr(((lngX1 - lngX2) * (pW + pB1)) ^ 2 + ((lngY1 - lngY2) * (pL + pB2)) ^ 2)
                intTD_Index = intTD_Index + 1
                LINE_LENGTH_TD(intTD_Index) = dblDistance
                dblDistance = 0
            End If
        Next intSubIndex
    Next intIndex
    
    intTT_Count = TB_COUNT + TD_COUNT
    intTT_Index = 0
    If intTT_Count < 6 Then
        intTT_Index = 0
        For intIndex = 1 To intTT_Count - 1
            For intSubIndex = intIndex + 1 To intTT_Count
                If Trim(DEFECT_DATA(arrDEFECT_TT(intIndex)).DATA_ADDRESS(1)) <> "" And Trim(DEFECT_DATA(arrDEFECT_TT(intSubIndex)).DATA_ADDRESS(1)) <> "" And _
                   Trim(DEFECT_DATA(arrDEFECT_TT(intIndex)).GATE_ADDRESS(1)) <> "" And Trim(DEFECT_DATA(arrDEFECT_TT(intSubIndex)).GATE_ADDRESS(1)) <> "" Then
                    lngX1 = CLng(DEFECT_DATA(arrDEFECT_TT(intIndex)).DATA_ADDRESS(pIndex))
                    lngX2 = CLng(DEFECT_DATA(arrDEFECT_TT(intSubIndex)).DATA_ADDRESS(pIndex))
                    lngY1 = CLng(DEFECT_DATA(arrDEFECT_TT(intIndex)).GATE_ADDRESS(pIndex))
                    lngY2 = CLng(DEFECT_DATA(arrDEFECT_TT(intSubIndex)).GATE_ADDRESS(pIndex))
                    dblDistance = Sqr(((lngX1 - lngX2) * (pW + pB1)) ^ 2 + ((lngY1 - lngY2) * (pL + pB2)) ^ 2)
                    intTT_Index = intTT_Index + 1
                    LINE_LENGTH_TT(intTT_Index) = dblDistance
                    dblDistance = 0
                End If
            Next intSubIndex
        Next intIndex
    End If
    
End Sub

Public Sub Get_Minimum_Distance(pTB_Distance As Double, pTD_Distance As Double, pTT_Distance As Double)

    Dim dblDistance                     As Double
    
    Dim intIndex                        As Integer
    
    If LINE_LENGTH_COUNT(cDEFECT_TYPE_TB) > 0 Then
        dblDistance = LINE_LENGTH_TB(1)
        For intIndex = 2 To LINE_LENGTH_COUNT(cDEFECT_TYPE_TB)
            If LINE_LENGTH_TB(intIndex) <> 0 Then
                If LINE_LENGTH_TB(intIndex) < dblDistance Then
                    dblDistance = LINE_LENGTH_TB(intIndex)
                End If
            End If
        Next intIndex
        pTB_Distance = dblDistance
    End If
    
    If LINE_LENGTH_COUNT(cDEFECT_TYPE_TD) > 0 Then
        dblDistance = LINE_LENGTH_TD(1)
        For intIndex = 2 To LINE_LENGTH_COUNT(cDEFECT_TYPE_TD)
            If LINE_LENGTH_TD(intIndex) <> 0 Then
                If LINE_LENGTH_TD(intIndex) < dblDistance Then
                    dblDistance = LINE_LENGTH_TD(intIndex)
                End If
            End If
        Next intIndex
        pTD_Distance = dblDistance
    End If
    
    If LINE_LENGTH_COUNT(cDEFECT_TYPE_TT) > 0 Then
        dblDistance = LINE_LENGTH_TT(1)
        For intIndex = 2 To LINE_LENGTH_COUNT(cDEFECT_TYPE_TT)
            If LINE_LENGTH_TT(intIndex) <> 0 Then
                If LINE_LENGTH_TT(intIndex) < dblDistance Then
                    dblDistance = LINE_LENGTH_TT(intIndex)
                End If
            End If
        Next intIndex
        pTT_Distance = dblDistance
    End If
    
End Sub

Public Sub Get_Point_Distance(pDISTANCE_TB() As Double, pDISTANCE_TD() As Double, pDISTANCE_TT() As Double)

    Dim intIndex                        As Integer
    
    ReDim pDISTANCE_TB(LINE_LENGTH_COUNT(cDEFECT_TYPE_TB))
    ReDim pDISTANCE_TD(LINE_LENGTH_COUNT(cDEFECT_TYPE_TD))
    ReDim pDISTANCE_TT(LINE_LENGTH_COUNT(cDEFECT_TYPE_TT))
    
    For intIndex = 1 To LINE_LENGTH_COUNT(cDEFECT_TYPE_TB)
        pDISTANCE_TB(intIndex) = LINE_LENGTH_TB(intIndex)
    Next intIndex
    
    For intIndex = 1 To LINE_LENGTH_COUNT(cDEFECT_TYPE_TD)
        pDISTANCE_TD(intIndex) = LINE_LENGTH_TD(intIndex)
    Next intIndex
    
    For intIndex = 1 To LINE_LENGTH_COUNT(cDEFECT_TYPE_TT)
        pDISTANCE_TT(intIndex) = LINE_LENGTH_TT(intIndex)
    Next intIndex
    
End Sub

Public Sub Set_START_TIME(ByVal pTime As String)

    START_TIME = pTime
    START_TIMER = Timer
    TRANSFER_TIME = START_TIMER - END_TIMER
    
End Sub

Public Function Get_START_TIME() As String

    Get_START_TIME = START_TIME
    
End Function

Public Sub Set_END_TIME(ByVal pTime As String)

    END_TIME = pTime
    END_TIMER = Timer
    TACT_TIME = END_TIMER - START_TIMER
    
End Sub

Public Function Get_END_TIME() As String

    Get_END_TIME = END_TIME
    
End Function

Public Sub Init_Timer()

    START_TIMER = Timer
    END_TIMER = Timer
    
End Sub

Public Function Get_Tact_Time() As Double

    Get_Tact_Time = TACT_TIME
    
End Function

Public Function Get_Transfer_Time() As Double

    Get_Transfer_Time = TRANSFER_TIME
    
End Function

Public Sub Reset_SK_SETTING()

    SK_PANEL_COUNT = 0
    SK_PANEL_SETTING = 0
    
End Sub

Public Function Increase_SK_Panel_Count() As Boolean

    SK_PANEL_COUNT = SK_PANEL_COUNT + 1
    If SK_PANEL_SETTING <> 0 Then
        If SK_PANEL_COUNT >= SK_PANEL_SETTING Then
            Increase_SK_Panel_Count = True
            SK_PANEL_COUNT = 0
        Else
            Increase_SK_Panel_Count = False
        End If
    Else
        Increase_SK_Panel_Count = False
    End If
    
End Function

Public Sub Set_Select_DEFECTCODE(ByVal pDEFECTCODE As String)

    SELECT_DEFECTCODE = pDEFECTCODE
    
End Sub

Public Function Get_Select_DEFECTCODE() As String

    Get_Select_DEFECTCODE = SELECT_DEFECTCODE
    
End Function
Public Function Get_Rank_Levels() As String()
'============Leo 2012.05.22 Add Rank Level Start
    Dim dbMyDB                              As Database
    Dim lstRecord                           As Recordset
    Dim intCount                         As Integer
    Dim strQuery                              As String
    Dim strDB_Path                          As String
    Dim strDB_FileName                      As String

    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Parameter.mdb"
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        strQuery = "SELECT * FROM Rank_Level order by RankLevel"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
         intCount = lstRecord.RecordCount
        If intCount > 0 Then
            ReDim RankLevel(intCount - 1)
            For intCount = 0 To lstRecord.RecordCount - 1
            RankLevel(intCount) = lstRecord("RankCode")
            lstRecord.MoveNext
            Next intCount
        End If
        
        lstRecord.Close
    End If
   ' Set RankLevel = Rank
    '============Leo 2012.05.22 Add Rank Level End
    Get_Rank_Levels = RankLevel
    
End Function
