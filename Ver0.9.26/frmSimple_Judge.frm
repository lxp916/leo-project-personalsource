VERSION 5.00
Begin VB.Form frmSimple_Judge 
   Caption         =   "Simple Judge"
   ClientHeight    =   1680
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   ScaleHeight     =   1680
   ScaleWidth      =   7275
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton cmdC 
      Caption         =   "C"
      Height          =   525
      Left            =   5400
      TabIndex        =   3
      Top             =   480
      Width           =   1245
   End
   Begin VB.CommandButton cmdNG 
      Caption         =   "K"
      Height          =   525
      Left            =   3870
      TabIndex        =   2
      Top             =   480
      Width           =   1245
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "L"
      Height          =   525
      Left            =   2250
      TabIndex        =   1
      Top             =   480
      Width           =   1245
   End
   Begin VB.CommandButton cmdY 
      Caption         =   "Y"
      Height          =   525
      Left            =   690
      TabIndex        =   0
      Top             =   480
      Width           =   1245
   End
End
Attribute VB_Name = "frmSimple_Judge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdC_Click()

    Dim dbMyDB                              As Database
    
    Dim typGRADE_DATA()                     As GRADE_DATA_STRUCTURE
    Dim typDEFECT_DATA()                    As DEFECT_DATA_STRUCTURE
    
    Dim typRANK_DATA                        As RANK_DATA_STRUCTURE
    Dim typGRADE_DEFECT                     As DEFECT_DATA_STRUCTURE
    Dim typPFCD_ADDRESS_DATA                As PFCD_ADDRESS_STRUCTURE
    
    Dim arrPOINT_DEFECT_COUNT(1 To 3)       As Integer
    
    Dim strDB_Path                          As String
    Dim strDB_FileName                      As String
    Dim strQuery                            As String
    Dim strNew_Grade                        As String
    Dim strPoint_Defect_Rank                As String
    Dim strGrade                            As String
    Dim strRank                             As String
    Dim strDEFECT_TYPE                      As String
    Dim strState                            As String
    
    Dim intPortNo                           As Integer
    Dim intRow                              As Integer
    Dim intCol                              As Integer
    Dim intIndex                            As Integer
    Dim intDefect_Count                     As Integer
    Dim intGrade_Defect_Index               As Integer
    Dim intGrade_Count                      As Integer
    Dim intPoint_Defect_Total               As Integer
    Dim intPTN_Index                        As Integer
    
    Call ENV.Get_Device_Data_by_Name("API", intPortNo, strState)

    If Left(frmMain.flxEQ_Information.TextMatrix(5, 1), 5) = "CATST" Then
        If intPortNo > 0 Then
            Call QUEUE.Put_Send_Command(intPortNo, "QBLV")
        End If
    End If

    Call RANK_OBJ.Set_DEFECT_DATA_COUNT(frmJudge.flxDefect_List.Rows - 1)
    intDefect_Count = frmJudge.flxDefect_List.Rows - 1
    ReDim typDEFECT_DATA(intDefect_Count)

    With frmJudge.flxDefect_List
        If intDefect_Count > 0 Then
            For intRow = 1 To intDefect_Count
                .TextMatrix(intRow, 8) = pubPANEL_INFO.PANELID
                typDEFECT_DATA(intRow).DEFECT_CODE = .TextMatrix(intRow, 0)
                typDEFECT_DATA(intRow).DEFECT_NAME = .TextMatrix(intRow, 1)
                typDEFECT_DATA(intRow).PANELID = .TextMatrix(intRow, 8)
                intIndex = 0
                For intCol = 2 To 7 Step 2
                    intIndex = intIndex + 1
                    typDEFECT_DATA(intRow).DATA_ADDRESS(intIndex) = .TextMatrix(intRow, intCol)
                    typDEFECT_DATA(intRow).GATE_ADDRESS(intIndex) = .TextMatrix(intRow, intCol + 1)
                Next intCol
                If .TextMatrix(intRow, 13) = "" Then
                    .TextMatrix(intRow, 13) = "0"
                End If
                Call RANK_OBJ.Set_DEFECT_DATA(intRow, pubPANEL_INFO.PANELID, .TextMatrix(intRow, 0), .TextMatrix(intRow, 1), .TextMatrix(intRow, 11), typDEFECT_DATA(intRow).DATA_ADDRESS, typDEFECT_DATA(intRow).GATE_ADDRESS, .TextMatrix(intRow, 12), CInt(.TextMatrix(intRow, 13)))
                If (frmJudge.flxDefect_List.TextMatrix(intRow, 9) = "") And (frmJudge.flxDefect_List.TextMatrix(intRow, 10) = "") Then
                    Call ACCUMULATE(pubCST_INFO, typDEFECT_DATA(intRow), intRow)
                Else
                    Call Get_Rank_Data(pubCST_INFO.PROCESS_NUM, typRANK_DATA, typGRADE_DATA, typDEFECT_DATA(intRow).DEFECT_CODE, intGrade_Count)
                    typDEFECT_DATA(intRow).PRIORITY = typRANK_DATA.PRIORITY
                    If typRANK_DATA.ACCUMULATION <> "X" Then
                        'Accumulation
                        Call Add_Point_Defect_Total(typDEFECT_DATA(intRow), CInt(frmJudge.flxDefect_List.TextMatrix(intRow, 10)))
                        intPoint_Defect_Total = Get_Point_Defect_Total(typDEFECT_DATA(intRow).DEFECT_CODE, typDEFECT_DATA(intRow).PANELID)

                        If typRANK_DATA.DETAIL_DIVISION = "B" Then
                            Call RANK_OBJ.Add_TB_Count(CInt(frmJudge.flxDefect_List.TextMatrix(intRow, 10)))
                        Else
                            Call RANK_OBJ.Add_TD_Count(CInt(frmJudge.flxDefect_List.TextMatrix(intRow, 10)))
                        End If
                    End If
                    If typRANK_DATA.JUDGE_OR_NOT = "O" Then
                        strGrade = ""
                        strRank = frmJudge.flxDefect_List.TextMatrix(intRow, 9)
                        For intIndex = 1 To intGrade_Count
                            If (strGrade = "") And (typDEFECT_DATA(intRow).DEFECT_CODE = typGRADE_DATA(intIndex).DEFECT_CODE) And (InStr(typGRADE_DATA(intIndex).RANK, strRank) > 0) Then
                                strGrade = typGRADE_DATA(intIndex).GRADE
                            End If
                        Next intIndex
                        typDEFECT_DATA(intRow).GRADE = strGrade
                        Call SaveLog("cmdGrade_Click", typDEFECT_DATA(intRow).DEFECT_CODE & "'s RANK : " & typDEFECT_DATA(intRow).RANK & ", GRADE : " & strGrade)
                        Call RANK_OBJ.Set_DEFECT_RANK(typDEFECT_DATA(intRow).DEFECT_CODE, strRank, typDEFECT_DATA(intRow).DATA_ADDRESS, typDEFECT_DATA(intRow).GATE_ADDRESS)
                        Call RANK_OBJ.Set_DEFECT_GRADE(typDEFECT_DATA(intRow).DEFECT_CODE, typDEFECT_DATA(intRow).DATA_ADDRESS, typDEFECT_DATA(intRow).GATE_ADDRESS, strGrade)
                    End If
                End If
            Next intRow

            Call RANK_OBJ.Init_DEFECT_PRIORITY
            For intIndex = 1 To 3
                arrPOINT_DEFECT_COUNT(intIndex) = 0
            Next intIndex
            For intIndex = 1 To intDefect_Count
                With typDEFECT_DATA(intIndex)
                    Call RANK_OBJ.Get_DEFECT_DATA_by_Index(intIndex, .PANELID, .DEFECT_CODE, .DEFECT_NAME, .DETAIL_DIVISION, .DATA_ADDRESS, .GATE_ADDRESS, .GRADE, .RANK, .COLOR, .GRAY_LEVEL)
                    strDEFECT_TYPE = Mid(.DEFECT_CODE, 2, 1)
                    If typDEFECT_DATA(intIndex).PRIORITY < RANK_OBJ.Get_DEFECT_PRIORITY_by_DEFECT_TYPE(strDEFECT_TYPE) Then
                        Call RANK_OBJ.Set_DEFECT_GRADE_by_PRIORITY(strDEFECT_TYPE, .GRADE, .PRIORITY, intIndex, .DEFECT_CODE, .RANK)
                    ElseIf typDEFECT_DATA(intIndex).PRIORITY = RANK_OBJ.Get_DEFECT_PRIORITY_by_DEFECT_TYPE(strDEFECT_TYPE) Then
                        If RANK_OBJ.Get_Rank_Priority_by_Rank(typDEFECT_DATA(intIndex).RANK) > RANK_OBJ.Get_Rank_Priority_by_Rank(strDEFECT_TYPE) Then
                            Call RANK_OBJ.Set_DEFECT_GRADE_by_PRIORITY(strDEFECT_TYPE, .GRADE, .PRIORITY, intIndex, .DEFECT_CODE, .RANK)
                        End If
                    End If
                    'Point Defect Count Check
                    Select Case .DETAIL_DIVISION
                    Case "B":
                        arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TB) = arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TB) + 1
                    Case "D":
                        arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TD) = arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TD) + 1
                    End Select
                    arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TT) = arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TB) + arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TD)
                End With
            Next intIndex

            'Line Lengh calcaulation between each point defect
            For intIndex = 1 To 3
                If arrPOINT_DEFECT_COUNT(intIndex) < 6 Then
                    'Reference W, L, B1 and B2 values in PFCD_Address.csv file
                    Call Get_PFCD_ADDRESS_DATA(typPFCD_ADDRESS_DATA, pubPANEL_INFO.PRODUCTID, CInt(Right(pubPANEL_INFO.PANELID, 2)))
                    With typPFCD_ADDRESS_DATA
                        Call RANK_OBJ.Calculate_Point_Distance(intIndex, .W, .L, .B1, .B2)
                    End With
                End If
            Next intIndex

            If strPoint_Defect_Rank = "" Then
                strPoint_Defect_Rank = pubPANEL_INFO.TFT_REPAIR_GRADE
            End If

            strNew_Grade = Get_Panel_Grade(strPoint_Defect_Rank)
            intGrade_Defect_Index = RANK_OBJ.Get_GRADE_DEFECT_INDEX
            strNew_Grade = PreJudgeGradeChange1(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strNew_Grade = PreJudgeGradeChange2(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA, intDefect_Count)
            strNew_Grade = PreJudgeGradeChange3(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index), strPoint_Defect_Rank)
            strNew_Grade = PostJudgeOtherRule1(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strNew_Grade = PostJudgeOtherRule2(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strNew_Grade = PostJudgeOtherRule3(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strNew_Grade = PostJudgeGradeChange1(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strNew_Grade = PostJudgeGradeChange2(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strNew_Grade = CheckPanelIDChangeGrade(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strNew_Grade = ChangeGrade(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strNew_Grade = ChangeGradeByDefectCode(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strNew_Grade = RepairPointTimes(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strNew_Grade = FlagChangeGrade(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index), pubJOB_INFO)
            strNew_Grade = SKChange(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
        Else
            'Get highst grade from Rank table
            strNew_Grade = RANK_OBJ.Get_Highest_Grade
        End If
    End With

    strNew_Grade = Me.cmdNG.Caption
    frmMain.lblPost_Judge.Caption = strNew_Grade
    With frmMain.flxJudge_History
        intRow = .Rows - 1
        If frmJudge.flxDefect_List.Rows > 1 Then
            .TextMatrix(intRow, 3) = strNew_Grade
            .TextMatrix(intRow, 4) = frmJudge.flxDefect_List.TextMatrix(frmJudge.flxDefect_List.Rows - 1, 0) 'typDEFECT_DATA(intGrade_Defect_Index).DEFECT_CODE
            .TextMatrix(intRow, 6) = Format(TIME, "HH:MM:SS")
        Else
            .TextMatrix(intRow, 3) = strNew_Grade
            .TextMatrix(intRow, 4) = ""
            .TextMatrix(intRow, 6) = Format(TIME, "HH:MM:SS")
        End If
    End With
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Result.mdb"

    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)

        strQuery = "UPDATE PANEL_DATA SET "
        strQuery = strQuery & "PANEL_GRADE='" & strNew_Grade & "', "
        strQuery = strQuery & "PANEL_LOSSCODE='" & frmMain.flxJudge_History.TextMatrix(intRow, 4) & "' WHERE "
        strQuery = strQuery & "KEYID='" & RANK_OBJ.Get_Current_KEYID & "'"

        dbMyDB.Execute (strQuery)

        dbMyDB.Close
    End If

    Call Send_Panel_Judge(pubPANEL_INFO.PANELID, strNew_Grade, frmMain.flxJudge_History.TextMatrix(intRow, 4), "")

    intPTN_Index = CInt(frmJudge.lblCurrent_PTN_Index.Caption)
    EQP.Set_PATTERN_END_by_Index (intPTN_Index)
    intPortNo = EQP.Get_PG_PortID
    Call QUEUE.Put_Send_Command(intPortNo, "QPPF")
    
    Unload Me

End Sub

Private Sub cmdNG_Click()

    Dim dbMyDB                              As Database
    
    Dim typGRADE_DATA()                     As GRADE_DATA_STRUCTURE
    Dim typDEFECT_DATA()                    As DEFECT_DATA_STRUCTURE
    
    Dim typRANK_DATA                        As RANK_DATA_STRUCTURE
    Dim typGRADE_DEFECT                     As DEFECT_DATA_STRUCTURE
    Dim typPFCD_ADDRESS_DATA                As PFCD_ADDRESS_STRUCTURE
    
    Dim arrPOINT_DEFECT_COUNT(1 To 3)       As Integer
    
    Dim strDB_Path                          As String
    Dim strDB_FileName                      As String
    Dim strQuery                            As String
    Dim strNew_Grade                        As String
    Dim strPoint_Defect_Rank                As String
    Dim strGrade                            As String
    Dim strRank                             As String
    Dim strDEFECT_TYPE                      As String
    Dim strState                            As String
    
    Dim intPortNo                           As Integer
    Dim intRow                              As Integer
    Dim intCol                              As Integer
    Dim intIndex                            As Integer
    Dim intDefect_Count                     As Integer
    Dim intGrade_Defect_Index               As Integer
    Dim intGrade_Count                      As Integer
    Dim intPoint_Defect_Total               As Integer
    Dim intPTN_Index                        As Integer
    
    Call ENV.Get_Device_Data_by_Name("API", intPortNo, strState)

    If Left(frmMain.flxEQ_Information.TextMatrix(5, 1), 5) = "CATST" Then
        If intPortNo > 0 Then
            Call QUEUE.Put_Send_Command(intPortNo, "QBLV")
        End If
    End If

    Call RANK_OBJ.Set_DEFECT_DATA_COUNT(frmJudge.flxDefect_List.Rows - 1)
    intDefect_Count = frmJudge.flxDefect_List.Rows - 1
    ReDim typDEFECT_DATA(intDefect_Count)

    With frmJudge.flxDefect_List
        If intDefect_Count > 0 Then
            For intRow = 1 To intDefect_Count
                .TextMatrix(intRow, 8) = pubPANEL_INFO.PANELID
                typDEFECT_DATA(intRow).DEFECT_CODE = .TextMatrix(intRow, 0)
                typDEFECT_DATA(intRow).DEFECT_NAME = .TextMatrix(intRow, 1)
                typDEFECT_DATA(intRow).PANELID = .TextMatrix(intRow, 8)
                intIndex = 0
                For intCol = 2 To 7 Step 2
                    intIndex = intIndex + 1
                    typDEFECT_DATA(intRow).DATA_ADDRESS(intIndex) = .TextMatrix(intRow, intCol)
                    typDEFECT_DATA(intRow).GATE_ADDRESS(intIndex) = .TextMatrix(intRow, intCol + 1)
                Next intCol
                If .TextMatrix(intRow, 13) = "" Then
                    .TextMatrix(intRow, 13) = "0"
                End If
                Call RANK_OBJ.Set_DEFECT_DATA(intRow, pubPANEL_INFO.PANELID, .TextMatrix(intRow, 0), .TextMatrix(intRow, 1), .TextMatrix(intRow, 11), typDEFECT_DATA(intRow).DATA_ADDRESS, typDEFECT_DATA(intRow).GATE_ADDRESS, .TextMatrix(intRow, 12), CInt(.TextMatrix(intRow, 13)))
                If (frmJudge.flxDefect_List.TextMatrix(intRow, 9) = "") And (frmJudge.flxDefect_List.TextMatrix(intRow, 10) = "") Then
                    Call ACCUMULATE(pubCST_INFO, typDEFECT_DATA(intRow), intRow)
                Else
                    Call Get_Rank_Data(pubCST_INFO.PROCESS_NUM, typRANK_DATA, typGRADE_DATA, typDEFECT_DATA(intRow).DEFECT_CODE, intGrade_Count)
                    typDEFECT_DATA(intRow).PRIORITY = typRANK_DATA.PRIORITY
                    If typRANK_DATA.ACCUMULATION <> "X" Then
                        'Accumulation
                        Call Add_Point_Defect_Total(typDEFECT_DATA(intRow), CInt(frmJudge.flxDefect_List.TextMatrix(intRow, 10)))
                        intPoint_Defect_Total = Get_Point_Defect_Total(typDEFECT_DATA(intRow).DEFECT_CODE, typDEFECT_DATA(intRow).PANELID)

                        If typRANK_DATA.DETAIL_DIVISION = "B" Then
                            Call RANK_OBJ.Add_TB_Count(CInt(frmJudge.flxDefect_List.TextMatrix(intRow, 10)))
                        Else
                            Call RANK_OBJ.Add_TD_Count(CInt(frmJudge.flxDefect_List.TextMatrix(intRow, 10)))
                        End If
                    End If
                    If typRANK_DATA.JUDGE_OR_NOT = "O" Then
                        strGrade = ""
                        strRank = frmJudge.flxDefect_List.TextMatrix(intRow, 9)
                        For intIndex = 1 To intGrade_Count
                            If (strGrade = "") And (typDEFECT_DATA(intRow).DEFECT_CODE = typGRADE_DATA(intIndex).DEFECT_CODE) And (InStr(typGRADE_DATA(intIndex).RANK, strRank) > 0) Then
                                strGrade = typGRADE_DATA(intIndex).GRADE
                            End If
                        Next intIndex
                        typDEFECT_DATA(intRow).GRADE = strGrade
                        Call SaveLog("cmdGrade_Click", typDEFECT_DATA(intRow).DEFECT_CODE & "'s RANK : " & typDEFECT_DATA(intRow).RANK & ", GRADE : " & strGrade)
                        Call RANK_OBJ.Set_DEFECT_RANK(typDEFECT_DATA(intRow).DEFECT_CODE, strRank, typDEFECT_DATA(intRow).DATA_ADDRESS, typDEFECT_DATA(intRow).GATE_ADDRESS)
                        Call RANK_OBJ.Set_DEFECT_GRADE(typDEFECT_DATA(intRow).DEFECT_CODE, typDEFECT_DATA(intRow).DATA_ADDRESS, typDEFECT_DATA(intRow).GATE_ADDRESS, strGrade)
                    End If
                End If
            Next intRow

            Call RANK_OBJ.Init_DEFECT_PRIORITY
            For intIndex = 1 To 3
                arrPOINT_DEFECT_COUNT(intIndex) = 0
            Next intIndex
            For intIndex = 1 To intDefect_Count
                With typDEFECT_DATA(intIndex)
                    Call RANK_OBJ.Get_DEFECT_DATA_by_Index(intIndex, .PANELID, .DEFECT_CODE, .DEFECT_NAME, .DETAIL_DIVISION, .DATA_ADDRESS, .GATE_ADDRESS, .GRADE, .RANK, .COLOR, .GRAY_LEVEL)
                    strDEFECT_TYPE = Mid(.DEFECT_CODE, 2, 1)
                    If typDEFECT_DATA(intIndex).PRIORITY < RANK_OBJ.Get_DEFECT_PRIORITY_by_DEFECT_TYPE(strDEFECT_TYPE) Then
                        Call RANK_OBJ.Set_DEFECT_GRADE_by_PRIORITY(strDEFECT_TYPE, .GRADE, .PRIORITY, intIndex, .DEFECT_CODE, .RANK)
                    ElseIf typDEFECT_DATA(intIndex).PRIORITY = RANK_OBJ.Get_DEFECT_PRIORITY_by_DEFECT_TYPE(strDEFECT_TYPE) Then
                        If RANK_OBJ.Get_Rank_Priority_by_Rank(typDEFECT_DATA(intIndex).RANK) > RANK_OBJ.Get_Rank_Priority_by_Rank(strDEFECT_TYPE) Then
                            Call RANK_OBJ.Set_DEFECT_GRADE_by_PRIORITY(strDEFECT_TYPE, .GRADE, .PRIORITY, intIndex, .DEFECT_CODE, .RANK)
                        End If
                    End If
                    'Point Defect Count Check
                    Select Case .DETAIL_DIVISION
                    Case "B":
                        arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TB) = arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TB) + 1
                    Case "D":
                        arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TD) = arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TD) + 1
                    End Select
                    arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TT) = arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TB) + arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TD)
                End With
            Next intIndex

            'Line Lengh calcaulation between each point defect
            For intIndex = 1 To 3
                If arrPOINT_DEFECT_COUNT(intIndex) < 6 Then
                    'Reference W, L, B1 and B2 values in PFCD_Address.csv file
                    Call Get_PFCD_ADDRESS_DATA(typPFCD_ADDRESS_DATA, pubPANEL_INFO.PRODUCTID, CInt(Right(pubPANEL_INFO.PANELID, 2)))
                    With typPFCD_ADDRESS_DATA
                        Call RANK_OBJ.Calculate_Point_Distance(intIndex, .W, .L, .B1, .B2)
                    End With
                End If
            Next intIndex

            If strPoint_Defect_Rank = "" Then
                strPoint_Defect_Rank = pubPANEL_INFO.TFT_REPAIR_GRADE
            End If

            strNew_Grade = Get_Panel_Grade(strPoint_Defect_Rank)
            intGrade_Defect_Index = RANK_OBJ.Get_GRADE_DEFECT_INDEX
            strNew_Grade = PreJudgeGradeChange1(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strNew_Grade = PreJudgeGradeChange2(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA, intDefect_Count)
            strNew_Grade = PreJudgeGradeChange3(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index), strPoint_Defect_Rank)
            strNew_Grade = PostJudgeOtherRule1(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strNew_Grade = PostJudgeOtherRule2(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strNew_Grade = PostJudgeOtherRule3(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strNew_Grade = PostJudgeGradeChange1(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strNew_Grade = PostJudgeGradeChange2(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strNew_Grade = CheckPanelIDChangeGrade(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strNew_Grade = ChangeGrade(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strNew_Grade = ChangeGradeByDefectCode(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strNew_Grade = RepairPointTimes(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strNew_Grade = FlagChangeGrade(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index), pubJOB_INFO)
            strNew_Grade = SKChange(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
        Else
            'Get highst grade from Rank table
            strNew_Grade = RANK_OBJ.Get_Highest_Grade
        End If
    End With

    strNew_Grade = Me.cmdNG.Caption
    frmMain.lblPost_Judge.Caption = strNew_Grade
    With frmMain.flxJudge_History
        intRow = .Rows - 1
        If frmJudge.flxDefect_List.Rows > 1 Then
            .TextMatrix(intRow, 4) = strNew_Grade
            .TextMatrix(intRow, 5) = frmJudge.flxDefect_List.TextMatrix(frmJudge.flxDefect_List.Rows - 1, 0) 'typDEFECT_DATA(intGrade_Defect_Index).DEFECT_CODE
            .TextMatrix(intRow, 6) = Format(TIME, "HH:MM:SS")
        Else
            .TextMatrix(intRow, 4) = strNew_Grade
            .TextMatrix(intRow, 5) = ""
            .TextMatrix(intRow, 6) = Format(TIME, "HH:MM:SS")
        End If
    End With
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Result.mdb"

    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)

        strQuery = "UPDATE PANEL_DATA SET "
        strQuery = strQuery & "PANEL_GRADE='" & strNew_Grade & "', "
        strQuery = strQuery & "PANEL_LOSSCODE='" & frmMain.flxJudge_History.TextMatrix(intRow, 5) & "' WHERE "
        strQuery = strQuery & "KEYID='" & RANK_OBJ.Get_Current_KEYID & "'"

        dbMyDB.Execute (strQuery)

        dbMyDB.Close
    End If

    Call Send_Panel_Judge(pubPANEL_INFO.PANELID, strNew_Grade, frmMain.flxJudge_History.TextMatrix(intRow, 5), "")

    intPTN_Index = CInt(frmJudge.lblCurrent_PTN_Index.Caption)
    EQP.Set_PATTERN_END_by_Index (intPTN_Index)
    intPortNo = EQP.Get_PG_PortID
    Call QUEUE.Put_Send_Command(intPortNo, "QPPF")
    
    Unload Me

End Sub

Private Sub cmdOK_Click()

    Dim dbMyDB                              As Database
    
    Dim typGRADE_DATA()                     As GRADE_DATA_STRUCTURE
    Dim typDEFECT_DATA()                    As DEFECT_DATA_STRUCTURE
    
    Dim typRANK_DATA                        As RANK_DATA_STRUCTURE
    Dim typGRADE_DEFECT                     As DEFECT_DATA_STRUCTURE
    Dim typPFCD_ADDRESS_DATA                As PFCD_ADDRESS_STRUCTURE
    
    Dim arrPOINT_DEFECT_COUNT(1 To 3)       As Integer
    
    Dim strDB_Path                          As String
    Dim strDB_FileName                      As String
    Dim strQuery                            As String
    Dim strNew_Grade                        As String
    Dim strPoint_Defect_Rank                As String
    Dim strGrade                            As String
    Dim strRank                             As String
    Dim strDEFECT_TYPE                      As String
    Dim strState                            As String
    
    Dim intPortNo                           As Integer
    Dim intRow                              As Integer
    Dim intCol                              As Integer
    Dim intIndex                            As Integer
    Dim intDefect_Count                     As Integer
    Dim intGrade_Defect_Index               As Integer
    Dim intGrade_Count                      As Integer
    Dim intPoint_Defect_Total               As Integer
    Dim intPTN_Index                        As Integer
    
    Call ENV.Get_Device_Data_by_Name("API", intPortNo, strState)

    If Left(frmMain.flxEQ_Information.TextMatrix(5, 1), 5) = "CATST" Then
        If intPortNo > 0 Then
            Call QUEUE.Put_Send_Command(intPortNo, "QBLV")
        End If
    End If

    Call RANK_OBJ.Set_DEFECT_DATA_COUNT(frmJudge.flxDefect_List.Rows - 1)
    intDefect_Count = frmJudge.flxDefect_List.Rows - 1
    ReDim typDEFECT_DATA(intDefect_Count)

    With frmJudge.flxDefect_List
        If intDefect_Count > 0 Then
            For intRow = 1 To intDefect_Count
                .TextMatrix(intRow, 8) = pubPANEL_INFO.PANELID
                typDEFECT_DATA(intRow).DEFECT_CODE = .TextMatrix(intRow, 0)
                typDEFECT_DATA(intRow).DEFECT_NAME = .TextMatrix(intRow, 1)
                typDEFECT_DATA(intRow).PANELID = .TextMatrix(intRow, 8)
                intIndex = 0
                For intCol = 2 To 7 Step 2
                    intIndex = intIndex + 1
                    typDEFECT_DATA(intRow).DATA_ADDRESS(intIndex) = .TextMatrix(intRow, intCol)
                    typDEFECT_DATA(intRow).GATE_ADDRESS(intIndex) = .TextMatrix(intRow, intCol + 1)
                Next intCol
                If .TextMatrix(intRow, 13) = "" Then
                    .TextMatrix(intRow, 13) = "0"
                End If
                Call RANK_OBJ.Set_DEFECT_DATA(intRow, pubPANEL_INFO.PANELID, .TextMatrix(intRow, 0), .TextMatrix(intRow, 1), .TextMatrix(intRow, 11), typDEFECT_DATA(intRow).DATA_ADDRESS, typDEFECT_DATA(intRow).GATE_ADDRESS, .TextMatrix(intRow, 12), CInt(.TextMatrix(intRow, 13)))
                If (frmJudge.flxDefect_List.TextMatrix(intRow, 9) = "") And (frmJudge.flxDefect_List.TextMatrix(intRow, 10) = "") Then
                    Call ACCUMULATE(pubCST_INFO, typDEFECT_DATA(intRow), intRow)
                Else
                    Call Get_Rank_Data(pubCST_INFO.PROCESS_NUM, typRANK_DATA, typGRADE_DATA, typDEFECT_DATA(intRow).DEFECT_CODE, intGrade_Count)
                    typDEFECT_DATA(intRow).PRIORITY = typRANK_DATA.PRIORITY
                    If typRANK_DATA.ACCUMULATION <> "X" Then
                        'Accumulation
                        Call Add_Point_Defect_Total(typDEFECT_DATA(intRow), CInt(frmJudge.flxDefect_List.TextMatrix(intRow, 10)))
                        intPoint_Defect_Total = Get_Point_Defect_Total(typDEFECT_DATA(intRow).DEFECT_CODE, typDEFECT_DATA(intRow).PANELID)

                        If typRANK_DATA.DETAIL_DIVISION = "B" Then
                            Call RANK_OBJ.Add_TB_Count(CInt(frmJudge.flxDefect_List.TextMatrix(intRow, 10)))
                        Else
                            Call RANK_OBJ.Add_TD_Count(CInt(frmJudge.flxDefect_List.TextMatrix(intRow, 10)))
                        End If
                    End If
                    If typRANK_DATA.JUDGE_OR_NOT = "O" Then
                        strGrade = ""
                        strRank = frmJudge.flxDefect_List.TextMatrix(intRow, 9)
                        For intIndex = 1 To intGrade_Count
                            If (strGrade = "") And (typDEFECT_DATA(intRow).DEFECT_CODE = typGRADE_DATA(intIndex).DEFECT_CODE) And (InStr(typGRADE_DATA(intIndex).RANK, strRank) > 0) Then
                                strGrade = typGRADE_DATA(intIndex).GRADE
                            End If
                        Next intIndex
                        typDEFECT_DATA(intRow).GRADE = strGrade
                        Call SaveLog("cmdGrade_Click", typDEFECT_DATA(intRow).DEFECT_CODE & "'s RANK : " & typDEFECT_DATA(intRow).RANK & ", GRADE : " & strGrade)
                        Call RANK_OBJ.Set_DEFECT_RANK(typDEFECT_DATA(intRow).DEFECT_CODE, strRank, typDEFECT_DATA(intRow).DATA_ADDRESS, typDEFECT_DATA(intRow).GATE_ADDRESS)
                        Call RANK_OBJ.Set_DEFECT_GRADE(typDEFECT_DATA(intRow).DEFECT_CODE, typDEFECT_DATA(intRow).DATA_ADDRESS, typDEFECT_DATA(intRow).GATE_ADDRESS, strGrade)
                    End If
                End If
            Next intRow

            Call RANK_OBJ.Init_DEFECT_PRIORITY
            For intIndex = 1 To 3
                arrPOINT_DEFECT_COUNT(intIndex) = 0
            Next intIndex
            For intIndex = 1 To intDefect_Count
                With typDEFECT_DATA(intIndex)
                    Call RANK_OBJ.Get_DEFECT_DATA_by_Index(intIndex, .PANELID, .DEFECT_CODE, .DEFECT_NAME, .DETAIL_DIVISION, .DATA_ADDRESS, .GATE_ADDRESS, .GRADE, .RANK, .COLOR, .GRAY_LEVEL)
                    strDEFECT_TYPE = Mid(.DEFECT_CODE, 2, 1)
                    If typDEFECT_DATA(intIndex).PRIORITY < RANK_OBJ.Get_DEFECT_PRIORITY_by_DEFECT_TYPE(strDEFECT_TYPE) Then
                        Call RANK_OBJ.Set_DEFECT_GRADE_by_PRIORITY(strDEFECT_TYPE, .GRADE, .PRIORITY, intIndex, .DEFECT_CODE, .RANK)
                    ElseIf typDEFECT_DATA(intIndex).PRIORITY = RANK_OBJ.Get_DEFECT_PRIORITY_by_DEFECT_TYPE(strDEFECT_TYPE) Then
                        If RANK_OBJ.Get_Rank_Priority_by_Rank(typDEFECT_DATA(intIndex).RANK) > RANK_OBJ.Get_Rank_Priority_by_Rank(strDEFECT_TYPE) Then
                            Call RANK_OBJ.Set_DEFECT_GRADE_by_PRIORITY(strDEFECT_TYPE, .GRADE, .PRIORITY, intIndex, .DEFECT_CODE, .RANK)
                        End If
                    End If
                    'Point Defect Count Check
                    Select Case .DETAIL_DIVISION
                    Case "B":
                        arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TB) = arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TB) + 1
                    Case "D":
                        arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TD) = arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TD) + 1
                    End Select
                    arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TT) = arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TB) + arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TD)
                End With
            Next intIndex

            'Line Lengh calcaulation between each point defect
            For intIndex = 1 To 3
                If arrPOINT_DEFECT_COUNT(intIndex) < 6 Then
                    'Reference W, L, B1 and B2 values in PFCD_Address.csv file
                    Call Get_PFCD_ADDRESS_DATA(typPFCD_ADDRESS_DATA, pubPANEL_INFO.PRODUCTID, CInt(Right(pubPANEL_INFO.PANELID, 2)))
                    With typPFCD_ADDRESS_DATA
                        Call RANK_OBJ.Calculate_Point_Distance(intIndex, .W, .L, .B1, .B2)
                    End With
                End If
            Next intIndex

            If strPoint_Defect_Rank = "" Then
                strPoint_Defect_Rank = pubPANEL_INFO.TFT_REPAIR_GRADE
            End If

            strNew_Grade = Get_Panel_Grade(strPoint_Defect_Rank)
            intGrade_Defect_Index = RANK_OBJ.Get_GRADE_DEFECT_INDEX
            strNew_Grade = PreJudgeGradeChange1(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strNew_Grade = PreJudgeGradeChange2(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA, intDefect_Count)
            strNew_Grade = PreJudgeGradeChange3(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index), strPoint_Defect_Rank)
            strNew_Grade = PostJudgeOtherRule1(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strNew_Grade = PostJudgeOtherRule2(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strNew_Grade = PostJudgeOtherRule3(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strNew_Grade = PostJudgeGradeChange1(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strNew_Grade = PostJudgeGradeChange2(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strNew_Grade = CheckPanelIDChangeGrade(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strNew_Grade = ChangeGrade(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strNew_Grade = ChangeGradeByDefectCode(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strNew_Grade = RepairPointTimes(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strNew_Grade = FlagChangeGrade(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index), pubJOB_INFO)
            strNew_Grade = SKChange(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
        Else
            'Get highst grade from Rank table
            strNew_Grade = RANK_OBJ.Get_Highest_Grade
        End If
    End With

    strNew_Grade = Me.cmdOK.Caption
    frmMain.lblPost_Judge.Caption = strNew_Grade
    With frmMain.flxJudge_History
        intRow = .Rows - 1
        If frmJudge.flxDefect_List.Rows > 1 Then
            .TextMatrix(intRow, 4) = strNew_Grade
            .TextMatrix(intRow, 5) = frmJudge.flxDefect_List.TextMatrix(frmJudge.flxDefect_List.Rows - 1, 0) 'typDEFECT_DATA(intGrade_Defect_Index).DEFECT_CODE
            .TextMatrix(intRow, 6) = Format(TIME, "HH:MM:SS")
        Else
            .TextMatrix(intRow, 4) = strNew_Grade
            .TextMatrix(intRow, 5) = ""
            .TextMatrix(intRow, 6) = Format(TIME, "HH:MM:SS")
        End If
    End With
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Result.mdb"

    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)

        strQuery = "UPDATE PANEL_DATA SET "
        strQuery = strQuery & "PANEL_GRADE='" & strNew_Grade & "', "
        strQuery = strQuery & "PANEL_LOSSCODE='" & frmMain.flxJudge_History.TextMatrix(intRow, 5) & "' WHERE "
        strQuery = strQuery & "KEYID='" & RANK_OBJ.Get_Current_KEYID & "'"

        dbMyDB.Execute (strQuery)

        dbMyDB.Close
    End If

    Call Send_Panel_Judge(pubPANEL_INFO.PANELID, strNew_Grade, frmMain.flxJudge_History.TextMatrix(intRow, 5), "")

    intPTN_Index = CInt(frmJudge.lblCurrent_PTN_Index.Caption)
    EQP.Set_PATTERN_END_by_Index (intPTN_Index)
    intPortNo = EQP.Get_PG_PortID
    Call QUEUE.Put_Send_Command(intPortNo, "QPPF")
    
    Unload Me

End Sub

Private Sub cmdY_Click()

    Dim dbMyDB                              As Database
    
    Dim typGRADE_DATA()                     As GRADE_DATA_STRUCTURE
    Dim typDEFECT_DATA()                    As DEFECT_DATA_STRUCTURE
    
    Dim typRANK_DATA                        As RANK_DATA_STRUCTURE
    Dim typGRADE_DEFECT                     As DEFECT_DATA_STRUCTURE
    Dim typPFCD_ADDRESS_DATA                As PFCD_ADDRESS_STRUCTURE
    
    Dim arrPOINT_DEFECT_COUNT(1 To 3)       As Integer
    
    Dim strDB_Path                          As String
    Dim strDB_FileName                      As String
    Dim strQuery                            As String
    Dim strNew_Grade                        As String
    Dim strPoint_Defect_Rank                As String
    Dim strGrade                            As String
    Dim strRank                             As String
    Dim strDEFECT_TYPE                      As String
    Dim strState                            As String
    
    Dim intPortNo                           As Integer
    Dim intRow                              As Integer
    Dim intCol                              As Integer
    Dim intIndex                            As Integer
    Dim intDefect_Count                     As Integer
    Dim intGrade_Defect_Index               As Integer
    Dim intGrade_Count                      As Integer
    Dim intPoint_Defect_Total               As Integer
    Dim intPTN_Index                        As Integer
    
    Call ENV.Get_Device_Data_by_Name("API", intPortNo, strState)

    If Left(frmMain.flxEQ_Information.TextMatrix(5, 1), 5) = "CATST" Then
        If intPortNo > 0 Then
            Call QUEUE.Put_Send_Command(intPortNo, "QBLV")
        End If
    End If

    Call RANK_OBJ.Set_DEFECT_DATA_COUNT(frmJudge.flxDefect_List.Rows - 1)
    intDefect_Count = frmJudge.flxDefect_List.Rows - 1
    ReDim typDEFECT_DATA(intDefect_Count)

    With frmJudge.flxDefect_List
        If intDefect_Count > 0 Then
            For intRow = 1 To intDefect_Count
                .TextMatrix(intRow, 8) = pubPANEL_INFO.PANELID
                typDEFECT_DATA(intRow).DEFECT_CODE = .TextMatrix(intRow, 0)
                typDEFECT_DATA(intRow).DEFECT_NAME = .TextMatrix(intRow, 1)
                typDEFECT_DATA(intRow).PANELID = .TextMatrix(intRow, 8)
                intIndex = 0
                For intCol = 2 To 7 Step 2
                    intIndex = intIndex + 1
                    typDEFECT_DATA(intRow).DATA_ADDRESS(intIndex) = .TextMatrix(intRow, intCol)
                    typDEFECT_DATA(intRow).GATE_ADDRESS(intIndex) = .TextMatrix(intRow, intCol + 1)
                Next intCol
                If .TextMatrix(intRow, 13) = "" Then
                    .TextMatrix(intRow, 13) = "0"
                End If
                Call RANK_OBJ.Set_DEFECT_DATA(intRow, pubPANEL_INFO.PANELID, .TextMatrix(intRow, 0), .TextMatrix(intRow, 1), .TextMatrix(intRow, 11), typDEFECT_DATA(intRow).DATA_ADDRESS, typDEFECT_DATA(intRow).GATE_ADDRESS, .TextMatrix(intRow, 12), CInt(.TextMatrix(intRow, 13)))
                If (frmJudge.flxDefect_List.TextMatrix(intRow, 9) = "") And (frmJudge.flxDefect_List.TextMatrix(intRow, 10) = "") Then
                    Call ACCUMULATE(pubCST_INFO, typDEFECT_DATA(intRow), intRow)
                Else
                    Call Get_Rank_Data(pubCST_INFO.PROCESS_NUM, typRANK_DATA, typGRADE_DATA, typDEFECT_DATA(intRow).DEFECT_CODE, intGrade_Count)
                    typDEFECT_DATA(intRow).PRIORITY = typRANK_DATA.PRIORITY
                    If typRANK_DATA.ACCUMULATION <> "X" Then
                        'Accumulation
                        Call Add_Point_Defect_Total(typDEFECT_DATA(intRow), CInt(frmJudge.flxDefect_List.TextMatrix(intRow, 10)))
                        intPoint_Defect_Total = Get_Point_Defect_Total(typDEFECT_DATA(intRow).DEFECT_CODE, typDEFECT_DATA(intRow).PANELID)

                        If typRANK_DATA.DETAIL_DIVISION = "B" Then
                            Call RANK_OBJ.Add_TB_Count(CInt(frmJudge.flxDefect_List.TextMatrix(intRow, 10)))
                        Else
                            Call RANK_OBJ.Add_TD_Count(CInt(frmJudge.flxDefect_List.TextMatrix(intRow, 10)))
                        End If
                    End If
                    If typRANK_DATA.JUDGE_OR_NOT = "O" Then
                        strGrade = ""
                        strRank = frmJudge.flxDefect_List.TextMatrix(intRow, 9)
                        For intIndex = 1 To intGrade_Count
                            If (strGrade = "") And (typDEFECT_DATA(intRow).DEFECT_CODE = typGRADE_DATA(intIndex).DEFECT_CODE) And (InStr(typGRADE_DATA(intIndex).RANK, strRank) > 0) Then
                                strGrade = typGRADE_DATA(intIndex).GRADE
                            End If
                        Next intIndex
                        typDEFECT_DATA(intRow).GRADE = strGrade
                        Call SaveLog("cmdGrade_Click", typDEFECT_DATA(intRow).DEFECT_CODE & "'s RANK : " & typDEFECT_DATA(intRow).RANK & ", GRADE : " & strGrade)
                        Call RANK_OBJ.Set_DEFECT_RANK(typDEFECT_DATA(intRow).DEFECT_CODE, strRank, typDEFECT_DATA(intRow).DATA_ADDRESS, typDEFECT_DATA(intRow).GATE_ADDRESS)
                        Call RANK_OBJ.Set_DEFECT_GRADE(typDEFECT_DATA(intRow).DEFECT_CODE, typDEFECT_DATA(intRow).DATA_ADDRESS, typDEFECT_DATA(intRow).GATE_ADDRESS, strGrade)
                    End If
                End If
            Next intRow

            Call RANK_OBJ.Init_DEFECT_PRIORITY
            For intIndex = 1 To 3
                arrPOINT_DEFECT_COUNT(intIndex) = 0
            Next intIndex
            For intIndex = 1 To intDefect_Count
                With typDEFECT_DATA(intIndex)
                    Call RANK_OBJ.Get_DEFECT_DATA_by_Index(intIndex, .PANELID, .DEFECT_CODE, .DEFECT_NAME, .DETAIL_DIVISION, .DATA_ADDRESS, .GATE_ADDRESS, .GRADE, .RANK, .COLOR, .GRAY_LEVEL)
                    strDEFECT_TYPE = Mid(.DEFECT_CODE, 2, 1)
                    If typDEFECT_DATA(intIndex).PRIORITY < RANK_OBJ.Get_DEFECT_PRIORITY_by_DEFECT_TYPE(strDEFECT_TYPE) Then
                        Call RANK_OBJ.Set_DEFECT_GRADE_by_PRIORITY(strDEFECT_TYPE, .GRADE, .PRIORITY, intIndex, .DEFECT_CODE, .RANK)
                    ElseIf typDEFECT_DATA(intIndex).PRIORITY = RANK_OBJ.Get_DEFECT_PRIORITY_by_DEFECT_TYPE(strDEFECT_TYPE) Then
                        If RANK_OBJ.Get_Rank_Priority_by_Rank(typDEFECT_DATA(intIndex).RANK) > RANK_OBJ.Get_Rank_Priority_by_Rank(strDEFECT_TYPE) Then
                            Call RANK_OBJ.Set_DEFECT_GRADE_by_PRIORITY(strDEFECT_TYPE, .GRADE, .PRIORITY, intIndex, .DEFECT_CODE, .RANK)
                        End If
                    End If
                    'Point Defect Count Check
                    Select Case .DETAIL_DIVISION
                    Case "B":
                        arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TB) = arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TB) + 1
                    Case "D":
                        arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TD) = arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TD) + 1
                    End Select
                    arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TT) = arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TB) + arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TD)
                End With
            Next intIndex

            'Line Lengh calcaulation between each point defect
            For intIndex = 1 To 3
                If arrPOINT_DEFECT_COUNT(intIndex) < 6 Then
                    'Reference W, L, B1 and B2 values in PFCD_Address.csv file
                    Call Get_PFCD_ADDRESS_DATA(typPFCD_ADDRESS_DATA, pubPANEL_INFO.PRODUCTID, CInt(Right(pubPANEL_INFO.PANELID, 2)))
                    With typPFCD_ADDRESS_DATA
                        Call RANK_OBJ.Calculate_Point_Distance(intIndex, .W, .L, .B1, .B2)
                    End With
                End If
            Next intIndex

            If strPoint_Defect_Rank = "" Then
                strPoint_Defect_Rank = pubPANEL_INFO.TFT_REPAIR_GRADE
            End If

            strNew_Grade = Get_Panel_Grade(strPoint_Defect_Rank)
            intGrade_Defect_Index = RANK_OBJ.Get_GRADE_DEFECT_INDEX
            strNew_Grade = PreJudgeGradeChange1(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strNew_Grade = PreJudgeGradeChange2(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA, intDefect_Count)
            strNew_Grade = PreJudgeGradeChange3(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index), strPoint_Defect_Rank)
            strNew_Grade = PostJudgeOtherRule1(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strNew_Grade = PostJudgeOtherRule2(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strNew_Grade = PostJudgeOtherRule3(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strNew_Grade = PostJudgeGradeChange1(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strNew_Grade = PostJudgeGradeChange2(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strNew_Grade = CheckPanelIDChangeGrade(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strNew_Grade = ChangeGrade(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strNew_Grade = ChangeGradeByDefectCode(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strNew_Grade = RepairPointTimes(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
            strNew_Grade = FlagChangeGrade(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index), pubJOB_INFO)
            strNew_Grade = SKChange(strNew_Grade, pubCST_INFO, pubPANEL_INFO, typDEFECT_DATA(intGrade_Defect_Index))
        Else
            'Get highst grade from Rank table
            strNew_Grade = RANK_OBJ.Get_Highest_Grade
        End If
    End With

    strNew_Grade = Me.cmdY.Caption
    frmMain.lblPost_Judge.Caption = strNew_Grade
    With frmMain.flxJudge_History
        intRow = .Rows - 1
        If frmJudge.flxDefect_List.Rows > 1 Then
            .TextMatrix(intRow, 4) = strNew_Grade
            .TextMatrix(intRow, 5) = frmJudge.flxDefect_List.TextMatrix(frmJudge.flxDefect_List.Rows - 1, 0) 'typDEFECT_DATA(intGrade_Defect_Index).DEFECT_CODE
            .TextMatrix(intRow, 6) = Format(TIME, "HH:MM:SS")
        Else
            .TextMatrix(intRow, 4) = strNew_Grade
            .TextMatrix(intRow, 5) = ""
            .TextMatrix(intRow, 6) = Format(TIME, "HH:MM:SS")
        End If
    End With
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Result.mdb"

    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)

        strQuery = "UPDATE PANEL_DATA SET "
        strQuery = strQuery & "PANEL_GRADE='" & strNew_Grade & "', "
        strQuery = strQuery & "PANEL_LOSSCODE='" & frmMain.flxJudge_History.TextMatrix(intRow, 5) & "' WHERE "
        strQuery = strQuery & "KEYID='" & RANK_OBJ.Get_Current_KEYID & "'"

        dbMyDB.Execute (strQuery)

        dbMyDB.Close
    End If

    Call Send_Panel_Judge(pubPANEL_INFO.PANELID, strNew_Grade, frmMain.flxJudge_History.TextMatrix(intRow, 5), "")

    intPTN_Index = CInt(frmJudge.lblCurrent_PTN_Index.Caption)
    EQP.Set_PATTERN_END_by_Index (intPTN_Index)
    intPortNo = EQP.Get_PG_PortID
    Call QUEUE.Put_Send_Command(intPortNo, "QPPF")
    
    Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Unload frmJudge
    
End Sub
