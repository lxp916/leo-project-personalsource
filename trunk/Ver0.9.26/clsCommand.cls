VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim RECEIVE_COMMAND_LIST(1 To 8, 0 To cQUEUE_MAX)       As String
Dim SEND_COMMAND_LIST(1 To 8, 0 To cQUEUE_MAX)          As String
Dim LOG_LIST(0 To cQUEUE_MAX)                           As String

Dim LAST_COMMAND(1 To 8)                                As String
Dim INPUT_COMMAND(1 To 8)                               As String

Dim RECEIVE_FRONT(1 To 8)                               As Integer
Dim RECEIVE_REAR(1 To 8)                                As Integer
Dim SEND_FRONT(1 To 8)                                  As Integer
Dim SEND_REAR(1 To 8)                                   As Integer
Dim RETRY_COUNT(1 To 8)                                 As Integer

Dim LOG_FRONT                                           As Integer
Dim LOG_REAR                                            As Integer

Dim TIMEOUT(1 To 8)                                     As Double

Public Sub Init_Class()

    Dim intPortNo               As Integer
    Dim intIndex                As Integer
    
    For intPortNo = 1 To 8
        For intIndex = 0 To cQUEUE_MAX
            RECEIVE_COMMAND_LIST(intPortNo, intIndex) = ""
            SEND_COMMAND_LIST(intPortNo, intIndex) = ""
        Next intIndex
        LAST_COMMAND(intPortNo) = ""
        RECEIVE_FRONT(intPortNo) = -1
        RECEIVE_REAR(intPortNo) = -1
        SEND_FRONT(intPortNo) = -1
        SEND_REAR(intPortNo) = -1
        TIMEOUT(intPortNo) = -1
        RETRY_COUNT(intPortNo) = 0
        INPUT_COMMAND(intPortNo) = ""
    Next intPortNo
    
    For intIndex = 0 To cQUEUE_MAX
        LOG_LIST(intIndex) = ""
    Next intIndex
    LOG_FRONT = -1
    LOG_REAR = -1
    
End Sub

Public Function Get_Receive_Que_Empty(ByVal pPortNo As Integer) As Integer

    If (0 < pPortNo) And (pPortNo < 9) Then
        If RECEIVE_FRONT(pPortNo) = RECEIVE_REAR(pPortNo) Then
            Get_Receive_Que_Empty = cCOMMAND_QUEUE_EMPTY
        Else
            Get_Receive_Que_Empty = cCOMMAND_QUEUE_NORMALCY
        End If
    Else
        Get_Receive_Que_Empty = cCOMMAND_QUEUE_INVALID_PORT_NO
    End If
    
End Function

Public Function Get_Receive_Que_Full(ByVal pPortNo As Integer) As Integer

    If (0 < pPortNo) And (pPortNo < 9) Then
        If ((RECEIVE_FRONT(pPortNo) + 1) Mod cQUEUE_MAX) = RECEIVE_REAR(pPortNo) Then
            Get_Receive_Que_Full = cCOMMAND_QUEUE_FULL
        Else
            Get_Receive_Que_Full = cCOMMAND_QUEUE_NORMALCY
        End If
    Else
        Get_Receive_Que_Full = cCOMMAND_QUEUE_INVALID_PORT_NO
    End If
    
End Function

Public Function Get_Receive_Command(ByVal pPortNo As Integer, pCommand As String) As Integer

    If (0 < pPortNo) And (pPortNo < 9) Then
        Get_Receive_Command = Get_Receive_Que_Empty(pPortNo)
        If Get_Receive_Command = cCOMMAND_QUEUE_NORMALCY Then
            RECEIVE_REAR(pPortNo) = (RECEIVE_REAR(pPortNo) + 1) Mod cQUEUE_MAX
            pCommand = RECEIVE_COMMAND_LIST(pPortNo, RECEIVE_REAR(pPortNo))
        Else
            pCommand = ""
        End If
    Else
        Get_Receive_Command = cCOMMAND_QUEUE_INVALID_PORT_NO
    End If
    
End Function

Public Function Put_Receive_Command(ByVal pPortNo As Integer, ByVal pCommand As String) As Integer

    If (0 < pPortNo) And (pPortNo < 9) Then
        Put_Receive_Command = Get_Receive_Que_Full(pPortNo)
        If Put_Receive_Command = cCOMMAND_QUEUE_NORMALCY Then
            RECEIVE_FRONT(pPortNo) = (RECEIVE_FRONT(pPortNo) + 1) Mod cQUEUE_MAX
            RECEIVE_COMMAND_LIST(pPortNo, RECEIVE_FRONT(pPortNo)) = pCommand
        End If
    Else
        Put_Receive_Command = cCOMMAND_QUEUE_INVALID_PORT_NO
    End If
    
End Function

Public Function Get_Send_Que_Empty(ByVal pPortNo As Integer) As Integer

    If (0 < pPortNo) And (pPortNo < 9) Then
        If SEND_FRONT(pPortNo) = SEND_REAR(pPortNo) Then
            Get_Send_Que_Empty = cCOMMAND_QUEUE_EMPTY
        Else
            Get_Send_Que_Empty = cCOMMAND_QUEUE_NORMALCY
        End If
    Else
        Get_Send_Que_Empty = cCOMMAND_QUEUE_INVALID_PORT_NO
    End If
    
End Function

Public Function Get_Send_Que_Full(ByVal pPortNo As Integer) As Integer

    If (0 < pPortNo) And (pPortNo < 9) Then
        If ((SEND_FRONT(pPortNo) + 1) Mod cQUEUE_MAX) = SEND_REAR(pPortNo) Then
            Get_Send_Que_Full = cCOMMAND_QUEUE_FULL
        Else
            Get_Send_Que_Full = cCOMMAND_QUEUE_NORMALCY
        End If
    Else
        Get_Send_Que_Full = cCOMMAND_QUEUE_INVALID_PORT_NO
    End If
    
End Function

Public Function Get_Send_Command(ByVal pPortNo As Integer, pCommand As String) As Integer

    If (0 < pPortNo) And (pPortNo < 9) Then
        Get_Send_Command = Get_Send_Que_Empty(pPortNo)
        If Get_Send_Command = cCOMMAND_QUEUE_NORMALCY Then
            SEND_REAR(pPortNo) = (SEND_REAR(pPortNo) + 1) Mod cQUEUE_MAX
            pCommand = SEND_COMMAND_LIST(pPortNo, SEND_REAR(pPortNo))
            LAST_COMMAND(pPortNo) = pCommand
            RETRY_COUNT(pPortNo) = 0
        Else
            pCommand = ""
        End If
    Else
        Get_Send_Command = cCOMMAND_QUEUE_INVALID_PORT_NO
    End If
    
End Function

Public Function Put_Send_Command(ByVal pPortNo As Integer, ByVal pCommand As String) As Integer

    If (0 < pPortNo) And (pPortNo < 9) Then
        Put_Send_Command = Get_Send_Que_Full(pPortNo)
        If Put_Send_Command = cCOMMAND_QUEUE_NORMALCY Then
            SEND_FRONT(pPortNo) = (SEND_FRONT(pPortNo) + 1) Mod cQUEUE_MAX
            SEND_COMMAND_LIST(pPortNo, SEND_FRONT(pPortNo)) = cSTX & pCommand & cETX
        End If
    Else
        Put_Send_Command = cCOMMAND_QUEUE_INVALID_PORT_NO
    End If
    
End Function

Public Function Set_TimeOut(ByVal pPortNo As Integer, ByVal pTime As Double) As Integer

    If (0 < pPortNo) And (pPortNo < 9) Then
        TIMEOUT(pPortNo) = pTime
        Set_TimeOut = cCOMMAND_QUEUE_NORMALCY
    Else
        Set_TimeOut = cCOMMAND_QUEUE_INVALID_PORT_NO
    End If
    
End Function

Public Function Get_TimeOut(ByVal pPortNo As Integer) As Double

    If (0 < pPortNo) And (pPortNo < 9) Then
        Get_TimeOut = TIMEOUT(pPortNo)
    Else
        Get_TimeOut = -1
    End If
    
End Function

Public Function Get_Last_Command(ByVal pPortNo As Integer, pCommand As String) As Integer

    If (0 < pPortNo) And (pPortNo < 9) Then
        Get_Last_Command = cCOMMAND_QUEUE_NORMALCY
        pCommand = LAST_COMMAND(pPortNo)
    Else
        Get_Last_Command = cCOMMAND_QUEUE_INVALID_PORT_NO
    End If
    
End Function

Public Sub Reset_Last_Command(ByVal pPortNo As Integer)

    LAST_COMMAND(pPortNo) = ""
    RETRY_COUNT(pPortNo) = 0
    
End Sub

Public Sub Reset_Retry_Count(ByVal pPortNo As Integer)

    If (0 < pPortNo) And (pPortNo < 9) Then
        LAST_COMMAND(pPortNo) = ""
        RETRY_COUNT(pPortNo) = 0
    End If

End Sub

Public Function Get_Retry_Count(ByVal pPortNo As Integer) As Integer

    If (0 < pPortNo) And (pPortNo < 9) Then
        Get_Retry_Count = RETRY_COUNT(pPortNo)
    Else
        Get_Retry_Count = -1
    End If

End Function

Public Sub Increase_Retry_Count(ByVal pPortNo As Integer)

    If (0 < pPortNo) And (pPortNo < 9) Then
        RETRY_COUNT(pPortNo) = RETRY_COUNT(pPortNo) + 1
    End If

End Sub

Public Sub Set_Input_Command(ByVal pPortNo As Integer, ByVal pCommand As String)

    INPUT_COMMAND(pPortNo) = pCommand
    
End Sub

Public Function Get_Input_Command(ByVal pPortNo As Integer) As String

    If (0 < pPortNo) And (pPortNo < 9) Then
        Get_Input_Command = INPUT_COMMAND(pPortNo)
    Else
        Get_Input_Command = ""
    End If
    
End Function

Public Function Get_Log_Que_Empty() As Integer

    If LOG_FRONT = LOG_REAR Then
        Get_Log_Que_Empty = cCOMMAND_QUEUE_EMPTY
    Else
        Get_Log_Que_Empty = cCOMMAND_QUEUE_NORMALCY
    End If
    
End Function

Public Function Get_Log_Que_Full() As Integer

    If ((LOG_FRONT + 1) Mod cQUEUE_MAX) = LOG_REAR Then
        Get_Log_Que_Full = cCOMMAND_QUEUE_FULL
    Else
        Get_Log_Que_Full = cCOMMAND_QUEUE_NORMALCY
    End If
    
End Function

Public Function Get_Log_Data(pLog As String) As Integer

    Get_Log_Data = Get_Log_Que_Empty
    If Get_Log_Data = cCOMMAND_QUEUE_NORMALCY Then
        LOG_REAR = (LOG_REAR + 1) Mod cQUEUE_MAX
        pLog = LOG_LIST(LOG_REAR)
    Else
        pLog = ""
    End If
    
End Function

Public Function Put_Log_Data(ByVal pLog As String) As Integer

    Put_Log_Data = Get_Log_Que_Full
    If Put_Log_Data = cCOMMAND_QUEUE_NORMALCY Then
        LOG_FRONT = (LOG_FRONT + 1) Mod cQUEUE_MAX
        LOG_LIST(LOG_FRONT) = pLog
    End If
    
End Function

