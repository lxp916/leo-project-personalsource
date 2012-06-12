VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRankLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Sub Init_Class()

End Sub

Public Function GetRankLevelIntoHashTable(ByVal ht As clsHashTable)
    Dim strDB_Path                          As String
    Dim strDB_FileName                      As String
    
    Dim intOldCount                         As Integer
    
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Parameter.mdb"
    
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        strQuery = "SELECT * FROM Rank_Level order by RankLevel"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        intOldCount = lstRecord.RecordCount
        If intOldCount > 0 Then
            For intOldCount = 0 To lstRecord.RecordCount - 1
                If ht.Exists(lstRecord("RankCode")) = False Then
                    ht.Add(lstRecord("RankCode")) = ""
                End If
                lstRecord.MoveNext
            Next intOldCount
        End If
    End If
    
    lstRecord.Close
End Function



