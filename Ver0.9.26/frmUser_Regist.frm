VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmUser_Regist 
   Caption         =   "User Regist"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   11565
   StartUpPosition =   2  'È­¸é °¡¿îµ¥
   Begin VB.Frame Frame1 
      Height          =   5955
      Left            =   9150
      TabIndex        =   3
      Top             =   0
      Width           =   2415
      Begin VB.CommandButton cmdDelete 
         Caption         =   "DELETE"
         Height          =   525
         Left            =   570
         TabIndex        =   17
         Top             =   5340
         Width           =   1245
      End
      Begin VB.CommandButton cmdAdd_Modify 
         Caption         =   "MODIFY"
         Height          =   525
         Left            =   570
         TabIndex        =   16
         Top             =   4830
         Width           =   1245
      End
      Begin VB.ComboBox cmbUser_Level 
         Height          =   300
         Left            =   180
         TabIndex        =   15
         Top             =   4380
         Width           =   2025
      End
      Begin VB.TextBox txtPassword2 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         Height          =   315
         Left            =   180
         MaxLength       =   8
         TabIndex        =   13
         Top             =   3600
         Width           =   2025
      End
      Begin VB.TextBox txtPassword1 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         Height          =   315
         Left            =   180
         MaxLength       =   6
         TabIndex        =   11
         Top             =   2820
         Width           =   2025
      End
      Begin VB.TextBox txtID_Card_Code 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         Height          =   315
         Left            =   180
         MaxLength       =   10
         TabIndex        =   9
         Top             =   2040
         Width           =   2025
      End
      Begin VB.TextBox txtUser_Name 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         Height          =   315
         Left            =   180
         MaxLength       =   8
         TabIndex        =   7
         Top             =   1260
         Width           =   2025
      End
      Begin VB.TextBox txtUser_Number 
         Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
         Height          =   315
         Left            =   180
         MaxLength       =   6
         TabIndex        =   5
         Top             =   510
         Width           =   2025
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "User Level"
         Height          =   180
         Left            =   210
         TabIndex        =   14
         Top             =   4110
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Password2"
         Height          =   180
         Left            =   210
         TabIndex        =   12
         Top             =   3330
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Password1"
         Height          =   180
         Left            =   210
         TabIndex        =   10
         Top             =   2550
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "ID Card Code"
         Height          =   180
         Left            =   210
         TabIndex        =   8
         Top             =   1770
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "User Name"
         Height          =   180
         Left            =   210
         TabIndex        =   6
         Top             =   990
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Private Number"
         Height          =   180
         Left            =   210
         TabIndex        =   4
         Top             =   270
         Width           =   1305
      End
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Accept"
      Height          =   525
      Left            =   3900
      TabIndex        =   2
      Top             =   6030
      Width           =   1245
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   525
      Left            =   6300
      TabIndex        =   1
      Top             =   6030
      Width           =   1245
   End
   Begin MSFlexGridLib.MSFlexGrid flxUser_Info 
      Height          =   5955
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   10504
      _Version        =   393216
      Rows            =   1
      Cols            =   6
      FixedCols       =   0
   End
End
Attribute VB_Name = "frmUser_Regist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAccept_Click()

    Dim dbMyDB                      As Database
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    Dim strPath                     As String
    Dim strFileName                 As String
    Dim strTemp                     As String
    
    Dim intFileNum                  As Integer
    Dim intRow                      As Integer
    
    Dim ErrMsg                      As String
    
On Error GoTo ErrorHandler

    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Parameter.mdb"
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
    
        strQuery = "DELETE * FROM USER_DATA"
        
        dbMyDB.Execute (strQuery)
        
        If Me.flxUser_Info.Rows > 1 Then
            strPath = App.PATH & "\Env\Standard_Info\"
            strFileName = "User.csv"
            intFileNum = FreeFile
            
            Open strPath & strFileName For Output As intFileNum
            
            strTemp = "Cardnum,name,password1,ID Card Code,password2,Right (S,E,P,T)"
            Print #intFileNum, strTemp
            With Me.flxUser_Info
                For intRow = 1 To .Rows - 1
                    strQuery = "INSERT INTO USER_DATA VALUES ("
                    strQuery = strQuery & "'" & .TextMatrix(intRow, 0) & "', "
                    strQuery = strQuery & "'" & .TextMatrix(intRow, 1) & "', "
                    strQuery = strQuery & "'" & .TextMatrix(intRow, 2) & "', "
                    strQuery = strQuery & "'" & .TextMatrix(intRow, 3) & "', "
                    strQuery = strQuery & "'" & .TextMatrix(intRow, 4) & "', "
                    strQuery = strQuery & "'" & .TextMatrix(intRow, 5) & "')"
                    
                    dbMyDB.Execute (strQuery)
                    
                    strTemp = .TextMatrix(intRow, 0) & "," & .TextMatrix(intRow, 1) & "," & .TextMatrix(intRow, 2) & ","
                    strTemp = strTemp & .TextMatrix(intRow, 3) & "," & .TextMatrix(intRow, 4) & "," & .TextMatrix(intRow, 5)
                    Print #intFileNum, strTemp
                Next intRow
            End With
            
            Close intFileNum
        End If
        dbMyDB.Close
        
        DBEngine.CompactDatabase strDB_Path & strDB_FileName, strDB_Path & "Parameter_Temp.mdb", dbLangChineseSimplified
        Kill strDB_Path & strDB_FileName
        Name strDB_Path & "Parameter_Temp.mdb" As strDB_Path & strDB_FileName
        
        Call Put_File_To_Host("User.csv", "User", strPath)
    End If
        
    Unload Me
    
    Exit Sub
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("cmdAccept_Click", ErrMsg)
    
End Sub

Private Sub cmdAdd_Modify_Click()

    Dim intRowIndex                 As Integer
    Dim intRow                      As Integer
    
    intRow = 0
    
    For intRowIndex = 1 To Me.flxUser_Info.Rows - 1
        If Me.txtUser_Number.Text = Me.flxUser_Info.TextMatrix(intRowIndex, 0) Then
            intRow = intRowIndex
        End If
    Next intRowIndex
    
    If intRow = 0 Then
        intRow = Add_Grid(Me.txtUser_Number.Text)
    End If
    With Me.flxUser_Info
        .TextMatrix(intRow, 1) = Me.txtUser_Name.Text
        .TextMatrix(intRow, 2) = Me.txtID_Card_Code.Text
        .TextMatrix(intRow, 3) = Me.txtPassword1.Text
        .TextMatrix(intRow, 4) = Me.txtPassword2.Text
        Select Case Me.cmbUser_Level.Text
        Case "SUPER":
            .TextMatrix(intRow, 5) = "S"
        Case "ENGINEER":
            .TextMatrix(intRow, 5) = "E"
        Case "PM":
            .TextMatrix(intRow, 5) = "P"
        Case "TA":
            .TextMatrix(intRow, 5) = "T"
        End Select
    End With
    
End Sub

Private Sub cmdClose_Click()

    Unload Me
    
End Sub

Private Sub cmdDelete_Click()

    Dim intRowIndex                 As Integer
    Dim intRow                      As Integer
    
    intRow = 0
    
    For intRowIndex = 1 To Me.flxUser_Info.Rows - 1
        If Me.txtUser_Name.Text = Me.flxUser_Info.TextMatrix(intRowIndex, 0) Then
            intRow = intRowIndex
        End If
    Next intRowIndex
    
    If intRow > 0 Then
        If Me.flxUser_Info.Rows = 2 Then
            Me.flxUser_Info.Rows = 1
        Else
            Me.flxUser_Info.RemoveItem (intRow)
        End If
    End If
    
End Sub

Private Sub flxUser_Info_Click()

    Dim intRow                      As Integer
    
    intRow = Me.flxUser_Info.Row
    
    If intRow > 0 Then
        With Me.flxUser_Info
            Me.txtUser_Number.Text = .TextMatrix(intRow, 0)
            Me.txtUser_Name.Text = .TextMatrix(intRow, 1)
            Me.txtID_Card_Code.Text = .TextMatrix(intRow, 2)
            Me.txtPassword1.Text = .TextMatrix(intRow, 3)
            Me.txtPassword2.Text = .TextMatrix(intRow, 4)
            Select Case .TextMatrix(intRow, 5)
            Case "S":
                Me.cmbUser_Level.Text = "SUPER"
            Case "E":
                Me.cmbUser_Level.Text = "ENGINEER"
            Case "P":
                Me.cmbUser_Level.Text = "PM"
            Case "T":
                Me.cmbUser_Level.Text = "TA"
            End Select
        End With
    End If
    
End Sub

Private Sub Form_Load()

    Call Init_Grid
    Call Init_Form
    Call Fill_Data
    
End Sub

Private Sub Init_Grid()

    Dim intCol                      As Integer
    Dim intRow                      As Integer
    
    With Me.flxUser_Info
        .Rows = 1
        For intRow = 0 To .Rows - 1
            For intCol = 0 To .Cols - 1
                .Row = intRow
                .Col = intCol
                .CellAlignment = flexAlignCenterCenter
                .RowHeight(intRow) = 350
                .ColWidth(intCol) = 1400
            Next intCol
        Next intRow
        
        .TextMatrix(0, 0) = "User No."
        .TextMatrix(0, 1) = "Name"
        .TextMatrix(0, 2) = "Card Code"
        .TextMatrix(0, 3) = "PW1"
        .TextMatrix(0, 4) = "PW2"
        .TextMatrix(0, 5) = "LEVEL"
    End With

End Sub

Private Sub Init_Form()

    With Me.cmbUser_Level
        .Clear
        .AddItem "SUPER"
        .AddItem "ENGINEER"
        .AddItem "PM"
        .AddItem "TA"
        .Text = "SUPER"
    End With
    
End Sub

Private Sub Fill_Data()

    Dim dbMyDB                          As Database
    
    Dim lstRecord                       As Recordset
    
    Dim strDB_Path                      As String
    Dim strDB_FileName                  As String
    Dim strQuery                        As String
    
    Dim intRow                          As Integer
    
    Dim ErrMsg                          As String
    
On Error GoTo ErrorHandler

    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Parameter.mdb"
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
        strQuery = "SELECT * FROM USER_DATA"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        
        If lstRecord.EOF = False Then
            lstRecord.MoveFirst
            While lstRecord.EOF = False
                With Me.flxUser_Info
                    intRow = Add_Grid(lstRecord.Fields("USER_ID"))
                    .TextMatrix(intRow, 1) = lstRecord.Fields("USER_NAME")
                    .TextMatrix(intRow, 2) = lstRecord.Fields("ID_CARD_CODE")
                    .TextMatrix(intRow, 3) = lstRecord.Fields("USER_PW1")
                    .TextMatrix(intRow, 4) = lstRecord.Fields("USER_PW2")
                    .TextMatrix(intRow, 5) = lstRecord.Fields("USER_LEVEL")
                End With
                lstRecord.MoveNext
            Wend
        End If
        lstRecord.Close
        
        dbMyDB.Close
    End If
    
    Exit Sub
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("frmUser_Regist_Fill_Data", ErrMsg)
    
End Sub

Private Function Add_Grid(ByVal pUser_Number As String) As Integer

    Dim intRow          As Integer
    Dim intCol          As Integer
    
    With Me.flxUser_Info
        intRow = .Rows
        .AddItem pUser_Number
        .RowHeight(intRow) = 350
        For intCol = 0 To .Cols - 1
            .Row = intRow
            .Col = intCol
            .CellAlignment = flexAlignCenterCenter
        Next intCol
        
        Add_Grid = .Rows - 1
    End With

End Function

