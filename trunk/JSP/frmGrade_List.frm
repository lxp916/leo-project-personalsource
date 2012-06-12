VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmGrade_List 
   Caption         =   "GRADE LIST"
   ClientHeight    =   2955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3765
   LinkTopic       =   "Form1"
   ScaleHeight     =   2955
   ScaleWidth      =   3765
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   525
      Left            =   2490
      TabIndex        =   5
      Top             =   1860
      Width           =   1245
   End
   Begin VB.CommandButton cmdRegist 
      Caption         =   "Regist"
      Height          =   525
      Left            =   2490
      TabIndex        =   4
      Top             =   1320
      Width           =   1245
   End
   Begin VB.TextBox txtGrade 
      Alignment       =   1  '오른쪽 맞춤
      Height          =   300
      Left            =   2490
      MaxLength       =   2
      TabIndex        =   3
      Top             =   450
      Width           =   1155
   End
   Begin MSFlexGridLib.MSFlexGrid flxGrade_List 
      Height          =   2955
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   5212
      _Version        =   393216
      Rows            =   1
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Save/Close"
      Height          =   525
      Left            =   2490
      TabIndex        =   0
      Top             =   2400
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "GRADE"
      Height          =   180
      Left            =   2460
      TabIndex        =   2
      Top             =   210
      Width           =   615
   End
End
Attribute VB_Name = "frmGrade_List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()

    Dim strPath                     As String
    Dim strFileName                 As String
    Dim strTemp                     As String
    
    Dim intFileNum                  As Integer
    Dim intRow                      As Integer
    Dim intCol                      As Integer
    Dim intTemp                     As Integer
    Dim intLoopCount                As Integer
    
    strPath = App.PATH & "\Env\"
    strFileName = "Rank_Interface_Base.cfg"
    intFileNum = FreeFile
    
    Open strPath & strFileName For Output As intFileNum
    
    With frmRank_Interface
        .flxGrade.Rows = 1
        For intRow = 1 To Me.flxGrade_List.Rows - 1
            strTemp = Me.flxGrade_List.TextMatrix(intRow, 1)
            Call Add_Grade_Grid(strTemp)
            Print #intFileNum, strTemp
        Next intRow
        For intCol = 0 To .flxRank_Data.Cols - 1
            If UCase(.flxRank_Data.TextMatrix(0, intCol)) = "PRIORITY" Then
                intTemp = intCol
            End If
        Next intCol
        .flxRank_Data.Cols = .flxGrade.Rows + intTemp
        For intCol = 1 To .flxGrade.Rows - 1
            .flxRank_Data.TextMatrix(0, intCol + intTemp) = .flxGrade.TextMatrix(intCol, 0)
            .flxRank_Data.Row = 0
            .flxRank_Data.Col = intCol + intTemp
            .flxRank_Data.CellAlignment = flexAlignCenterCenter
        Next intCol
    End With
    
    Close intFileNum
    
    Unload Me
    
End Sub

Private Sub cmdDelete_Click()

    Dim intRow                      As Integer
    Dim intRowCount                 As Integer
    Dim intRowIndex                 As Integer
    
    intRowCount = Me.flxGrade_List.Rows - 1
    
    If intRowCount > 0 Then
        If Me.txtGrade.Text <> "" Then
            intRowIndex = 0
            For intRow = 1 To intRowCount
                If Me.flxGrade_List.TextMatrix(intRow, 1) = Me.txtGrade.Text Then
                    intRowIndex = intRow
                End If
            Next intRow
            If intRowIndex > 0 Then
                If intRowCount = 1 Then
                    Me.flxGrade_List.Rows = 1
                Else
                    Me.flxGrade_List.RemoveItem (intRowIndex)
                End If
            End If
            Me.txtGrade.Text = ""
            For intRow = 1 To Me.flxGrade_List.Rows - 1
                Me.flxGrade_List.TextMatrix(intRow, 0) = intRow
            Next intRow
        End If
    End If
    
End Sub

Private Sub cmdRegist_Click()

    Dim intRow                      As Integer
    Dim intRowCount                 As Integer
    Dim intRowIndex                 As Integer
    
    intRowCount = Me.flxGrade_List.Rows - 1
    
'    If intRowCount < 10 Then
        If Me.txtGrade.Text <> "" Then
            intRowIndex = 0
            For intRow = 1 To intRowCount
                If Me.flxGrade_List.TextMatrix(intRow, 1) = Me.txtGrade.Text Then
                    intRowIndex = intRow
                End If
            Next intRow
            
            If intRowIndex > 0 Then
                Me.flxGrade_List.TextMatrix(intRowIndex, 1) = UCase(Me.txtGrade.Text)
                Me.txtGrade.Text = ""
            Else
                intRowIndex = Add_Grid
                Me.flxGrade_List.TextMatrix(intRowIndex, 1) = UCase(Me.txtGrade.Text)
                Me.txtGrade.Text = ""
            End If
        End If
'    Else
'        Call MsgBox("Maximum grade count is 10.", vbOKOnly, "Warning")
'    End If
    
End Sub

Private Sub flxGrade_List_Click()

    Dim intRow                      As Integer
    
    intRow = Me.flxGrade_List.Row
    
    If intRow > 0 Then
        Me.txtGrade.Text = Me.flxGrade_List.TextMatrix(intRow, 1)
    End If
    
End Sub

Private Sub Form_Load()

    Call Init_Grid
    Call Fill_Grid
    
End Sub

Private Sub Init_Grid()

    Dim intCol                      As Integer
    Dim intRow                      As Integer
    
    With Me.flxGrade_List
        For intRow = 0 To .Rows - 1
            For intCol = 0 To .Cols - 1
                .Row = intRow
                .Col = intCol
                .CellAlignment = flexAlignCenterCenter
                .RowHeight(intRow) = 350
            Next intCol
        Next intRow
        
        .ColWidth(0) = 800
        .ColWidth(1) = 1200
        
        .TextMatrix(0, 0) = "No."
        .TextMatrix(0, 1) = "GRADE"
    End With

End Sub

Private Sub Fill_Grid()

    Dim strPath                     As String
    Dim strFileName                 As String
    Dim strTemp                     As String
    
    Dim intFileNum                  As Integer
    Dim intRow                      As Integer
    
    strPath = App.PATH & "\Env\"
    strFileName = "Rank_Interface_Base.cfg"
    
    If Dir(strPath & strFileName, vbNormal) <> "" Then
        intFileNum = FreeFile
        
        Open strPath & strFileName For Input As intFileNum
        
        While Not EOF(intFileNum)
            Line Input #intFileNum, strTemp
            
            intRow = Add_Grid
            Me.flxGrade_List.TextMatrix(intRow, 1) = strTemp
        Wend
        
        Close intFileNum
    End If
    
End Sub

Private Function Add_Grid() As Integer

    Dim intRow          As Integer
    Dim intCol          As Integer
    
    With Me.flxGrade_List
        intRow = .Rows
        .AddItem intRow
        .RowHeight(intRow) = 350
        For intCol = 0 To .Cols - 1
            .Row = intRow
            .Col = intCol
            .CellAlignment = flexAlignCenterCenter
        Next intCol
        
        Add_Grid = .Rows - 1
    End With

End Function

Private Sub Add_Grade_Grid(ByVal pGrade As String)

    Dim intRow          As Integer
    Dim intCol          As Integer
    
    With frmRank_Interface.flxGrade
        intRow = .Rows
        .AddItem pGrade
        .RowHeight(intRow) = 350
        For intCol = 0 To .Cols - 1
            .Row = intRow
            .Col = intCol
            .CellAlignment = flexAlignCenterCenter
        Next intCol
    End With

End Sub

