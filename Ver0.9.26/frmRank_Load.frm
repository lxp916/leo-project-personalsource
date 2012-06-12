VERSION 5.00
Begin VB.Form frmRank_Load 
   Caption         =   "File Load"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   ScaleHeight     =   4995
   ScaleWidth      =   6840
   StartUpPosition =   2  '화면 가운데
   Begin VB.TextBox txtFileName 
      Height          =   300
      Left            =   4440
      TabIndex        =   7
      Top             =   3690
      Width           =   2385
   End
   Begin VB.ComboBox cmbPattern 
      Height          =   300
      Left            =   4440
      TabIndex        =   6
      Top             =   4020
      Width           =   2385
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   525
      Left            =   3690
      TabIndex        =   5
      Top             =   4410
      Width           =   1245
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   525
      Left            =   1800
      TabIndex        =   4
      Top             =   4410
      Width           =   1245
   End
   Begin VB.TextBox txtPath 
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6855
   End
   Begin VB.FileListBox fleFile 
      Height          =   3330
      Left            =   3450
      TabIndex        =   2
      Top             =   330
      Width           =   3405
   End
   Begin VB.DirListBox dirDirectory 
      Height          =   3660
      Left            =   0
      TabIndex        =   1
      Top             =   660
      Width           =   3435
   End
   Begin VB.DriveListBox driDrive 
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   330
      Width           =   3435
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "File Pattern"
      Height          =   180
      Left            =   3450
      TabIndex        =   9
      Top             =   4080
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "File Name"
      Height          =   180
      Left            =   3480
      TabIndex        =   8
      Top             =   3750
      Width           =   870
   End
End
Attribute VB_Name = "frmRank_Load"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbPattern_Click()

    Me.fleFile.Pattern = Me.cmbPattern.Text
    Me.fleFile.Refresh
    
End Sub

Private Sub cmdClose_Click()

    Unload Me
    
End Sub

Private Sub cmdLoad_Click()

    Dim arrGrade(1 To 50)       As String
    
    Dim strPath                 As String
    Dim strFileName             As String
    Dim strTemp                 As String
    Dim strSection              As String
    Dim strDefect_Code          As String
        
    Dim intFileNum              As Integer
    Dim intPos                  As Integer
    Dim intArray_Index          As Integer
    Dim intCol                  As Integer
    Dim intGrade_Col            As Integer
    Dim intRow                  As Integer
    
    strPath = App.PATH & "\Env\"
    strFileName = "Last_Path.dat"
    intFileNum = FreeFile
    
    Open strPath & strFileName For Output As intFileNum
    
    strTemp = Me.dirDirectory.PATH
    If Right(strTemp, 1) <> "\" Then
        strTemp = strTemp & "\"
    End If
    
    Print #intFileNum, strTemp
    
    Close intFileNum
    
    strPath = Me.txtPath.Text
    If Right(strPath, 1) <> "\" Then
        strPath = strPath & "\"
    End If
    
    strFileName = Me.txtFileName.Text
    
    If Dir(strPath & strFileName, vbNormal) <> "" Then
        intFileNum = FreeFile
        frmRank_Interface.flxGrade.Rows = 1
        Open strPath & strFileName For Input As intFileNum
        
        frmRank_Interface.flxRank_Data.Visible = False
        While Not EOF(intFileNum)
            Line Input #intFileNum, strTemp
            
            intPos = InStr(strTemp, ",")
            If intPos > 0 Then
                strTemp = Mid(strTemp, intPos + 1)
                
                intPos = InStr(strTemp, ",")
                If InStr(Left(strTemp, intPos - 1), "CODE") = 0 Then
                    With frmRank_Interface.flxRank_Data
                        strDefect_Code = Left(strTemp, intPos - 1)
                        Select Case Mid(strDefect_Code, 2, 1)
                        Case "D":
                            intRow = Add_Rank_Grid("POINT D/F")
                        Case "L":
                            intRow = Add_Rank_Grid("LINE D/F")
                        Case "G":
                            intRow = Add_Rank_Grid("GAP D/F")
                        Case "M":
                            intRow = Add_Rank_Grid("MURA D/F")
                        Case "F":
                            intRow = Add_Rank_Grid("CF D/F")
                        Case "P":
                            intRow = Add_Rank_Grid("POLARIZE D/F")
                        Case "A":
                            intRow = Add_Rank_Grid("APPEARANCE D/F")
                        Case "C":
                            intRow = Add_Rank_Grid("CELL D/F")
                        Case "O":
                            intRow = Add_Rank_Grid("OTHER D/F")
                        End Select
                        .TextMatrix(intRow, 1) = strDefect_Code
                        strTemp = Mid(strTemp, intPos + 1)
                        
                        For intCol = 2 To (.Cols - 2)
                            intPos = InStr(strTemp, ",")
                            .TextMatrix(intRow, intCol) = Left(strTemp, intPos - 1)
                            strTemp = Mid(strTemp, intPos + 1)
                            DoEvents
                        Next intCol
                        .TextMatrix(intRow, .Cols - 1) = strTemp
                    End With
                Else
                    strTemp = Mid(strTemp, intPos + 1)
                    intCol = 1
                    
                    intPos = InStr(strTemp, ",")
                    While UCase(Left(strTemp, intPos - 1)) <> "PRIORITY"
                        strTemp = Mid(strTemp, intPos + 1)
                        intCol = intCol + 1
                        intPos = InStr(strTemp, ",")
                    Wend
                    strTemp = Mid(strTemp, intPos + 1)
                    intCol = intCol + 1
                    
                    intPos = InStr(strTemp, ",")
                    intArray_Index = 0
                    While intPos > 0
                        intArray_Index = intArray_Index + 1
                        arrGrade(intArray_Index) = Left(strTemp, intPos - 1)
                        strTemp = Mid(strTemp, intPos + 1)
                        intPos = InStr(strTemp, ",")
                    Wend
                    intArray_Index = intArray_Index + 1
                    arrGrade(intArray_Index) = strTemp
                    
                    With frmRank_Interface
                        .flxRank_Data.Cols = intArray_Index + intCol + 1
                        For intGrade_Col = (intCol + 1) To (intArray_Index + intCol)
                            .flxRank_Data.TextMatrix(0, intGrade_Col) = arrGrade(intGrade_Col - intCol)
                            .flxRank_Data.Row = 0
                            .flxRank_Data.Col = intGrade_Col
                            .flxRank_Data.CellAlignment = flexAlignCenterCenter
                        Next intGrade_Col
                    End With
                End If
            End If
        Wend
        
        Close intFileNum
        frmRank_Interface.flxRank_Data.Visible = True
        
        Unload Me
    Else
        Call MsgBox(strFileName & " does not exist", vbOKOnly, "File error")
    End If
    
End Sub

Private Sub dirDirectory_Change()

    Dim strPath                 As String
    
    strPath = Me.dirDirectory.PATH
    
    If Right(strPath, 1) <> "\" Then
        strPath = strPath & "\"
    End If
    Me.fleFile.Pattern = Me.cmbPattern.Text
    Me.fleFile.PATH = strPath
    Me.txtPath.Text = strPath
    
End Sub

Private Sub driDrive_Change()

    Dim strDrive                As String
    
    Dim intPos                  As Integer
    
    strDrive = Me.driDrive.Drive
    intPos = InStr(strDrive, ":")
    strDrive = Left(strDrive, intPos)
    
    If Right(strDrive, 1) <> "\" Then
        strDrive = strDrive & "\"
    End If
    Me.dirDirectory.PATH = strDrive
    Me.fleFile.Pattern = Me.cmbPattern.Text
    Me.fleFile.PATH = strDrive
    Me.txtPath.Text = strDrive
    
End Sub

Private Sub fleFile_Click()

    Dim strFileName             As String
    
    strFileName = Me.fleFile.FILENAME
    
    Me.txtFileName.Text = strFileName
    
End Sub

Private Sub Form_Load()

    Dim strPath                     As String
    Dim strFileName                 As String
    Dim strTemp                     As String
    
    Dim intFileNum                  As Integer
    Dim intPos                      As Integer
    
    Me.cmbPattern.Clear
    Me.cmbPattern.AddItem "*.ran"
    Me.cmbPattern.AddItem "*.*"
    Me.cmbPattern.Text = "*.ran"
    
    strPath = App.PATH & "\Env\"
    strFileName = "Last_Path.dat"
    
    If Dir(strPath & strFileName, vbNormal) <> "" Then
        intFileNum = FreeFile
        
        Open strPath & strFileName For Input As intFileNum
        
        Line Input #intFileNum, strTemp
        
        Close intFileNum
        
        If Right(strTemp, 1) <> "\" Then
            strTemp = strTemp & "\"
        End If
        intPos = InStr(strTemp, ":")
        If intPos > 0 Then
            Me.driDrive.Drive = Left(strTemp, intPos)
        End If
        Me.dirDirectory.PATH = strTemp
        Me.fleFile.Pattern = Me.cmbPattern.Text
        Me.fleFile.PATH = strTemp
    Else
        strTemp = App.PATH
        intPos = InStr(strTemp, ":")
        If intPos > 0 Then
            Me.driDrive.Drive = Left(strTemp, intPos)
        End If
        Me.dirDirectory.PATH = strTemp
        Me.fleFile.Pattern = Me.cmbPattern.Text
        Me.fleFile.PATH = strTemp
    End If
    
    Me.txtPath.Text = Me.dirDirectory.PATH
    
End Sub

Private Function Add_Rank_Grid(ByVal pData As String) As Integer

    Dim intRow          As Integer
    Dim intCol          As Integer
    
    With frmRank_Interface.flxRank_Data
        intRow = .Rows
        .AddItem pData
        .RowHeight(intRow) = 350
        For intCol = 0 To .Cols - 1
            .Row = intRow
            .Col = intCol
            .CellAlignment = flexAlignCenterCenter
        Next intCol
        
        Add_Rank_Grid = .Rows - 1
    End With

End Function
