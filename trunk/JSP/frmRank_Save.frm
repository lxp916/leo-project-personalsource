VERSION 5.00
Begin VB.Form frmRank_Save 
   Caption         =   "File Save"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   6840
   StartUpPosition =   2  '화면 가운데
   Begin VB.DriveListBox driDrive 
      Height          =   300
      Left            =   0
      TabIndex        =   7
      Top             =   330
      Width           =   3435
   End
   Begin VB.DirListBox dirDirectory 
      Height          =   3660
      Left            =   0
      TabIndex        =   6
      Top             =   660
      Width           =   3435
   End
   Begin VB.FileListBox fleFile 
      Height          =   3330
      Left            =   3450
      TabIndex        =   5
      Top             =   330
      Width           =   3405
   End
   Begin VB.TextBox txtPath 
      Height          =   300
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   525
      Left            =   1800
      TabIndex        =   3
      Top             =   4410
      Width           =   1245
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   525
      Left            =   3690
      TabIndex        =   2
      Top             =   4410
      Width           =   1245
   End
   Begin VB.ComboBox cmbPattern 
      Height          =   300
      Left            =   4440
      TabIndex        =   1
      Top             =   4020
      Width           =   2385
   End
   Begin VB.TextBox txtFileName 
      Height          =   300
      Left            =   4440
      TabIndex        =   0
      Top             =   3690
      Width           =   2385
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "File Name"
      Height          =   180
      Left            =   3480
      TabIndex        =   9
      Top             =   3750
      Width           =   870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "File Pattern"
      Height          =   180
      Left            =   3450
      TabIndex        =   8
      Top             =   4080
      Width           =   945
   End
End
Attribute VB_Name = "frmRank_Save"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()

    Unload Me
    
End Sub

Private Sub cmdSave_Click()

    Dim strPath                 As String
    Dim strFileName             As String
    Dim strTemp                 As String
    
    Dim intFileNum              As Integer
    
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
    If InStr(strFileName, "*") = 0 Then
        If Dir(strPath & strFileName, vbNormal) <> "" Then
            If MsgBox(strFileName & " already exist. Overwrite it?", vbYesNo, "File save") = vbYes Then
                Call Save_File(strPath, strFileName)
                
                Unload Me
            End If
        Else
            Call Save_File(strPath, strFileName)
            
            Unload Me
        End If
    Else
        Call MsgBox("Check file name.", vbOKOnly, "File save")
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
    Me.txtFileName.Text = "*.ran"
    
End Sub

Private Sub Save_File(ByVal pPath As String, ByVal pFileName As String)

    Dim strTemp                     As String
    
    Dim intFileNum                  As Integer
    Dim intRow                      As Integer
    Dim intCol                      As Integer
       
    intFileNum = FreeFile
    Open pPath & pFileName For Output As intFileNum
    
    With frmRank_Interface.flxRank_Data
        If .Rows > 1 Then
            For intRow = 0 To .Rows - 1
                strTemp = .TextMatrix(intRow, 0) & ","
                For intCol = 1 To .Cols - 2
                    strTemp = strTemp & .TextMatrix(intRow, intCol) & ","
                Next intCol
                strTemp = strTemp & .TextMatrix(intRow, .Cols - 1)
                Print #intFileNum, strTemp
                DoEvents
            Next intRow
        End If
    End With
    
    Close intFileNum
    
End Sub
