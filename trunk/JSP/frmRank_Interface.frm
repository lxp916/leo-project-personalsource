VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRank_Interface 
   Caption         =   "RANK Intterface"
   ClientHeight    =   11112
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   20256
   LinkTopic       =   "Form1"
   ScaleHeight     =   11112
   ScaleWidth      =   20256
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   525
      Left            =   7770
      TabIndex        =   31
      Top             =   10560
      Width           =   1245
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   525
      Left            =   9330
      TabIndex        =   28
      Top             =   10560
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Height          =   10515
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   20235
      Begin VB.CommandButton cmdData_Delete 
         Caption         =   "DELETE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1125
         Left            =   18540
         TabIndex        =   37
         Top             =   9330
         Width           =   1605
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   525
         Left            =   16770
         TabIndex        =   36
         Top             =   9660
         Width           =   1245
      End
      Begin VB.CommandButton cmdRank_Add 
         Caption         =   "Modify"
         Height          =   525
         Left            =   16770
         TabIndex        =   35
         Top             =   9150
         Width           =   1245
      End
      Begin MSFlexGridLib.MSFlexGrid flxGrade 
         Height          =   1575
         Left            =   14370
         TabIndex        =   32
         Top             =   8640
         Width           =   2055
         _ExtentX        =   3620
         _ExtentY        =   2773
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
      End
      Begin VB.TextBox txtGrade_Rank 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   17160
         TabIndex        =   30
         Top             =   8730
         Width           =   975
      End
      Begin VB.CommandButton cmdRegist 
         Caption         =   "REGIST"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1125
         Left            =   18540
         TabIndex        =   29
         Top             =   8220
         Width           =   1605
      End
      Begin VB.TextBox txtPriority 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   15300
         TabIndex        =   26
         Top             =   8280
         Width           =   675
      End
      Begin VB.TextBox txtRank 
         Alignment       =   2  'Center
         Height          =   900
         Index           =   0
         Left            =   990
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   9300
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.ComboBox cmbAddress_Count 
         Height          =   300
         Left            =   13470
         TabIndex        =   22
         Top             =   8670
         Width           =   735
      End
      Begin VB.TextBox txtAccumulation 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   13470
         TabIndex        =   20
         Top             =   8280
         Width           =   735
      End
      Begin VB.ComboBox cmbDetail_Division 
         Height          =   300
         Left            =   10950
         TabIndex        =   18
         Top             =   8670
         Width           =   735
      End
      Begin VB.ComboBox cmbDefect_Type 
         Height          =   300
         Left            =   8490
         TabIndex        =   17
         Top             =   8280
         Width           =   735
      End
      Begin VB.ComboBox cmbXY_Axis 
         Height          =   300
         Left            =   10950
         TabIndex        =   15
         Top             =   8280
         Width           =   735
      End
      Begin VB.ComboBox cmbJudge 
         Height          =   300
         Left            =   8490
         TabIndex        =   13
         Top             =   8670
         Width           =   735
      End
      Begin VB.TextBox txtDefect_Name_China 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   5490
         TabIndex        =   10
         Top             =   8670
         Width           =   1755
      End
      Begin VB.TextBox txtDefect_Name_English 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   5490
         TabIndex        =   8
         Top             =   8280
         Width           =   1755
      End
      Begin VB.TextBox txtDefect_Code 
         Alignment       =   2  'Center
         Height          =   300
         Left            =   1650
         MaxLength       =   5
         TabIndex        =   6
         Top             =   8670
         Width           =   1755
      End
      Begin VB.ComboBox cmbSection 
         Height          =   300
         Left            =   1170
         TabIndex        =   4
         Top             =   8280
         Width           =   1755
      End
      Begin MSFlexGridLib.MSFlexGrid flxRank_Data 
         Height          =   7995
         Left            =   90
         TabIndex        =   2
         Top             =   180
         Width           =   20055
         _ExtentX        =   35370
         _ExtentY        =   14097
         _Version        =   393216
         Rows            =   1
         Cols            =   25
         FixedCols       =   0
      End
      Begin VB.Label lblGrade 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   17160
         TabIndex        =   34
         Top             =   8280
         Width           =   585
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "GRADE"
         Height          =   180
         Left            =   16470
         TabIndex        =   33
         Top             =   8340
         Width           =   615
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "RANK"
         Height          =   180
         Left            =   16530
         TabIndex        =   27
         Top             =   8790
         Width           =   495
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "PRIORITY"
         Height          =   180
         Left            =   14400
         TabIndex        =   25
         Top             =   8340
         Width           =   825
      End
      Begin VB.Label lblRank 
         AutoSize        =   -1  'True
         Caption         =   "Y RANK"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   23
         Top             =   9360
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ADDRESS COUNT"
         Height          =   180
         Left            =   11880
         TabIndex        =   21
         Top             =   8730
         Width           =   1545
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "ACCUMULATION"
         Height          =   180
         Left            =   11940
         TabIndex        =   19
         Top             =   8340
         Width           =   1455
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "DETAIL DIVISION"
         Height          =   180
         Left            =   9420
         TabIndex        =   16
         Top             =   8730
         Width           =   1455
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "X/Y AXIS"
         Height          =   180
         Left            =   10080
         TabIndex        =   14
         Top             =   8340
         Width           =   795
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "JUDGE"
         Height          =   180
         Left            =   7830
         TabIndex        =   12
         Top             =   8730
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "D/F TYPE"
         Height          =   180
         Left            =   7560
         TabIndex        =   11
         Top             =   8340
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "DEFECT NAME(CHN)"
         Height          =   180
         Left            =   3510
         TabIndex        =   9
         Top             =   8760
         Width           =   1860
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "DEFECT NAME(ENG)"
         Height          =   180
         Left            =   3510
         TabIndex        =   7
         Top             =   8340
         Width           =   1860
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "DEFECT CODE"
         Height          =   180
         Left            =   270
         TabIndex        =   5
         Top             =   8730
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "SECTION"
         Height          =   180
         Left            =   270
         TabIndex        =   3
         Top             =   8340
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   525
      Left            =   10890
      TabIndex        =   0
      Top             =   10560
      Width           =   1245
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuGrade_List_Input 
         Caption         =   "Grade List Input"
      End
   End
End
Attribute VB_Name = "frmRank_Interface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mCURRENT_COL                As Integer
Dim mCURRENT_ROW                As Integer

Private Sub cmbAddress_Count_Change()

    Dim intIndex                As Integer
    
    Dim bolFind                 As Boolean
    
    bolFind = False
    With Me.cmbAddress_Count
        For intIndex = 0 To .ListCount - 1
            If .Text = .List(intIndex) Then
                bolFind = True
            End If
        Next intIndex
        
        If bolFind = False Then
            .Text = ""
        Else
            If Me.txtDefect_Code.Text = "" Then
                Call MsgBox("Please key in defect code.", vbOKOnly, "Data error")
                Me.cmbAddress_Count.Text = ""
            Else
                Select Case Me.cmbAddress_Count.Text
                Case "X":
'                    If (Mid(Me.txtDefect_Code.Text, 2, 1) = "D") Or (Mid(Me.txtDefect_Code.Text, 2, 1) = "L") Or (Mid(Me.txtDefect_Code.Text, 2, 1) = "M") Then
'                        Call MsgBox("Defect code type and address count does not match", vbOKOnly, "Data error")
'                        Me.cmbAddress_Count.Text = ""
'                    End If
                Case "1":
                    If Mid(Me.txtDefect_Code.Text, 2, 1) <> "D" Then
                        Call MsgBox("Defect code type and address count does not match", vbOKOnly, "Data error")
                        Me.cmbAddress_Count.Text = ""
                    End If
                Case "2":
                    If Mid(Me.txtDefect_Code.Text, 2, 1) <> "L" Then
                        Call MsgBox("Defect code type and address count does not match", vbOKOnly, "Data error")
                        Me.cmbAddress_Count.Text = ""
                    End If
                Case "3":
                    If (Mid(Me.txtDefect_Code.Text, 2, 1) = "D") Or (Mid(Me.txtDefect_Code.Text, 2, 1) = "L") Then
                        Call MsgBox("Defect code type and address count does not match", vbOKOnly, "Data error")
                        Me.cmbAddress_Count.Text = ""
                    End If
                End Select
            End If
        End If
    End With
    
End Sub

Private Sub cmbAddress_Count_Click()

    Dim intIndex                As Integer
    
    Dim bolFind                 As Boolean
    
    bolFind = False
    With Me.cmbAddress_Count
        For intIndex = 0 To .ListCount - 1
            If .Text = .List(intIndex) Then
                bolFind = True
            End If
        Next intIndex
        
        If bolFind = False Then
            .Text = ""
        Else
            If Me.txtDefect_Code.Text = "" Then
                Call MsgBox("Please key in defect code.", vbOKOnly, "Data error")
                Me.cmbAddress_Count.Text = ""
            Else
                Select Case Me.cmbAddress_Count.Text
                Case "X":
'                    If (Mid(Me.txtDefect_Code.Text, 2, 1) = "D") Or (Mid(Me.txtDefect_Code.Text, 2, 1) = "L") Or (Mid(Me.txtDefect_Code.Text, 2, 1) = "M") Then
'                        Call MsgBox("Defect code type and address count does not match", vbOKOnly, "Data error")
'                        Me.cmbAddress_Count.Text = ""
'                    End If
                Case "1":
                    If Mid(Me.txtDefect_Code.Text, 2, 1) <> "D" Then
                        Call MsgBox("Defect code type and address count does not match", vbOKOnly, "Data error")
                        Me.cmbAddress_Count.Text = ""
                    End If
                Case "2":
                    If Mid(Me.txtDefect_Code.Text, 2, 1) <> "L" Then
                        Call MsgBox("Defect code type and address count does not match", vbOKOnly, "Data error")
                        Me.cmbAddress_Count.Text = ""
                    End If
                Case "3":
                    If (Mid(Me.txtDefect_Code.Text, 2, 1) = "D") Or (Mid(Me.txtDefect_Code.Text, 2, 1) = "L") Then
                        Call MsgBox("Defect code type and address count does not match", vbOKOnly, "Data error")
                        Me.cmbAddress_Count.Text = ""
                    End If
                End Select
            End If
        End If
    End With

End Sub

Private Sub cmbDefect_Type_Change()

    Dim intIndex                As Integer
    
    Dim bolFind                 As Boolean
    
    bolFind = False
    With Me.cmbDefect_Type
        For intIndex = 0 To .ListCount - 1
            If .Text = .List(intIndex) Then
                bolFind = True
            End If
        Next intIndex
        
        If bolFind = False Then
            .Text = ""
        End If
    End With
    
End Sub

Private Sub cmbDefect_Type_Click()

    If Me.txtDefect_Code.Text <> "" Then
        If Me.cmbDefect_Type.Text = "P" Then
            If Mid(Me.txtDefect_Code.Text, 2, 1) <> "D" Then
                Call MsgBox("Defect code and type does not match", vbOKOnly, "Data error")
            End If
        End If
    Else
        Call MsgBox("Please key in defect code.", vbOKOnly, "Data error")
        Me.cmbDefect_Type.Text = ""
    End If
    
End Sub

Private Sub cmbDetail_Division_Change()

    Dim intIndex                As Integer
    
    Dim bolFind                 As Boolean
    
    bolFind = False
    With Me.cmbDetail_Division
        For intIndex = 0 To .ListCount - 1
            If .Text = .List(intIndex) Then
                bolFind = True
            End If
        Next intIndex
        
        If bolFind = False Then
            .Text = ""
        Else
            If Me.txtDefect_Code.Text <> "" Then
                If Me.cmbDetail_Division.Text = "X" Then
                    Me.txtAccumulation.Text = "X"
                Else
                    If Mid(Me.txtDefect_Code.Text, 2, 1) <> "D" Then
                        Call MsgBox("Defect code and detail division does not match", vbOKOnly, "Data error")
                    End If
                End If
            Else
                Call MsgBox("Please key in defect code.", vbOKOnly, "Data error")
                Me.cmbDefect_Type.Text = ""
            End If
        End If
    End With
    
End Sub

Private Sub cmbDetail_Division_Click()

    If Me.txtDefect_Code.Text <> "" Then
        If Me.cmbDetail_Division.Text = "X" Then
            Me.txtAccumulation.Text = "X"
        Else
            If Mid(Me.txtDefect_Code.Text, 2, 1) <> "D" Then
                Call MsgBox("Defect code and detail division does not match", vbOKOnly, "Data error")
            End If
        End If
    Else
        Call MsgBox("Please key in defect code.", vbOKOnly, "Data error")
        Me.cmbDefect_Type.Text = ""
    End If
    
End Sub

Private Sub cmbJudge_Change()

    Dim intIndex                As Integer
    
    Dim bolFind                 As Boolean
    
    bolFind = False
    With Me.cmbJudge
        For intIndex = 0 To .ListCount - 1
            If .Text = .List(intIndex) Then
                bolFind = True
            End If
        Next intIndex
        
        If bolFind = False Then
            .Text = ""
        End If
    End With
    
End Sub

Private Sub cmbSection_Change()

    Dim intIndex                As Integer
    
    Dim bolFind                 As Boolean
    
    bolFind = False
    With Me.cmbSection
        For intIndex = 0 To .ListCount - 1
            If .Text = .List(intIndex) Then
                bolFind = True
            End If
        Next intIndex
        
        If bolFind = False Then
            .Text = ""
        End If
    End With
    
End Sub

Private Sub cmbXY_Axis_Change()

    Dim intIndex                As Integer
    
    Dim bolFind                 As Boolean
    
    bolFind = False
    With Me.cmbXY_Axis
        For intIndex = 0 To .ListCount - 1
            If .Text = .List(intIndex) Then
                bolFind = True
            End If
        Next intIndex
        
        If bolFind = False Then
            .Text = ""
        End If
    End With
    
End Sub

Private Sub cmdClose_Click()

    Unload Me
    
End Sub

Private Sub cmdData_Delete_Click()

    Dim intRow                      As Integer
    Dim intIndex                    As Integer
    
    Dim bolFind                     As Boolean
    
    intIndex = 0
    bolFind = False
    If Me.flxRank_Data.Rows > 1 Then
        While bolFind = False
            intIndex = intIndex + 1
            If Me.flxRank_Data.TextMatrix(intIndex, 1) = Me.txtDefect_Code.Text Then
                intRow = intIndex
                bolFind = True
            Else
                If intIndex = Me.flxRank_Data.Rows - 1 Then
                    bolFind = False
                    intRow = 0
                End If
            End If
        Wend
        If intRow > 0 Then
            If Me.flxRank_Data.Rows = 2 Then
                Me.flxRank_Data.Rows = 1
            Else
                Me.flxGrade.RemoveItem (intRow)
            End If
        Else
            Call MsgBox("Rank data does not exist", vbOKOnly, "Data error")
        End If
    Else
        Call MsgBox("Rank data list does not exist", vbOKOnly, "Data error")
    End If
    
End Sub

Private Sub cmdDelete_Click()

    Dim intRow                      As Integer
    Dim intIndex                    As Integer
    
    If Me.lblGrade.Caption <> "" Then
        intRow = 0
        For intIndex = 1 To Me.flxGrade.Rows - 1
            If Me.flxGrade.TextMatrix(intIndex, 0) = Me.lblGrade.Caption Then
                intRow = intIndex
            End If
            If intRow > 0 Then
                Me.flxGrade.TextMatrix(intRow, 1) = ""
            End If
        Next intIndex
        Me.lblGrade.Caption = ""
        Me.txtGrade_Rank.Text = ""
    End If
    
End Sub

Private Sub cmdLoad_Click()

    Load frmRank_Load
    frmRank_Load.Show
    
End Sub

Private Sub cmdRank_Add_Click()

    Dim intRow                      As Integer
    Dim intIndex                    As Integer
    
    If Me.lblGrade.Caption <> "" Then
        intRow = 0
        For intIndex = 1 To Me.flxGrade.Rows - 1
            If Me.flxGrade.TextMatrix(intIndex, 0) = Me.lblGrade.Caption Then
                intRow = intIndex
            End If
        Next intIndex
        If intRow > 0 Then
            Me.flxGrade.TextMatrix(intRow, 1) = UCase(Me.txtGrade_Rank.Text)
        End If
        Me.lblGrade.Caption = ""
        Me.txtGrade_Rank.Text = ""
    End If
    
End Sub

Private Sub cmdRegist_Click()

    Dim intRow                      As Integer
    Dim intCol                      As Integer
    Dim intIndex                    As Integer
    
    Dim bolFind                     As Boolean
    
    '============Leo 2012.05.22 Add Rank Level Start
    Dim RankColBegin As Integer
    Dim RankColEnd As Integer
    RankColBegin = 10
    RankColEnd = RankColBegin + UBound(RankLevel)
    '============Leo 2012.05.22 Add Rank Level end
    
    If Check_Data = True Then
        intIndex = 0
        bolFind = False
        If Me.flxRank_Data.Rows > 1 Then
            While bolFind = False
                intIndex = intIndex + 1
                If Me.flxRank_Data.TextMatrix(intIndex, 1) = Me.txtDefect_Code.Text Then
                    intRow = intIndex
                    bolFind = True
                Else
                    If intIndex = Me.flxRank_Data.Rows - 1 Then
                        bolFind = True
                    End If
                End If
            Wend
        End If
        If intRow = 0 Then
            intRow = Add_Grid(Me.cmbSection.Text)
        End If
        With Me.flxRank_Data
            .TextMatrix(intRow, 1) = Me.txtDefect_Code.Text
            .TextMatrix(intRow, 2) = Me.txtDefect_Name_English.Text
            .TextMatrix(intRow, 3) = Me.txtDefect_Name_China.Text
            .TextMatrix(intRow, 4) = Me.cmbDefect_Type.Text
            .TextMatrix(intRow, 5) = Me.cmbJudge.Text
            .TextMatrix(intRow, 6) = Me.cmbXY_Axis.Text
            .TextMatrix(intRow, 7) = Me.cmbDetail_Division.Text
            .TextMatrix(intRow, 8) = Me.txtAccumulation.Text
            .TextMatrix(intRow, 9) = Me.cmbAddress_Count.Text
            
  '============Leo 2012.05.22 Add Rank Level Start
        For intIndex = 0 To UBound(RankLevel)
            If txtRank(intIndex).Text = "" Then
                .TextMatrix(intRow, RankColBegin + intIndex) = "-"
            Else
                .TextMatrix(intRow, RankColBegin + intIndex) = txtRank(intIndex).Text
            End If
        Next intIndex
     
'            If Me.txtRank_Y.Text = "" Then
'                .TextMatrix(intRow, 10) = "-"
'            Else
'                .TextMatrix(intRow, 10) = Me.txtRank_Y.Text
'            End If
'            If Me.txtRank_L.Text = "" Then
'                .TextMatrix(intRow, 11) = "-"
'            Else
'                .TextMatrix(intRow, 11) = Me.txtRank_L.Text
'            End If
'            If Me.txtRank_K.Text = "" Then
'                .TextMatrix(intRow, 12) = "-"
'            Else
'                .TextMatrix(intRow, 12) = Me.txtRank_K.Text
'            End If
'            If Me.txtRank_C.Text = "" Then
'                .TextMatrix(intRow, 13) = "-"
'            Else
'                .TextMatrix(intRow, 13) = Me.txtRank_C.Text
'            End If
'            If Me.txtRank_S.Text = "" Then
'                .TextMatrix(intRow, 14) = "-"
'            Else
'                .TextMatrix(intRow, 14) = Me.txtRank_S.Text
'            End If
'            .TextMatrix(intRow, 15) = Me.txtPriority.Text
            .TextMatrix(intRow, RankColEnd + 1) = Me.txtPriority.Text
'            For intCol = 16 To .Cols - 1
             For intCol = RankColEnd + 2 To .Cols - 1
'              If Me.flxGrade.TextMatrix(intCol - 15, 1) = "" Then
                If Me.flxGrade.TextMatrix(intCol - (RankColEnd + 1), 1) = "" Then
                    .TextMatrix(intRow, intCol) = "-"
                Else
'                    .TextMatrix(intRow, intCol) = Me.flxGrade.TextMatrix(intCol - 15, 1)
                    .TextMatrix(intRow, intCol) = Me.flxGrade.TextMatrix(intCol - (RankColEnd + 1), 1)
                End If
            Next intCol
             '============Leo 2012.05.22 Add Rank Level End
            Me.cmbSection.Text = ""
            Me.txtDefect_Code.Text = ""
            Me.txtDefect_Name_English.Text = ""
            Me.txtDefect_Name_China.Text = ""
            Me.cmbDefect_Type = ""
            Me.cmbJudge.Text = ""
            Me.cmbXY_Axis.Text = ""
            Me.cmbDetail_Division.Text = ""
            Me.txtAccumulation.Text = ""
            Me.cmbAddress_Count.Text = ""
  '============Leo 2012.05.22 Add Rank Level Start
        For intIndex = 0 To UBound(RankLevel)
            txtRank(intIndex).Text = ""
        Next intIndex
      '============Leo 2012.05.22 Add Rank Level End
            Me.txtPriority.Text = ""
            For intRow = 1 To Me.flxGrade.Rows - 1
                Me.flxGrade.TextMatrix(intRow, 1) = ""
            Next intRow
        End With
    Else
        Call MsgBox("Check input data", vbOKOnly, "Data error")
    End If
    
End Sub

Private Sub cmdSave_Click()

    Load frmRank_Save
    frmRank_Save.Show
    
End Sub

Private Sub flxGrade_Click()

    Dim intRow                      As Integer
    
    intRow = Me.flxGrade.Row
    
    If intRow > 0 Then
        Me.lblGrade.Caption = Me.flxGrade.TextMatrix(intRow, 0)
        Me.txtGrade_Rank.Text = Me.flxGrade.TextMatrix(intRow, 1)
    End If
    
End Sub

Private Sub flxRank_Data_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim intCol                  As Integer
    Dim intRow                  As Integer
    
    Select Case KeyCode
    Case vbKeyUp:
        If Me.flxRank_Data.Rows > 22 Then
            If mCURRENT_ROW > 1 Then
                mCURRENT_ROW = mCURRENT_ROW - 1
                Me.flxRank_Data.TopRow = mCURRENT_ROW
                Me.flxRank_Data.Row = mCURRENT_ROW
            Else
                Me.flxRank_Data.Row = Me.flxRank_Data.Row - 1
            End If
        Else
            Me.flxRank_Data.Row = Me.flxRank_Data.Row - 1
        End If
    Case vbKeyDown:
        If Me.flxRank_Data.Rows > 22 Then
            If mCURRENT_ROW + 22 < Me.flxRank_Data.Rows Then
                mCURRENT_ROW = mCURRENT_ROW + 1
                Me.flxRank_Data.TopRow = mCURRENT_ROW
                Me.flxRank_Data.Row = mCURRENT_ROW
            Else
                Me.flxRank_Data.Row = Me.flxRank_Data.Row + 1
            End If
        Else
            Me.flxRank_Data.Row = Me.flxRank_Data.Row + 1
        End If
    Case vbKeyRight:
        If mCURRENT_COL + 13 < Me.flxRank_Data.Cols Then
            mCURRENT_COL = mCURRENT_COL + 1
            Me.flxRank_Data.LeftCol = mCURRENT_COL
            Me.flxGrade.Col = mCURRENT_COL
        Else
            Me.flxRank_Data.Col = Me.flxRank_Data.Col + 1
        End If
    Case vbKeyLeft:
        If mCURRENT_COL > 0 Then
            mCURRENT_COL = mCURRENT_COL - 1
            Me.flxRank_Data.LeftCol = mCURRENT_COL
            Me.flxGrade.Col = mCURRENT_COL
        Else
            Me.flxRank_Data.Col = Me.flxRank_Data.Col - 1
        End If
    Case vbKeyPageDown:
        If Me.flxRank_Data.Rows > 22 Then
            If mCURRENT_ROW + 44 < Me.flxRank_Data.Rows Then
                mCURRENT_ROW = mCURRENT_ROW + 22
                Me.flxRank_Data.TopRow = mCURRENT_ROW
            Else
                mCURRENT_ROW = Me.flxRank_Data.Rows - 21
                Me.flxRank_Data.TopRow = mCURRENT_ROW
            End If
            Me.flxRank_Data.Row = mCURRENT_ROW
        End If
    Case vbKeyPageUp:
        If Me.flxRank_Data.Rows > 22 Then
            If mCURRENT_ROW - 22 > 1 Then
                mCURRENT_ROW = mCURRENT_ROW - 21
                Me.flxRank_Data.TopRow = mCURRENT_ROW
            ElseIf mCURRENT_ROW < 22 Then
                mCURRENT_ROW = 1
                Me.flxRank_Data.TopRow = mCURRENT_ROW
            End If
            Me.flxRank_Data.Row = mCURRENT_ROW
        End If
    End Select
    
End Sub

Private Sub flxRank_Data_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim intRow                      As Integer
    Dim intCol                      As Integer
    Dim intGrade_Row                As Integer
     '============Leo 2012.05.22 Add Rank Level Start
    Dim intIndex As Integer
    Dim RankColBegin As Integer
    Dim RankColEnd As Integer
    RankColBegin = 10
    RankColEnd = RankColBegin + UBound(RankLevel)
    '============Leo 2012.05.22 Add Rank Level end
    
    If Button = vbRightButton Then
        Me.PopupMenu Me.mnuMenu
    Else
        intRow = Me.flxRank_Data.Row
        If intRow > 0 Then
            With Me.flxRank_Data
                Me.cmbSection.Text = .TextMatrix(intRow, 0)
                Me.txtDefect_Code.Text = .TextMatrix(intRow, 1)
                Me.txtDefect_Name_English.Text = .TextMatrix(intRow, 2)
                Me.txtDefect_Name_China.Text = .TextMatrix(intRow, 3)
                Me.cmbDefect_Type.Text = .TextMatrix(intRow, 4)
                Me.cmbJudge.Text = .TextMatrix(intRow, 5)
                Me.cmbXY_Axis.Text = .TextMatrix(intRow, 6)
                Me.cmbDetail_Division.Text = .TextMatrix(intRow, 7)
                Me.txtAccumulation.Text = .TextMatrix(intRow, 8)
                Me.cmbAddress_Count.Text = .TextMatrix(intRow, 9)
                 '============Leo 2012.05.22 Add Rank Level Start
                 For intIndex = 0 To UBound(RankLevel)
                    txtRank(intIndex).Text = .TextMatrix(intRow, RankColBegin + intIndex)
                 Next intIndex
                 Me.txtPriority.Text = .TextMatrix(intRow, RankColEnd + 1)
'                Me.txtRank_Y.Text = .TextMatrix(intRow, 10)
'                Me.txtRank_L.Text = .TextMatrix(intRow, 11)
'                Me.txtRank_K.Text = .TextMatrix(intRow, 12)
'                Me.txtRank_C.Text = .TextMatrix(intRow, 13)
'                Me.txtRank_S.Text = .TextMatrix(intRow, 14)
'                Me.txtPriority.Text = .TextMatrix(intRow, 15)
                Me.flxGrade.Rows = 1
'                For intCol = 16 To .Cols - 1
                For intCol = RankColEnd + 2 To .Cols - 1
                    Call Add_Grade_Grid(.TextMatrix(0, intCol))
                    intGrade_Row = Me.flxGrade.Rows - 1
                    Me.flxGrade.TextMatrix(intGrade_Row, 1) = .TextMatrix(intRow, intCol)
                Next intCol
                 '============Leo 2012.05.22 Add Rank Level end
            End With
        End If
    End If
    
End Sub

Private Sub Form_Load()

    Call Init_Form
    Call Init_Grid
     '============Leo 2012.05.22 Add Rank Level
    Call Init_RankLevel
    mCURRENT_COL = 0
    mCURRENT_ROW = 1
    
    Me.Left = 1
    Me.Top = 1
    Me.Width = 20370
    Me.Height = 11535
    
End Sub

Private Sub Init_Form()

    Dim strPath                     As String
    Dim strFileName                 As String
    Dim strTemp                     As String
    
    Dim intFileNum                  As Integer
    
    With Me
        .cmbSection.Clear
        .cmbSection.AddItem "POINT D/F"
        .cmbSection.AddItem "LINE D/F"
        .cmbSection.AddItem "GAP D/F"
        .cmbSection.AddItem "MURA D/F"
        .cmbSection.AddItem "CF D/F"
        .cmbSection.AddItem "POLARIZE D/F"
        .cmbSection.AddItem "APPEARANCE D/F"
        .cmbSection.AddItem "CELL D/F"
        .cmbSection.AddItem "OTHER D/F"
        
        .cmbDefect_Type.Clear
        .cmbDefect_Type.AddItem "P"
        .cmbDefect_Type.AddItem "R"
        .cmbDefect_Type.AddItem "A"
        
        .cmbJudge.Clear
        .cmbJudge.AddItem "O"
        .cmbJudge.AddItem "X"
        
        .cmbXY_Axis.Clear
        .cmbXY_Axis.AddItem "O"
        .cmbXY_Axis.AddItem "X"
        
        .cmbDetail_Division.Clear
        .cmbDetail_Division.AddItem "B"
        .cmbDetail_Division.AddItem "D"
        .cmbDetail_Division.AddItem "LB"
        .cmbDetail_Division.AddItem "LD"
        .cmbDetail_Division.AddItem "TB"
        .cmbDetail_Division.AddItem "TD"
        .cmbDetail_Division.AddItem "L2"
        .cmbDetail_Division.AddItem "T2"
        .cmbDetail_Division.AddItem "X"
        
        .cmbAddress_Count.Clear
        .cmbAddress_Count.AddItem "X"
        .cmbAddress_Count.AddItem "1"
        .cmbAddress_Count.AddItem "2"
        .cmbAddress_Count.AddItem "3"
        
        strPath = App.PATH & "\Env\"
        strFileName = "Rank_Interface_Base.cfg"
        
        If Dir(strPath & strFileName, vbNormal) <> "" Then
            intFileNum = FreeFile
            Me.flxGrade.Rows = 1
            Open strPath & strFileName For Input As intFileNum
            
            While Not EOF(intFileNum)
                Line Input #intFileNum, strTemp
                
                Call Add_Grade_Grid(strTemp)
            Wend
            
            Close intFileNum
        End If
        
    End With
    
End Sub

Private Sub Init_Grid()

    Dim intCol                      As Integer
    Dim intRow                      As Integer
    Dim intLoopCount                As Integer
    '============Leo 2012.05.22 Add Rank Level Start
    Dim RankColBegin As Integer
    Dim RankColEnd As Integer
    RankColBegin = 10
    RankColEnd = RankColBegin + UBound(RankLevel)
    '============Leo 2012.05.22 Add Rank Level end
    
    With Me.flxRank_Data
    '.Cols = Me.flxGrade.Rows + 15
        .Cols = Me.flxGrade.Rows + RankColEnd + 1
        For intRow = 0 To .Rows - 1
            For intCol = 0 To .Cols - 1
                .Row = intRow
                .Col = intCol
                .CellAlignment = flexAlignCenterCenter
                .RowHeight(intRow) = 350
            Next intCol
        Next intRow
        
        .ColWidth(0) = 1000
        .ColWidth(1) = 1500
        .ColWidth(2) = 1900
        .ColWidth(3) = 1900
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColWidth(6) = 1200
        .ColWidth(7) = 1900
        .ColWidth(8) = 1700
        .ColWidth(9) = 1700
        '============Leo 2012.05.22 Add Rank Level Start
'        .ColWidth(10) = 1500
'        .ColWidth(11) = 1500
'        .ColWidth(12) = 1500
'        .ColWidth(13) = 1500
'        .ColWidth(14) = 1500
'        .ColWidth(15) = 1000
        .ColWidth(RankColEnd + 1) = 1000
         '============Leo 2012.05.22 Add Rank Level End
        
        .TextMatrix(0, 0) = "SECTION"
        .TextMatrix(0, 1) = "DEFECT CODE"
        .TextMatrix(0, 2) = "D/F NAME (English)"
        .TextMatrix(0, 3) = "D/F NAME (Chinese)"
        .TextMatrix(0, 4) = "D/F TYPE"
        .TextMatrix(0, 5) = "JUDGE"
        .TextMatrix(0, 6) = "X/Y AXIS"
        .TextMatrix(0, 7) = "DETAIL DIVISION"
        .TextMatrix(0, 8) = "ACCUMULATION"
        .TextMatrix(0, 9) = "ADDRESS COUNT"
        '============Leo 2012.05.22 Add Rank Level Start
         For intLoopCount = 0 To UBound(RankLevel)
             .ColWidth(RankColBegin + intLoopCount) = 1500
            .TextMatrix(0, RankColBegin + intLoopCount) = RankLevel(intLoopCount)
         Next intLoopCount
'        .TextMatrix(0, 10) = "Y"
'        .TextMatrix(0, 11) = "L"
'        .TextMatrix(0, 12) = "K"
'        .TextMatrix(0, 13) = "C"
'        .TextMatrix(0, 14) = "S"
'        .TextMatrix(0, 15) = "PRIORITY"
        .TextMatrix(0, RankColEnd + 1) = "PRIORITY"
        
        intLoopCount = Me.flxGrade.Rows - 1
        For intCol = 0 To intLoopCount - 1
'            .TextMatrix(0, intCol + 16) = Me.flxGrade.TextMatrix(intCol + 1, 0)
            .TextMatrix(0, intCol + RankColEnd + 2) = Me.flxGrade.TextMatrix(intCol + 1, 0)
        Next intCol
        '============Leo 2012.05.22 Add Rank Level End
    End With
    
    With Me.flxGrade
        For intRow = 0 To .Rows - 1
            For intCol = 0 To .Cols - 1
                .Row = intRow
                .Col = intCol
                .CellAlignment = flexAlignCenterCenter
                .RowHeight(intRow) = 350
            Next intCol
        Next intRow
        
        .ColWidth(0) = 400
        .ColWidth(0) = 800
        .TextMatrix(0, 0) = "GRADE"
        .TextMatrix(0, 1) = "RANK"
    End With
    
End Sub
 '============Leo 2012.05.22 Add Rank Level Start
Private Sub Init_RankLevel()
    Dim intIndex As Integer
    Dim lblToLeft As Integer
    Dim lblWidth As Integer
    Dim txtToLbl As Integer
    Dim txtWidth As Integer
    lblToLeft = 240
    lblWidth = 600
    txtToLbl = 50
    txtWidth = 1100
    
    For intIndex = 0 To UBound(RankLevel)
        If intIndex <> 0 Then
            Load lblRank(intIndex)
            Load txtRank(intIndex)
        End If
            Me.lblRank(intIndex).Caption = RankLevel(intIndex) & " RANK"
            Me.lblRank(intIndex).Width = lblWidth
            Me.lblRank(intIndex).Top = 9360
            Me.lblRank(intIndex).Left = lblToLeft + (lblToLeft + lblWidth + txtToLbl + txtWidth) * intIndex
            Me.lblRank(intIndex).Visible = True
            
            Me.txtRank(intIndex).Width = txtWidth
            Me.txtRank(intIndex).Top = 9300
            Me.txtRank(intIndex).Left = lblToLeft + lblWidth + txtToLbl + (lblToLeft + lblWidth + txtToLbl + txtWidth) * intIndex
            Me.txtRank(intIndex).Visible = True
    Next intIndex
End Sub
Private Function Add_Grid(ByVal pSection As String) As Integer

    Dim intRow          As Integer
    Dim intCol          As Integer
    
    With Me.flxRank_Data
        intRow = .Rows
        .AddItem pSection
        .RowHeight(intRow) = 350
        For intCol = 0 To .Cols - 1
            .Row = intRow
            .Col = intCol
            .CellAlignment = flexAlignCenterCenter
        Next intCol
        
        Add_Grid = .Rows - 1
    End With

End Function

Private Sub mnuGrade_List_Input_Click()

    Load frmGrade_List
    frmGrade_List.Show
    
End Sub

Private Sub txtAccumulation_Change()

    If Me.cmbDetail_Division.Text = "X" Then
        Me.txtAccumulation.Text = "X"
    End If
    
End Sub

Private Sub txtDefect_Code_LostFocus()

    Me.txtDefect_Code.Text = UCase(Me.txtDefect_Code.Text)
    
    If Me.cmbSection.Text = "" Then
        Call MsgBox("Please select defect section first", vbOKOnly, "Data error")
        Me.txtDefect_Code.Text = ""
    Else
        Select Case Mid(Me.txtDefect_Code.Text, 2, 1)
        Case "D":
            If Me.cmbSection.Text <> "POINT D/F" Then
                Call MsgBox("Defect section and defect code type does not match", vbOKOnly, "Data error")
                Me.txtDefect_Code.Text = ""
            End If
        Case "L":
            If Me.cmbSection.Text <> "LINE D/F" Then
                Call MsgBox("Defect section and defect code type does not match", vbOKOnly, "Data error")
                Me.txtDefect_Code.Text = ""
            End If
        Case "M":
            If Me.cmbSection.Text <> "MURA D/F" Then
                Call MsgBox("Defect section and defect code type does not match", vbOKOnly, "Data error")
                Me.txtDefect_Code.Text = ""
            End If
        Case "G":
            If Me.cmbSection.Text <> "GAP D/F" Then
                Call MsgBox("Defect section and defect code type does not match", vbOKOnly, "Data error")
                Me.txtDefect_Code.Text = ""
            End If
        Case "F":
            If Me.cmbSection.Text <> "CF D/F" Then
                Call MsgBox("Defect section and defect code type does not match", vbOKOnly, "Data error")
                Me.txtDefect_Code.Text = ""
            End If
        Case "A":
            If Me.cmbSection.Text <> "APPEARANCE D/F" Then
                Call MsgBox("Defect section and defect code type does not match", vbOKOnly, "Data error")
                Me.txtDefect_Code.Text = ""
            End If
        Case "C":
            If Me.cmbSection.Text <> "CELL D/F" Then
                Call MsgBox("Defect section and defect code type does not match", vbOKOnly, "Data error")
                Me.txtDefect_Code.Text = ""
            End If
        Case "O":
            If Me.cmbSection.Text <> "OTHER D/F" Then
                Call MsgBox("Defect section and defect code type does not match", vbOKOnly, "Data error")
                Me.txtDefect_Code.Text = ""
            End If
        Case "P":
            If Me.cmbSection.Text <> "POLARIZE D/F" Then
                Call MsgBox("Defect section and defect code type does not match", vbOKOnly, "Data error")
                Me.txtDefect_Code.Text = ""
            End If
        Case Else
            Call MsgBox("Defect section and defect code type does not match", vbOKOnly, "Data error")
            Me.txtDefect_Code.Text = ""
        End Select
    End If
    
End Sub

Private Sub txtGrade_Rank_Change()

    Dim intIndex                        As Integer
    
    Dim bolMatch_Rank                   As Boolean
    Dim bolDuplicate                    As Boolean
    
    Me.txtGrade_Rank.Text = UCase(Me.txtGrade_Rank.Text)
    SendKeys "{End}"
    
    bolMatch_Rank = False
    For intIndex = 10 To 14
        If Me.flxRank_Data.TextMatrix(0, intIndex) = Right(Me.txtGrade_Rank.Text, 1) Then
            bolMatch_Rank = True
        End If
    Next intIndex
    
    If bolMatch_Rank = False Then
        If InStr(Me.txtGrade_Rank.Text, "-") > 0 Then
            Me.txtGrade_Rank.Text = "-"
        Else
            If Me.txtGrade_Rank.Text <> "" Then
                Me.txtGrade_Rank.Text = Left(Me.txtGrade_Rank, Len(Me.txtGrade_Rank.Text) - 1)
            Else
                Me.txtGrade_Rank.Text = ""
            End If
        End If
    Else
        If InStr(Me.txtGrade_Rank.Text, "-") > 0 Then
            Me.txtGrade_Rank.Text = "-"
        Else
            bolDuplicate = False
            If Me.txtGrade_Rank.Text <> "" Then
                For intIndex = 1 To Len(Me.txtGrade_Rank.Text) - 1
                    If Mid(Me.txtGrade_Rank.Text, intIndex, 1) = Right(Me.txtGrade_Rank.Text, 1) Then
                        bolDuplicate = True
                    End If
                Next intIndex
                If bolDuplicate = True Then
                    Me.txtGrade_Rank.Text = Left(Me.txtGrade_Rank, Len(Me.txtGrade_Rank.Text) - 1)
                End If
            End If
        End If
    End If

End Sub

Private Sub txtPriority_Change()

    If Me.txtPriority.Text <> "" Then
        If IsNumeric(Me.txtPriority.Text) = False Then
            Call MsgBox("Data format error. Please type in numeric data.", vbOKOnly, "Data error")
            Me.txtPriority.Text = ""
        End If
    End If
    
End Sub

 '============Leo 2012.05.22 Add Rank Level Start

Private Sub txtRank_Change(Index As Integer)

    If Me.cmbDefect_Type.Text = "P" Then
        If txtRank(Index).Text <> "" Then
            If Left(txtRank(Index).Text, 1) <> "-" Then
                If IsNumeric(Me.txtRank(Index).Text) = False Then
                    Call MsgBox("Data type mismatch. Check a defect type data", vbOKOnly, "Data error")
                    Me.txtRank(Index).Text = ""
                ElseIf InStr(Me.txtRank(Index).Text, vbCrLf) > 0 Then
                    Call MsgBox("Data typ mismatch. Check a defect type data", vbOKOnly, "Data error")
                    Me.txtRank(Index).Text = ""
                End If
            Else
                If Len(Me.txtRank(Index).Text) > 1 Then
                    Me.txtRank(Index).Text = "-"
                    SendKeys "{End}"
                End If
            End If
        End If
    End If

    
'    If Me.cmbDefect_Type.Text = "P" Then
'        If txtRank.Text <> "" Then
'            If Left(Me.txtRank_Y.Text, 1) <> "-" Then
'                If IsNumeric(Me.txtRank_Y.Text) = False Then
'                    Call MsgBox("Data type mismatch. Check a defect type data", vbOKOnly, "Data error")
'                    Me.txtRank_Y.Text = ""
'                ElseIf InStr(Me.txtRank_Y.Text, vbCrLf) > 0 Then
'                    Call MsgBox("Data typ mismatch. Check a defect type data", vbOKOnly, "Data error")
'                    Me.txtRank_Y.Text = ""
'                End If
'            Else
'                If Len(Me.txtRank_Y.Text) > 1 Then
'                    Me.txtRank_Y.Text = "-"
'                    SendKeys "{End}"
'                End If
'            End If
'        End If
'    End If

End Sub

Private Function Check_Data() As Boolean

    Dim intRow                          As Integer
    Dim intIndex                        As Integer
    
    Dim bolData_Check                   As Boolean
    Dim bolFind                         As Boolean
    
    With Me
        bolData_Check = True
        If .cmbSection.Text = "" Then
            bolData_Check = False
        End If
        If .txtDefect_Code.Text = "" Then
            bolData_Check = False
        End If
        If .txtDefect_Name_English.Text = "" Then
            bolData_Check = False
        End If
        If .txtDefect_Name_China.Text = "" Then
            bolData_Check = False
        End If
        If .cmbDefect_Type.Text = "" Then
            bolData_Check = False
        End If
        If .cmbJudge.Text = "" Then
            bolData_Check = False
        End If
        If .cmbXY_Axis.Text = "" Then
            bolData_Check = False
        End If
        If .cmbDetail_Division.Text = "" Then
            bolData_Check = False
        End If
        If .txtAccumulation.Text = "" Then
            .txtAccumulation.Text = "X"
        End If
        If .cmbAddress_Count.Text = "" Then
            bolData_Check = False
        End If
        If .txtPriority.Text = "" Then
            bolData_Check = False
        Else
            If .flxRank_Data.Rows > 1 Then
                intIndex = 0
                bolFind = False
                While bolFind = False
                    intIndex = intIndex + 1
                    If (.flxRank_Data.TextMatrix(intIndex, 15) = .txtPriority.Text) And (.flxRank_Data.TextMatrix(intIndex, 0) = .cmbSection.Text) Then
                        bolFind = True
                        intRow = intIndex
                    Else
                        If intIndex = .flxRank_Data.Rows - 1 Then
                            bolFind = True
                        End If
                    End If
                Wend
                If intRow > 0 Then
                    If .flxRank_Data.TextMatrix(intRow, 1) <> .txtDefect_Code.Text Then
                        bolData_Check = False
                    End If
                End If
            End If
        End If
        
        Check_Data = bolData_Check
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

