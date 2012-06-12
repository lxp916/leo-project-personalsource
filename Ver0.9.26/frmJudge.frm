VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmJudge 
   Caption         =   "JUDGE WINDOW"
   ClientHeight    =   11028
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   19080
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   11028
   ScaleWidth      =   19080
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Frame Frame3 
      Height          =   2895
      Left            =   10260
      TabIndex        =   10
      Top             =   8100
      Width           =   9885
      Begin MSFlexGridLib.MSFlexGrid flxDefect_List 
         Height          =   2625
         Left            =   60
         TabIndex        =   47
         Top             =   180
         Width           =   9765
         _ExtentX        =   17230
         _ExtentY        =   4636
         _Version        =   393216
         Rows            =   1
         Cols            =   15
         FixedCols       =   0
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2565
      Left            =   10260
      TabIndex        =   7
      Top             =   5520
      Width           =   5295
      Begin VB.CommandButton cmdSet_Address 
         Caption         =   "Set Address"
         Height          =   525
         Left            =   2190
         TabIndex        =   51
         Top             =   1890
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.TextBox txtY_Gate 
         Height          =   315
         Index           =   2
         Left            =   3540
         MaxLength       =   5
         TabIndex        =   34
         Top             =   1440
         Width           =   1365
      End
      Begin VB.TextBox txtX_Data 
         Height          =   300
         Index           =   2
         Left            =   1200
         MaxLength       =   5
         TabIndex        =   33
         Top             =   1470
         Width           =   1365
      End
      Begin VB.TextBox txtY_Gate 
         Height          =   315
         Index           =   1
         Left            =   3540
         MaxLength       =   5
         TabIndex        =   32
         Top             =   870
         Width           =   1365
      End
      Begin VB.TextBox txtX_Data 
         Height          =   300
         Index           =   1
         Left            =   1200
         MaxLength       =   5
         TabIndex        =   31
         Top             =   900
         Width           =   1365
      End
      Begin VB.TextBox txtY_Gate 
         Height          =   315
         Index           =   0
         Left            =   3540
         MaxLength       =   5
         TabIndex        =   30
         Top             =   300
         Width           =   1365
      End
      Begin VB.TextBox txtX_Data 
         Height          =   300
         Index           =   0
         Left            =   1200
         MaxLength       =   5
         TabIndex        =   29
         Top             =   300
         Width           =   1365
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "X (Data)"
         Height          =   180
         Index           =   2
         Left            =   420
         TabIndex        =   28
         Top             =   1500
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Y (Gate)"
         Height          =   180
         Index           =   2
         Left            =   2730
         TabIndex        =   27
         Top             =   1500
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "X (Data)"
         Height          =   180
         Index           =   1
         Left            =   420
         TabIndex        =   26
         Top             =   930
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Y (Gate)"
         Height          =   180
         Index           =   1
         Left            =   2730
         TabIndex        =   25
         Top             =   930
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Y (Gate)"
         Height          =   180
         Index           =   0
         Left            =   2730
         TabIndex        =   9
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "X (Data)"
         Height          =   180
         Index           =   0
         Left            =   420
         TabIndex        =   8
         Top             =   360
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      Height          =   11115
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20235
      Begin MSFlexGridLib.MSFlexGrid flxDefect_I 
         Height          =   3315
         Left            =   6750
         TabIndex        =   44
         Top             =   7650
         Width           =   3405
         _ExtentX        =   5990
         _ExtentY        =   5842
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         SelectionMode   =   1
      End
      Begin MSFlexGridLib.MSFlexGrid flxDefect_H 
         Height          =   3315
         Left            =   3390
         TabIndex        =   43
         Top             =   7650
         Width           =   3375
         _ExtentX        =   5948
         _ExtentY        =   5842
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         SelectionMode   =   1
      End
      Begin MSFlexGridLib.MSFlexGrid flxDefect_G 
         Height          =   3315
         Left            =   60
         TabIndex        =   42
         Top             =   7650
         Width           =   3375
         _ExtentX        =   5948
         _ExtentY        =   5842
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         SelectionMode   =   1
      End
      Begin VB.Frame Frame7 
         Height          =   825
         Left            =   15600
         TabIndex        =   17
         Top             =   7260
         Width           =   4545
         Begin VB.CommandButton cmdDefect_Delete 
            Caption         =   "DELETE DEFECT"
            Height          =   525
            Left            =   1560
            TabIndex        =   18
            Top             =   180
            Width           =   1665
         End
      End
      Begin VB.Frame Frame6 
         Height          =   5385
         Left            =   10260
         TabIndex        =   16
         Top             =   120
         Width           =   9885
         Begin VB.Frame Frame5 
            Height          =   1845
            Left            =   5340
            TabIndex        =   48
            Top             =   3420
            Width           =   4425
            Begin MSFlexGridLib.MSFlexGrid flxPG_Data 
               Height          =   1635
               Left            =   60
               TabIndex        =   49
               Top             =   150
               Width           =   4305
               _ExtentX        =   7599
               _ExtentY        =   2879
               _Version        =   393216
               Rows            =   1
               Cols            =   4
               FixedCols       =   0
            End
         End
         Begin VB.Frame Frame9 
            Height          =   1845
            Left            =   120
            TabIndex        =   22
            Top             =   3420
            Width           =   5175
            Begin VB.CommandButton cmdGrade 
               Caption         =   "GRADE"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   15.6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1485
               Left            =   2730
               TabIndex        =   24
               Top             =   210
               Width           =   2175
            End
            Begin VB.CommandButton cmdMain_Window 
               Caption         =   "MAIN"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   15.6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1485
               Left            =   210
               TabIndex        =   23
               Top             =   210
               Width           =   2175
            End
         End
         Begin VB.Frame Frame8 
            Height          =   3285
            Left            =   120
            TabIndex        =   19
            Top             =   120
            Width           =   9645
            Begin VB.Timer tmrPattern_Delay 
               Enabled         =   0   'False
               Interval        =   1000
               Left            =   8160
               Top             =   150
            End
            Begin VB.PictureBox picCurrent_Pattern 
               BackColor       =   &H00000000&
               Height          =   2655
               Left            =   120
               ScaleHeight     =   2604
               ScaleWidth      =   9180
               TabIndex        =   20
               Top             =   480
               Width           =   9225
               Begin VB.TextBox Text15 
                  Height          =   270
                  Left            =   2640
                  TabIndex        =   66
                  Top             =   2280
                  Width           =   615
               End
               Begin VB.TextBox Text14 
                  Height          =   270
                  Left            =   2040
                  TabIndex        =   65
                  Top             =   2280
                  Width           =   615
               End
               Begin VB.TextBox Text13 
                  Height          =   270
                  Left            =   1320
                  TabIndex        =   64
                  Top             =   2280
                  Width           =   615
               End
               Begin VB.TextBox Text12 
                  Height          =   270
                  Left            =   720
                  TabIndex        =   63
                  Top             =   2280
                  Width           =   615
               End
               Begin VB.TextBox Text11 
                  Height          =   270
                  Left            =   0
                  TabIndex        =   62
                  Top             =   2280
                  Width           =   615
               End
               Begin VB.TextBox Text10 
                  Height          =   270
                  Left            =   2640
                  TabIndex        =   61
                  Top             =   2040
                  Width           =   615
               End
               Begin VB.TextBox Text9 
                  Height          =   270
                  Left            =   2040
                  TabIndex        =   60
                  Top             =   2040
                  Width           =   615
               End
               Begin VB.TextBox Text8 
                  Height          =   270
                  Left            =   1320
                  TabIndex        =   59
                  Top             =   2040
                  Width           =   615
               End
               Begin VB.TextBox Text6 
                  Height          =   270
                  Left            =   0
                  TabIndex        =   58
                  Top             =   2040
                  Width           =   615
               End
               Begin VB.TextBox Text7 
                  Height          =   270
                  Left            =   720
                  TabIndex        =   57
                  Top             =   2040
                  Width           =   615
               End
               Begin VB.TextBox Text5 
                  Height          =   270
                  Left            =   2640
                  TabIndex        =   56
                  Top             =   1800
                  Width           =   615
               End
               Begin VB.TextBox Text4 
                  Height          =   270
                  Left            =   2040
                  TabIndex        =   55
                  Top             =   1800
                  Width           =   615
               End
               Begin VB.TextBox Text3 
                  Height          =   270
                  Left            =   1320
                  TabIndex        =   54
                  Top             =   1800
                  Width           =   615
               End
               Begin VB.TextBox Text2 
                  Height          =   270
                  Left            =   720
                  TabIndex        =   53
                  Top             =   1800
                  Width           =   615
               End
               Begin VB.TextBox Text1 
                  Height          =   270
                  Left            =   0
                  TabIndex        =   52
                  Top             =   1800
                  Width           =   615
               End
               Begin VB.Image imgPG_Image 
                  Height          =   1095
                  Left            =   0
                  Stretch         =   -1  'True
                  Top             =   0
                  Width           =   4695
               End
            End
            Begin VB.Label lblCurrent_PTN_Index 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "0"
               Height          =   180
               Left            =   2010
               TabIndex        =   50
               Top             =   270
               Visible         =   0   'False
               Width           =   105
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Current Pattern"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   210
               TabIndex        =   21
               Top             =   240
               Width           =   1635
            End
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1725
         Left            =   15600
         TabIndex        =   11
         Top             =   5520
         Width           =   4545
         Begin VB.CommandButton cmdDelete 
            Caption         =   "DELETE"
            Height          =   555
            Left            =   2370
            TabIndex        =   15
            Top             =   1020
            Width           =   1425
         End
         Begin VB.CommandButton cmdAppend 
            Caption         =   "APPEND"
            Height          =   555
            Left            =   720
            TabIndex        =   14
            Top             =   1020
            Width           =   1425
         End
         Begin VB.ComboBox cmbUseful_Defect 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            TabIndex        =   13
            Top             =   540
            Width           =   4335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "USEUL DEFECT"
            Height          =   180
            Left            =   240
            TabIndex        =   12
            Top             =   270
            Width           =   1365
         End
      End
      Begin MSFlexGridLib.MSFlexGrid flxDefect_F 
         Height          =   3315
         Left            =   6780
         TabIndex        =   6
         Top             =   4050
         Width           =   3405
         _ExtentX        =   6011
         _ExtentY        =   5842
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         SelectionMode   =   1
      End
      Begin MSFlexGridLib.MSFlexGrid flxDefect_E 
         Height          =   3315
         Left            =   3420
         TabIndex        =   5
         Top             =   4050
         Width           =   3405
         _ExtentX        =   6011
         _ExtentY        =   5842
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         SelectionMode   =   1
      End
      Begin MSFlexGridLib.MSFlexGrid flxDefect_D 
         Height          =   3315
         Left            =   60
         TabIndex        =   4
         Top             =   4050
         Width           =   3405
         _ExtentX        =   6011
         _ExtentY        =   5842
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         SelectionMode   =   1
      End
      Begin MSFlexGridLib.MSFlexGrid flxDefect_C 
         Height          =   3315
         Left            =   6840
         TabIndex        =   3
         Top             =   480
         Width           =   3405
         _ExtentX        =   6011
         _ExtentY        =   5842
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         SelectionMode   =   1
      End
      Begin MSFlexGridLib.MSFlexGrid flxDefect_B 
         Height          =   3315
         Left            =   3420
         TabIndex        =   2
         Top             =   480
         Width           =   3405
         _ExtentX        =   6011
         _ExtentY        =   5842
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         SelectionMode   =   1
      End
      Begin MSFlexGridLib.MSFlexGrid flxDefect_A 
         Height          =   3315
         Left            =   60
         TabIndex        =   1
         Top             =   480
         Width           =   3405
         _ExtentX        =   6011
         _ExtentY        =   5842
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         SelectionMode   =   1
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Other D/F"
         Height          =   180
         Index           =   8
         Left            =   6780
         TabIndex        =   46
         Top             =   7440
         Width           =   825
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Cell D/F"
         Height          =   180
         Index           =   7
         Left            =   3450
         TabIndex        =   45
         Top             =   7440
         Width           =   705
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "APPEARANCE D/F"
         Height          =   180
         Index           =   6
         Left            =   120
         TabIndex        =   41
         Top             =   7440
         Width           =   1605
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Polarized D/F"
         Height          =   180
         Index           =   5
         Left            =   6810
         TabIndex        =   40
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "CF D/F"
         Height          =   180
         Index           =   4
         Left            =   3480
         TabIndex        =   39
         Top             =   3840
         Width           =   615
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Mura D/F"
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   38
         Top             =   3840
         Width           =   810
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "GAP D/F"
         Height          =   180
         Index           =   2
         Left            =   6810
         TabIndex        =   37
         Top             =   270
         Width           =   750
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Line D/F"
         Height          =   180
         Index           =   1
         Left            =   3480
         TabIndex        =   36
         Top             =   270
         Width           =   735
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Point D/F"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   35
         Top             =   270
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmJudge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_DEFECT_CODE                   As String
Dim m_DEFECT_NAME                   As String
Dim m_DEFECT_KIND                   As String

Dim m_CURRENT_DEFECT_INDEX          As Integer
Dim m_PATTERN_COUNT                 As Integer

Dim PATTERN_LIST()                  As PATTERN_LIST_DATA

Private Sub cmbUseful_Defect_Click()

    Dim typRANK_DATA        As RANK_DATA_STRUCTURE
    Dim typPATTERN_DATA     As PATTERN_LIST_DATA
    
    Dim intRow              As Integer
    Dim intIndex            As Integer
    Dim intPos              As Integer
    Dim intPTN_Index        As Integer
    
    Dim strDefect_Code      As String
    Dim strDEFECT_NAME      As String
    Dim strDEFECT_KIND      As String
    Dim strAddress_Count    As String
    
    If Me.cmbUseful_Defect.Text <> "" Then
        intPos = InStr(Me.cmbUseful_Defect.Text, " ")
        strDefect_Code = Left(Me.cmbUseful_Defect.Text, intPos - 1)
        strDEFECT_NAME = Mid(Me.cmbUseful_Defect.Text, intPos + 1)
        
        If Get_Defect_Type(strDefect_Code) <> "A" Then
            strDEFECT_KIND = Mid(strDefect_Code, 2, 1)
            
            For intIndex = 0 To 2
                Me.txtX_Data(intIndex).Enabled = False
                Me.txtY_Gate(intIndex).Enabled = False
            Next intIndex
            Me.txtX_Data(0).Enabled = True
            Me.txtY_Gate(0).Enabled = True
            
            intRow = Add_Grid(strDefect_Code, Me.flxDefect_List)
            With Me.flxDefect_List
                .TextMatrix(intRow, 1) = strDEFECT_NAME
                .TextMatrix(intRow, 8) = frmMain.flxMES_Data.TextMatrix(17, 1)
            
                intPTN_Index = CInt(Me.lblCurrent_PTN_Index.Caption)
                With typPATTERN_DATA
                    Call EQP.Get_PATTERN_LIST_by_Index(intPTN_Index, .PATTERN_CODE, .PATTERN_NAME, .DELAY_TIME, .LEVEL, .DH, .DL, .VGH, .VGL, .RESCUE_HIGH, .RESCUE_LOW, .VCOM)
                End With
                .TextMatrix(intRow, 12) = typPATTERN_DATA.PATTERN_NAME
            End With
        
            m_DEFECT_CODE = ""
            m_DEFECT_NAME = ""
            m_DEFECT_KIND = ""
            
            strAddress_Count = Get_Defect_Address_Count(strDefect_Code)
            If strAddress_Count = "X" Then
                Call Load_Manual_Judge
            ElseIf strAddress_Count = "0" Then
            Else
                Call Set_Interlock
            End If
        End If
    End If
    
End Sub

Private Sub cmdAppend_Click()

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    
    Dim ErrMsg                      As String
    
On Error GoTo ErrorHandler
    
    If (m_DEFECT_CODE <> "") And (m_DEFECT_NAME <> "") And (m_DEFECT_KIND <> "") Then
        strDB_Path = App.PATH & "\DB\"
        strDB_FileName = "Parameter.mdb"
        
        If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
            Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
            
            strQuery = "INSERT INTO USEFUL_DEFECT VALUES ("
            strQuery = strQuery & "'" & m_DEFECT_CODE & "', "
            strQuery = strQuery & "'" & m_DEFECT_NAME & "', "
            strQuery = strQuery & "'" & m_DEFECT_KIND & "')"
            
            dbMyDB.Execute strQuery
            
            Me.cmbUseful_Defect.Clear
            
            strQuery = "SELECT * FROM USEFUL_DEFECT"
            
            Set lstRecord = dbMyDB.OpenRecordset(strQuery)
            
            If lstRecord.EOF = False Then
                lstRecord.MoveFirst
                While lstRecord.EOF = False
                    Me.cmbUseful_Defect.AddItem lstRecord.Fields("DEFECT_CODE") & " " & lstRecord.Fields("DEFECT_NAME")
                    lstRecord.MoveNext
                Wend
            End If
            lstRecord.Close
            Me.cmbUseful_Defect.Text = Me.cmbUseful_Defect.List(0)
            
            dbMyDB.Close
        End If
        
        m_DEFECT_CODE = ""
        m_DEFECT_NAME = ""
        m_DEFECT_KIND = ""
    End If
    
    Exit Sub
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("frmJudge_cmdAppend_Click", ErrMsg)
    
End Sub
'
'Private Sub cmdAssign_Click()
'
'    Dim typRANK_DATA                    As RANK_DATA_STRUCTURE
'    Dim typGRADE_DATA()                 As GRADE_DATA_STRUCTURE
'
'    Dim strDATA_ADDRESS(1 To 3)         As String
'    Dim strGATE_ADDRESS(1 To 3)         As String
'
'    Dim strDEFECT_CODE                  As String
'
'    Dim intDefect_Count                 As Integer
'    Dim intIndex                        As Integer
'    Dim intCol                          As Integer
'    Dim intRow                          As Integer
'    Dim intGrade_Count                  As Integer
'
'    Call Reset_Interlock
'
'    intRow = frmJudge.flxDefect_List.Rows - 1
'    strDEFECT_CODE = frmJudge.flxDefect_List.TextMatrix(intRow, 0)
'
'    If Mid(strDEFECT_CODE, 2, 1) = "M" Then
'        For intIndex = 1 To 3
'            If Me.txtX_Data(intIndex - 1) <> "" Then
'                strDATA_ADDRESS(intIndex) = Me.txtX_Data(intIndex - 1).Text
'            Else
'                strDATA_ADDRESS(intIndex) = Space(5)
'            End If
'            Me.txtX_Data(intIndex - 1).Text = ""
'            If Me.txtY_Gate(intIndex - 1) <> "" Then
'                strGATE_ADDRESS(intIndex) = Me.txtY_Gate(intIndex - 1).Text
'            Else
'                strGATE_ADDRESS(intIndex) = Space(5)
'            End If
'            Me.txtY_Gate(intIndex - 1).Text = ""
'        Next intIndex
'    End If
'
'    intIndex = 0
'    For intCol = 2 To 7 Step 2
'        intIndex = intIndex + 1
'        With frmJudge.flxDefect_List
'            .TextMatrix(intRow, intCol) = strDATA_ADDRESS(intIndex)
'            .TextMatrix(intRow, intCol + 1) = strGATE_ADDRESS(intIndex)
'        End With
'    Next intCol
'
'    'Check Defect Type
'    Call Get_Rank_Data(pubCST_INFO.PROCESS_NUM, typRANK_DATA, typGRADE_DATA, strDEFECT_CODE, intGrade_Count)
'
'    If typRANK_DATA.DEFECT_TYPE = "R" Then
'        'Manual judge
'        If typRANK_DATA.POP_UP = "E" Then
'            'Sub window pop up
'            Load frmManual_Judge
'
'            With frmManual_Judge
'                If (Trim(typRANK_DATA.RANK_Y) <> "0") And (Trim(typRANK_DATA.RANK_Y) <> "-") Then
'                    .lblGrade(0).Caption = "Y"
'                    .optSpec_Value(0).Caption = typRANK_DATA.RANK_Y
'                    .lblGrade(0).Visible = True
'                    .optSpec_Value(0).Visible = True
'                End If
'
'                If (Trim(typRANK_DATA.RANK_L) <> "0") And (Trim(typRANK_DATA.RANK_L) <> "-") Then
'                    .lblGrade(1).Caption = "L"
'                    .optSpec_Value(1).Caption = typRANK_DATA.RANK_L
'                    .lblGrade(1).Visible = True
'                    .optSpec_Value(1).Visible = True
'                End If
'
'                If (Trim(typRANK_DATA.RANK_K) <> "0") And (Trim(typRANK_DATA.RANK_K) <> "-") Then
'                    .lblGrade(2).Caption = "K"
'                    .optSpec_Value(2).Caption = typRANK_DATA.RANK_K
'                    .lblGrade(2).Visible = True
'                    .optSpec_Value(2).Visible = True
'                End If
'
'                If (Trim(typRANK_DATA.RANK_C) <> "0") And (Trim(typRANK_DATA.RANK_C) <> "-") Then
'                    .lblGrade(3).Caption = "C"
'                    .optSpec_Value(3).Caption = typRANK_DATA.RANK_C
'                    .lblGrade(3).Visible = True
'                    .optSpec_Value(3).Visible = True
'                End If
'
'                If (Trim(typRANK_DATA.RANK_S) <> "0") And (Trim(typRANK_DATA.RANK_S) <> "-") Then
'                    .lblGrade(4).Caption = "S"
'                    .optSpec_Value(4).Caption = typRANK_DATA.RANK_S
'                    .lblGrade(4).Visible = True
'                    .optSpec_Value(4).Visible = True
'                End If
'                .lblDefect_Code.Caption = strDEFECT_CODE
'                .lblDefect_Name.Caption = frmJudge.flxDefect_List.TextMatrix(intRow, 1)
'                .lstData_Address.Clear
'                .lstGate_Address.Clear
'
'                If Mid(.lblDefect_Code.Caption, 2, 1) = "M" Then
'                    For intIndex = 1 To 3
'                        .lstData_Address.AddItem strDATA_ADDRESS(intIndex)
'                        .lstGate_Address.AddItem strGATE_ADDRESS(intIndex)
'                    Next intIndex
'                Else
'                    .lstData_Address.AddItem strDATA_ADDRESS(1)
'                    .lstGate_Address.AddItem strGATE_ADDRESS(1)
'                End If
'            End With
'
'            frmManual_Judge.Show
'        End If
'    End If
'    frmJudge.flxDefect_List.TextMatrix(frmJudge.flxDefect_List.Rows - 1, 11) = typRANK_DATA.DETAIL_DIVISION
'
'End Sub

Private Sub cmdDefect_Delete_Click()

    Dim intIndex                    As Integer
    
    intIndex = Me.flxDefect_List.Row
    
    If Me.flxDefect_List.Rows > 2 Then
        Me.flxDefect_List.RemoveItem (intIndex)
        Call RANK_OBJ.Set_Select_DEFECTCODE(Me.flxDefect_List.TextMatrix(Me.flxDefect_List.Rows - 1, 0))
    ElseIf Me.flxDefect_List.Rows = 2 Then
        Me.flxDefect_List.Rows = 1
        Call RANK_OBJ.Set_Select_DEFECTCODE("")
    End If
    
    m_DEFECT_CODE = ""
    m_DEFECT_NAME = ""
    m_DEFECT_KIND = ""
    
    Call Reset_Interlock
    
End Sub

Private Sub cmdDelete_Click()

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strDefect_Code              As String
    Dim strDEFECT_NAME              As String
    Dim strQuery                    As String
    
    Dim intPosition                 As Integer
    
    Dim ErrMsg                      As String
    
On Error GoTo ErrorHandler

    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Parameter.mdb"
    
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
        intPosition = InStr(Me.cmbUseful_Defect.Text, " ")
        If intPosition > 0 Then
            strDefect_Code = Left(Me.cmbUseful_Defect.Text, intPosition - 1)
            strDEFECT_NAME = Mid(Me.cmbUseful_Defect.Text, intPosition + 1)
            
            strQuery = "DELETE FROM USEFUL_DEFECT WHERE "
            strQuery = strQuery & "DEFECT_CODE = '" & strDefect_Code & "' AND "
            strQuery = strQuery & "DEFECT_NAME = '" & strDEFECT_NAME & "'"
            
            dbMyDB.Execute strQuery
        End If
        
        Me.cmbUseful_Defect.Clear
        
        strQuery = "SELECT * FROM USEFUL_DEFECT"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        
        If lstRecord.EOF = False Then
            lstRecord.MoveFirst
            While lstRecord.EOF = False
                Me.cmbUseful_Defect.AddItem lstRecord.Fields("DEFECT_CODE") & " " & lstRecord.Fields("DEFECT_NAME")
                lstRecord.MoveNext
            Wend
        End If
        lstRecord.Close
        If Me.cmbUseful_Defect.ListCount > 0 Then
            Me.cmbUseful_Defect.Text = Me.cmbUseful_Defect.List(0)
        End If
        
        dbMyDB.Close
    
        DBEngine.CompactDatabase strDB_Path & strDB_FileName, strDB_Path & "Parameter_Temp.mdb", dbLangChineseSimplified
        Kill strDB_Path & strDB_FileName
        Name strDB_Path & "Parameter_Temp.mdb" As strDB_Path & strDB_FileName
    End If
    
    Exit Sub
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    Call SaveLog("frmJudge_cmdDelete_Click", ErrMsg)
    
End Sub

Private Sub cmdGrade_Click()
    
    Dim dbMyDB                              As Database

    Dim typGRADE_DATA()                     As GRADE_DATA_STRUCTURE
    Dim typDEFECT_DATA()                    As DEFECT_DATA_STRUCTURE

    Dim typRANK_DATA                        As RANK_DATA_STRUCTURE
    Dim typGRADE_DEFECT                     As DEFECT_DATA_STRUCTURE
    Dim typPFCD_ADDRESS_DATA                As PFCD_ADDRESS_STRUCTURE
    Dim typGRADE_DEFECT_DATA                As DEFECT_DATA_STRUCTURE
    
    Dim arrPOINT_DEFECT_COUNT(1 To 3)       As Integer

    Dim arrPOINT_DISTANCE(1 To 3)           As Double
    
    Dim strDB_Path                          As String
    Dim strDB_FileName                      As String
    Dim strQuery                            As String
    Dim strNew_Grade                        As String
    Dim strPoint_Defect_Rank                As String
    Dim strGrade                            As String
    Dim strRank                             As String
    Dim strDEFECT_Rank                      As String
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
    Dim intSource_Rank_Priority             As Integer
    Dim intTarget_Rank_Priority             As Integer
    Dim msec                                As Long
    
    Call ENV.Get_Device_Data_by_Name("API", intPortNo, strState)
    
    intPortNo = EQP.Get_PG_PortID
    If intPortNo > 0 Then
        Call QUEUE.Put_Send_Command(intPortNo, "QPPF")
    End If

    If Left(frmMain.flxEQ_Information.TextMatrix(5, 1), 5) = "CATST" Then
        If intPortNo > 0 Then
            Call QUEUE.Put_Send_Command(intPortNo, "QBLV")
        End If
    End If

    intDefect_Count = Me.flxDefect_List.Rows - 1
    Call RANK_OBJ.Set_DEFECT_DATA_COUNT(intDefect_Count + 6)
    ReDim typDEFECT_DATA(intDefect_Count + 6)
    ReDim pubDEFECT_DATA(intDefect_Count)       '2012.03.26 Added by K.H.KIM
    pubDefect_Count = intDefect_Count           '2012.03.26 Added by K.H.KIM
    
    With Me.flxDefect_List
        If intDefect_Count > 0 Then
            For intRow = 1 To intDefect_Count
                pubDEFECT_DATA(intRow).DEFECT_CODE = .TextMatrix(intRow, 0)     '2012.03.26 Added by K.H.KIM
                pubDEFECT_DATA(intRow).DEFECT_NAME = .TextMatrix(intRow, 1)     '2012.03.26 Added by K.H.KIM
                pubDEFECT_DATA(intRow).PANELID = .TextMatrix(intRow, 8)         '2012.03.26 Added by K.H.KIM
                pubDEFECT_DATA(intRow).DEFECT_NO = intRow                       '2012.03.26 Added by K.H.KIM
                
                .TextMatrix(intRow, 8) = pubPANEL_INFO.PANELID
                typDEFECT_DATA(intRow).DEFECT_NO = intRow
                typDEFECT_DATA(intRow).DEFECT_CODE = .TextMatrix(intRow, 0)
                typDEFECT_DATA(intRow).DEFECT_NAME = .TextMatrix(intRow, 1)
                typDEFECT_DATA(intRow).PANELID = .TextMatrix(intRow, 8)
                intIndex = 0
                For intCol = 2 To 7 Step 2
                    intIndex = intIndex + 1
                    typDEFECT_DATA(intRow).DATA_ADDRESS(intIndex) = .TextMatrix(intRow, intCol)
                    typDEFECT_DATA(intRow).GATE_ADDRESS(intIndex) = .TextMatrix(intRow, intCol + 1)
                    pubDEFECT_DATA(intRow).DATA_ADDRESS(intIndex) = .TextMatrix(intRow, intCol)             '2012.03.26 Added by K.H.KIM
                    pubDEFECT_DATA(intRow).GATE_ADDRESS(intIndex) = .TextMatrix(intRow, intCol + 1)         '2012.03.26 Added by K.H.KIM
                Next intCol
                If .TextMatrix(intRow, 13) = "" Then
                    .TextMatrix(intRow, 13) = "0"
                End If
                typDEFECT_DATA(intRow).COLOR = .TextMatrix(intRow, 12)
                typDEFECT_DATA(intRow).GRAY_LEVEL = CInt(.TextMatrix(intRow, 13))
                If (Me.flxDefect_List.TextMatrix(intRow, 9) = "") And (Me.flxDefect_List.TextMatrix(intRow, 10) = "") Then
                    Call ACCUMULATE(pubCST_INFO, typDEFECT_DATA(intRow), intRow)
                    pubDEFECT_DATA(intRow).Rank = typDEFECT_DATA(intRow).Rank       '2012.03.26 Added by K.H.KIM
                    Me.flxDefect_List.TextMatrix(intRow, 14) = typDEFECT_DATA(intRow).GRADE
                    pubDEFECT_DATA(intRow).GRADE = typDEFECT_DATA(intRow).GRADE     '2012.03.26 Added by K.H.KIM
                    Call RANK_OBJ.Set_DEFECT_DATA(intRow, pubPANEL_INFO.PANELID, .TextMatrix(intRow, 0), .TextMatrix(intRow, 1), typDEFECT_DATA(intRow).DETAIL_DIVISION, _
                                                  typDEFECT_DATA(intRow).DATA_ADDRESS, typDEFECT_DATA(intRow).GATE_ADDRESS, .TextMatrix(intRow, 12), _
                                                  CInt(.TextMatrix(intRow, 13)), typDEFECT_DATA(intRow).ACCUMULATION)
                Else
                    typDEFECT_DATA(intRow).Rank = .TextMatrix(intRow, 9)
                    pubDEFECT_DATA(intRow).Rank = .TextMatrix(intRow, 9)            '2012.03.26 Added by K.H.KIM
                    Call Get_Rank_Data(pubCST_INFO.PROCESS_NUM, typRANK_DATA, typGRADE_DATA, typDEFECT_DATA(intRow).DEFECT_CODE, intGrade_Count)
                    typDEFECT_DATA(intRow).GRADE = Get_Grade_by_Rank(typGRADE_DATA, intGrade_Count, typDEFECT_DATA(intRow).DEFECT_CODE, typDEFECT_DATA(intRow).Rank)
                    Me.flxDefect_List.TextMatrix(intRow, 14) = typDEFECT_DATA(intRow).GRADE
                    pubDEFECT_DATA(intRow).GRADE = typDEFECT_DATA(intRow).GRADE     '2012.03.26 Added by K.H.KIM
                    Call RANK_OBJ.Set_DEFECT_DATA(intRow, pubPANEL_INFO.PANELID, .TextMatrix(intRow, 0), .TextMatrix(intRow, 1), .TextMatrix(intRow, 11), _
                                                  typDEFECT_DATA(intRow).DATA_ADDRESS, typDEFECT_DATA(intRow).GATE_ADDRESS, .TextMatrix(intRow, 12), _
                                                  CInt(.TextMatrix(intRow, 13)), .TextMatrix(intRow, 10))
                    typDEFECT_DATA(intRow).PRIORITY = typRANK_DATA.PRIORITY
                    If typRANK_DATA.ACCUMULATION <> "X" Then
                        'Accumulation
                        If IsNumeric(Me.flxDefect_List.TextMatrix(intRow, 10)) = True Then
                            Call Add_Point_Defect_Total(typDEFECT_DATA(intRow), CInt(Me.flxDefect_List.TextMatrix(intRow, 10)))
                            intPoint_Defect_Total = Get_Point_Defect_Total(typDEFECT_DATA(intRow).DEFECT_CODE, typDEFECT_DATA(intRow).PANELID)
                        Else
                            Call Add_Other_Defect_Data(typDEFECT_DATA(intRow))
                        End If
                    Else
                        Call Add_Other_Defect_Data(typDEFECT_DATA(intRow))
                    End If
                    If IsNumeric(Me.flxDefect_List.TextMatrix(intRow, 10)) = True Then
                        If typRANK_DATA.DETAIL_DIVISION = "B" Then
                            Call RANK_OBJ.Add_TB_Count(CInt(Me.flxDefect_List.TextMatrix(intRow, 10)))
                        ElseIf typRANK_DATA.DETAIL_DIVISION = "D" Then
                            Call RANK_OBJ.Add_TD_Count(CInt(Me.flxDefect_List.TextMatrix(intRow, 10)))
                        End If
                    Else
                        If typRANK_DATA.DETAIL_DIVISION = "B" Then
                            Call RANK_OBJ.Add_TB_Count(1)
                        ElseIf typRANK_DATA.DETAIL_DIVISION = "D" Then
                            Call RANK_OBJ.Add_TD_Count(1)
                        End If
                    End If
                    If typRANK_DATA.JUDGE_OR_NOT = "O" Then
                        strGrade = ""
                        strRank = Me.flxDefect_List.TextMatrix(intRow, 9)
                        For intIndex = 1 To intGrade_Count
                            If (strGrade = "") And (typDEFECT_DATA(intRow).DEFECT_CODE = typGRADE_DATA(intIndex).DEFECT_CODE) And _
                               (InStr(typGRADE_DATA(intIndex).Rank, strRank) > 0) Then
                                strGrade = typGRADE_DATA(intIndex).GRADE
                            End If
                        Next intIndex
                        If strGrade = "" Then
                            RANK_OBJ.Get_Highest_Grade
                        End If
                        typDEFECT_DATA(intRow).GRADE = strGrade
                        typDEFECT_DATA(intRow).Rank = strRank
                        Call SaveLog("cmdGrade_Click", typDEFECT_DATA(intRow).DEFECT_CODE & "'s RANK : " & typDEFECT_DATA(intRow).Rank & ", GRADE : " & strGrade)
                        Call Update_Defect_Grade(typDEFECT_DATA(intRow))
                        Call RANK_OBJ.Set_DEFECT_RANK(typDEFECT_DATA(intRow).DEFECT_CODE, strRank, typDEFECT_DATA(intRow).DATA_ADDRESS, typDEFECT_DATA(intRow).GATE_ADDRESS)
                        Call RANK_OBJ.Set_DEFECT_GRADE(typDEFECT_DATA(intRow).DEFECT_CODE, typDEFECT_DATA(intRow).DATA_ADDRESS, typDEFECT_DATA(intRow).GATE_ADDRESS, strGrade)
                    Else
                        If Me.flxDefect_List.Rows = 2 Then
                            strGrade = ""
                            strRank = Me.flxDefect_List.TextMatrix(intRow, 9)
                            For intIndex = 1 To intGrade_Count
                                If (strGrade = "") And (typDEFECT_DATA(intRow).DEFECT_CODE = typGRADE_DATA(intIndex).DEFECT_CODE) And _
                                   (InStr(typGRADE_DATA(intIndex).Rank, strRank) > 0) Then
                                    strGrade = typGRADE_DATA(intIndex).GRADE
                                End If
                            Next intIndex
                            If strGrade = "" Then
                                RANK_OBJ.Get_Highest_Grade
                            End If
                            typDEFECT_DATA(intRow).GRADE = strGrade
                            typDEFECT_DATA(intRow).Rank = strRank
                            Call SaveLog("cmdGrade_Click", typDEFECT_DATA(intRow).DEFECT_CODE & "'s RANK : " & typDEFECT_DATA(intRow).Rank & ", GRADE : " & strGrade)
                            Call Update_Defect_Grade(typDEFECT_DATA(intRow))
                            Call RANK_OBJ.Set_DEFECT_RANK(typDEFECT_DATA(intRow).DEFECT_CODE, strRank, typDEFECT_DATA(intRow).DATA_ADDRESS, typDEFECT_DATA(intRow).GATE_ADDRESS)
                            Call RANK_OBJ.Set_DEFECT_GRADE(typDEFECT_DATA(intRow).DEFECT_CODE, typDEFECT_DATA(intRow).DATA_ADDRESS, typDEFECT_DATA(intRow).GATE_ADDRESS, strGrade)
                        End If
                    End If
                End If
            Next intRow
            
            Call Get_Rank_Data(pubCST_INFO.PROCESS_NUM, typRANK_DATA, typGRADE_DATA, "CDBTT", intGrade_Count)
            '============Leo 2012.05.22 Add Rank Level Start
            Dim IsRanks As Boolean
            Dim Rank_SIndex As Integer
            IsRanks = False
            For intIndex = 0 To UBound(RankLevel)
                If RankLevel(intIndex) = "S" Then
                    IsRanks = True
                    Rank_SIndex = intIndex
                    Exit For
                End If
            Next intIndex
                
            If IsRanks Then
                If (typRANK_DATA.Rank(Rank_SIndex) <> "-") And (typRANK_DATA.Rank(Rank_SIndex) <> "") Then
                    If CInt(typRANK_DATA.Rank(Rank_SIndex)) < RANK_OBJ.Get_TB_Count Then
                        Call Get_Rank_Data(pubCST_INFO.PROCESS_NUM, typRANK_DATA, typGRADE_DATA, "CDDKT", intGrade_Count)
                        If RANK_OBJ.Get_TB_Count + RANK_OBJ.Get_TD_Count <= typRANK_DATA.Rank(Rank_SIndex) Then
                            If (pubCST_INFO.PROCESS_NUM = "4600") Or (pubCST_INFO.PROCESS_NUM = "4650") Then
                                strNew_Grade = "P2"
                            Else
                                strNew_Grade = "RD"
                            End If
                        End If
                    End If
                End If
            End If
            
            '============Leo 2012.05.22 Add Rank Level end
            If (strNew_Grade <> "RD") And (strNew_Grade <> "P2") Then
                Call RANK_OBJ.Init_DEFECT_PRIORITY
                For intIndex = 1 To 3
                    arrPOINT_DEFECT_COUNT(intIndex) = 0
                Next intIndex
                For intIndex = 1 To intDefect_Count
                    With typDEFECT_DATA(intIndex)
                        Call RANK_OBJ.Get_DEFECT_DATA_by_Index(intIndex, .PANELID, .DEFECT_CODE, .DEFECT_NAME, .DETAIL_DIVISION, .DATA_ADDRESS, _
                                                               .GATE_ADDRESS, .GRADE, .Rank, .COLOR, .GRAY_LEVEL, .ACCUMULATION)
                        strDEFECT_TYPE = Mid(.DEFECT_CODE, 2, 1)
                        intSource_Rank_Priority = RANK_OBJ.Get_Rank_Priority_by_Rank(.Rank)
                        intTarget_Rank_Priority = RANK_OBJ.Get_Rank_Priority_by_Rank(RANK_OBJ.Get_DEFECT_PRIORITY_RANK_by_DEFECT_TYPE(strDEFECT_TYPE))
                        If intSource_Rank_Priority > intTarget_Rank_Priority Then
                            Call RANK_OBJ.Set_DEFECT_GRADE_by_PRIORITY(strDEFECT_TYPE, .GRADE, .PRIORITY, intIndex, .DEFECT_CODE, .Rank)
                        Else
                            If intSource_Rank_Priority = intTarget_Rank_Priority Then
                                If .PRIORITY < RANK_OBJ.Get_DEFECT_PRIORITY_by_DEFECT_TYPE(strDEFECT_TYPE) Then
                                    Call RANK_OBJ.Set_DEFECT_GRADE_by_PRIORITY(strDEFECT_TYPE, .GRADE, .PRIORITY, intIndex, .DEFECT_CODE, .Rank)
                                End If
                            End If
                        End If
                        
                        'Point Defect Count Check
                        Select Case .DETAIL_DIVISION
                        Case "B":
                            arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TB) = arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TB) + .ACCUMULATION
                            arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TT) = arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TB) + arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TD)
                        Case "D":
                            arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TD) = arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TD) + .ACCUMULATION
                            arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TT) = arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TB) + arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TD)
                        End Select
                    End With
                Next intIndex
                            
                'Line Lengh calcaulation between each point defect
    '            For intIndex = 1 To 3
                intIndex = 1
                If arrPOINT_DEFECT_COUNT(intIndex) < 6 Then
                    'Reference W, L, B1 and B2 values in PFCD_Address.csv file
                    Call Get_PFCD_ADDRESS_DATA(typPFCD_ADDRESS_DATA, pubPANEL_INFO.PRODUCTID, CInt(Right(pubPANEL_INFO.PANELID, 2)))
                    With typPFCD_ADDRESS_DATA
                        Call RANK_OBJ.Calculate_Point_Distance(intIndex, .W, .L, .B1, .B2)
                    End With
                    Call RANK_OBJ.Get_Minimum_Distance(arrPOINT_DISTANCE(cDEFECT_TYPE_TB), arrPOINT_DISTANCE(cDEFECT_TYPE_TD), arrPOINT_DISTANCE(cDEFECT_TYPE_TT))
                End If
    '            Next intIndex
                
                If pubCST_INFO.PROCESS_NUM <> "3000" Then
                    'CDBTD : Bright Point Defect Minimum Distance
                    With typDEFECT_DATA(intDefect_Count + 1)
                        .PANELID = pubPANEL_INFO.PANELID
                        .DEFECT_CODE = "CDBDT"
                        Call Get_Rank_Data(pubCST_INFO.PROCESS_NUM, typRANK_DATA, typGRADE_DATA, .DEFECT_CODE, intGrade_Count)
                        Call Get_Rank(typRANK_DATA, typGRADE_DATA, intGrade_Count, strRank, strGrade, arrPOINT_DISTANCE(cDEFECT_TYPE_TB))
                        .PRIORITY = typRANK_DATA.PRIORITY
                        .Rank = strRank
                        .GRADE = strGrade
                        .ACCUMULATION = arrPOINT_DISTANCE(cDEFECT_TYPE_TB)
                        Call SaveLog("cmdGrade_Click", "CDBDT's RANK : " & strRank & ", GRADE : " & strGrade & ", Bright Distance : " & arrPOINT_DISTANCE(cDEFECT_TYPE_TB))
                        Call RANK_OBJ.Set_DEFECT_DATA(intDefect_Count + 1, pubPANEL_INFO.PANELID, .DEFECT_CODE, .DEFECT_NAME, .DETAIL_DIVISION, .DATA_ADDRESS, .GATE_ADDRESS, _
                                                      .COLOR, .GRAY_LEVEL, Trim(Str$(.ACCUMULATION)))
                        Call RANK_OBJ.Set_DEFECT_RANK(.DEFECT_CODE, .Rank, .DATA_ADDRESS, .GATE_ADDRESS)
                        Call RANK_OBJ.Set_DEFECT_GRADE(.DEFECT_CODE, .DATA_ADDRESS, .GATE_ADDRESS, .GRADE)
                    End With
                
                    'CDDKD : Dark Point Defect Minimum Distance
                    With typDEFECT_DATA(intDefect_Count + 3)
                        .PANELID = pubPANEL_INFO.PANELID
                        .DEFECT_CODE = "CDDKD"
                        Call Get_Rank_Data(pubCST_INFO.PROCESS_NUM, typRANK_DATA, typGRADE_DATA, .DEFECT_CODE, intGrade_Count)
                        Call Get_Rank(typRANK_DATA, typGRADE_DATA, intGrade_Count, strRank, strGrade, arrPOINT_DISTANCE(cDEFECT_TYPE_TD))
                        .PRIORITY = typRANK_DATA.PRIORITY
                        .Rank = strRank
                        .GRADE = strGrade
                        .ACCUMULATION = arrPOINT_DISTANCE(cDEFECT_TYPE_TD)
                        Call SaveLog("cmdGrade_Click", "CDDKD's RANK : " & strRank & ", GRADE : " & strGrade & ", Bright Distance : " & arrPOINT_DISTANCE(cDEFECT_TYPE_TD))
                        Call RANK_OBJ.Set_DEFECT_DATA(intDefect_Count + 3, pubPANEL_INFO.PANELID, .DEFECT_CODE, .DEFECT_NAME, .DETAIL_DIVISION, .DATA_ADDRESS, .GATE_ADDRESS, _
                                                      .COLOR, .GRAY_LEVEL, Trim(Str$(.ACCUMULATION)))
                        Call RANK_OBJ.Set_DEFECT_RANK(.DEFECT_CODE, .Rank, .DATA_ADDRESS, .GATE_ADDRESS)
                        Call RANK_OBJ.Set_DEFECT_GRADE(.DEFECT_CODE, .DATA_ADDRESS, .GATE_ADDRESS, .GRADE)
                    End With
                
                    'CDBDD : Point Defect Minimum Distance
                    With typDEFECT_DATA(intDefect_Count + 5)
                        .PANELID = pubPANEL_INFO.PANELID
                        .DEFECT_CODE = "CDBDD"
                        Call Get_Rank_Data(pubCST_INFO.PROCESS_NUM, typRANK_DATA, typGRADE_DATA, .DEFECT_CODE, intGrade_Count)
                        Call Get_Rank(typRANK_DATA, typGRADE_DATA, intGrade_Count, strRank, strGrade, arrPOINT_DISTANCE(cDEFECT_TYPE_TT))
                        .PRIORITY = typRANK_DATA.PRIORITY
                        .Rank = strRank
                        .GRADE = strGrade
                        .ACCUMULATION = arrPOINT_DISTANCE(cDEFECT_TYPE_TT)
                        Call SaveLog("cmdGrade_Click", "CDBDD's RANK : " & strRank & ", GRADE : " & strGrade & ", Bright Distance : " & arrPOINT_DISTANCE(cDEFECT_TYPE_TT))
                        Call RANK_OBJ.Set_DEFECT_DATA(intDefect_Count + 5, pubPANEL_INFO.PANELID, .DEFECT_CODE, .DEFECT_NAME, .DETAIL_DIVISION, .DATA_ADDRESS, .GATE_ADDRESS, _
                                                      .COLOR, .GRAY_LEVEL, Trim(Str$(.ACCUMULATION)))
                        Call RANK_OBJ.Set_DEFECT_RANK(.DEFECT_CODE, .Rank, .DATA_ADDRESS, .GATE_ADDRESS)
                        Call RANK_OBJ.Set_DEFECT_GRADE(.DEFECT_CODE, .DATA_ADDRESS, .GATE_ADDRESS, .GRADE)
                    End With
                End If
                
                'CDBTT : Bright Point Defect Total Count
                With typDEFECT_DATA(intDefect_Count + 2)
                    .PANELID = pubPANEL_INFO.PANELID
                    .DEFECT_CODE = "CDBTT"
                    Call Get_Rank_Data(pubCST_INFO.PROCESS_NUM, typRANK_DATA, typGRADE_DATA, .DEFECT_CODE, intGrade_Count)
                    Call Get_Rank(typRANK_DATA, typGRADE_DATA, intGrade_Count, strRank, strGrade, arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TB))
                    .PRIORITY = typRANK_DATA.PRIORITY
                    .Rank = strRank
                    .GRADE = strGrade
                    .ACCUMULATION = arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TB)
                    Call SaveLog("cmdGrade_Click", "CDBTT's RANK : " & strRank & ", GRADE : " & strGrade & ", Bright Total : " & .ACCUMULATION)
                    Call RANK_OBJ.Set_DEFECT_DATA(intDefect_Count + 2, pubPANEL_INFO.PANELID, .DEFECT_CODE, .DEFECT_NAME, .DETAIL_DIVISION, .DATA_ADDRESS, .GATE_ADDRESS, _
                                                  .COLOR, .GRAY_LEVEL, Trim(Str$(.ACCUMULATION)))
                    Call RANK_OBJ.Set_DEFECT_RANK(.DEFECT_CODE, .Rank, .DATA_ADDRESS, .GATE_ADDRESS)
                    Call RANK_OBJ.Set_DEFECT_GRADE(.DEFECT_CODE, .DATA_ADDRESS, .GATE_ADDRESS, .GRADE)
                End With
                
                'CDDKT : Dark Point Defect Total Count
                With typDEFECT_DATA(intDefect_Count + 4)
                    .PANELID = pubPANEL_INFO.PANELID
                    .DEFECT_CODE = "CDDKT"
                    Call Get_Rank_Data(pubCST_INFO.PROCESS_NUM, typRANK_DATA, typGRADE_DATA, .DEFECT_CODE, intGrade_Count)
                    Call Get_Rank(typRANK_DATA, typGRADE_DATA, intGrade_Count, strRank, strGrade, arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TD))
                    .PRIORITY = typRANK_DATA.PRIORITY
                    .Rank = strRank
                    .GRADE = strGrade
                    .ACCUMULATION = arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TD)
                    Call SaveLog("cmdGrade_Click", "CDDKT's RANK : " & strRank & ", GRADE : " & strGrade & ", Dark Total : " & .ACCUMULATION)
                    Call RANK_OBJ.Set_DEFECT_DATA(intDefect_Count + 4, pubPANEL_INFO.PANELID, .DEFECT_CODE, .DEFECT_NAME, .DETAIL_DIVISION, .DATA_ADDRESS, .GATE_ADDRESS, _
                                                  .COLOR, .GRAY_LEVEL, Trim(Str$(.ACCUMULATION)))
                    Call RANK_OBJ.Set_DEFECT_RANK(.DEFECT_CODE, .Rank, .DATA_ADDRESS, .GATE_ADDRESS)
                    Call RANK_OBJ.Set_DEFECT_GRADE(.DEFECT_CODE, .DATA_ADDRESS, .GATE_ADDRESS, .GRADE)
                End With
                                
                'CDBDT : Point Defect Total Count
                With typDEFECT_DATA(intDefect_Count + 6)
                    .PANELID = pubPANEL_INFO.PANELID
                    .DEFECT_CODE = "CDBDT"
                    Call Get_Rank_Data(pubCST_INFO.PROCESS_NUM, typRANK_DATA, typGRADE_DATA, .DEFECT_CODE, intGrade_Count)
                    Call Get_Rank(typRANK_DATA, typGRADE_DATA, intGrade_Count, strRank, strGrade, arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TT))
                    .PRIORITY = typRANK_DATA.PRIORITY
                    .Rank = strRank
                    .GRADE = strGrade
                    .ACCUMULATION = arrPOINT_DEFECT_COUNT(cDEFECT_TYPE_TT)
                    Call SaveLog("cmdGrade_Click", "CDBDT's RANK : " & strRank & ", GRADE : " & strGrade & ", Point Total : " & .ACCUMULATION)
                    Call RANK_OBJ.Set_DEFECT_DATA(intDefect_Count + 6, pubPANEL_INFO.PANELID, .DEFECT_CODE, .DEFECT_NAME, .DETAIL_DIVISION, .DATA_ADDRESS, .GATE_ADDRESS, _
                                                  .COLOR, .GRAY_LEVEL, Trim(Str$(.ACCUMULATION)))
                    Call RANK_OBJ.Set_DEFECT_RANK(.DEFECT_CODE, .Rank, .DATA_ADDRESS, .GATE_ADDRESS)
                    Call RANK_OBJ.Set_DEFECT_GRADE(.DEFECT_CODE, .DATA_ADDRESS, .GATE_ADDRESS, .GRADE)
                End With
        
                With typDEFECT_DATA(1)
                    Call RANK_OBJ.Get_DEFECT_DATA_by_Index(1, .PANELID, .DEFECT_CODE, .DEFECT_NAME, .DETAIL_DIVISION, .DATA_ADDRESS, .GATE_ADDRESS, _
                                                           .GRADE, .Rank, .COLOR, .GRAY_LEVEL, .ACCUMULATION)
                    If .Rank <> "" Then
                        Call RANK_OBJ.Set_DEFECT_GRADE_by_PRIORITY(Mid(.DEFECT_CODE, 2, 1), .GRADE, .PRIORITY, 1, .DEFECT_CODE, .Rank)
                    End If
                End With
                For intIndex = 2 To intDefect_Count
                    With typDEFECT_DATA(intIndex)
                        Call RANK_OBJ.Get_DEFECT_DATA_by_Index(intIndex, .PANELID, .DEFECT_CODE, .DEFECT_NAME, .DETAIL_DIVISION, .DATA_ADDRESS, .GATE_ADDRESS, _
                                                               .GRADE, .Rank, .COLOR, .GRAY_LEVEL, .ACCUMULATION)
                        If .Rank <> "" Then
                            strDEFECT_TYPE = Mid(.DEFECT_CODE, 2, 1)
                            intSource_Rank_Priority = RANK_OBJ.Get_Rank_Priority_by_Rank(.Rank)
                            intTarget_Rank_Priority = RANK_OBJ.Get_Rank_Priority_by_Rank(RANK_OBJ.Get_DEFECT_PRIORITY_RANK_by_DEFECT_TYPE(strDEFECT_TYPE))
                            If intSource_Rank_Priority > intTarget_Rank_Priority Then
                                Call RANK_OBJ.Set_DEFECT_GRADE_by_PRIORITY(strDEFECT_TYPE, .GRADE, .PRIORITY, intIndex, .DEFECT_CODE, .Rank)
                            Else
                                If intSource_Rank_Priority = intTarget_Rank_Priority Then
                                    If .PRIORITY < RANK_OBJ.Get_DEFECT_PRIORITY_by_DEFECT_TYPE(strDEFECT_TYPE) Then
                                        Call RANK_OBJ.Set_DEFECT_GRADE_by_PRIORITY(strDEFECT_TYPE, .GRADE, .PRIORITY, intIndex, .DEFECT_CODE, .Rank)
                                    End If
                                End If
                            End If
                        End If
                    End With
                Next intIndex
                For intIndex = intDefect_Count + 1 To intDefect_Count + 6
                    With typDEFECT_DATA(intIndex)
                        Call RANK_OBJ.Get_DEFECT_DATA_by_Index(intIndex, .PANELID, .DEFECT_CODE, .DEFECT_NAME, .DETAIL_DIVISION, .DATA_ADDRESS, .GATE_ADDRESS, _
                                                               .GRADE, .Rank, .COLOR, .GRAY_LEVEL, .ACCUMULATION)
                        If .Rank <> "" Then
                            strDEFECT_TYPE = Mid(.DEFECT_CODE, 2, 1)
                            intSource_Rank_Priority = RANK_OBJ.Get_Rank_Priority_by_Rank(.Rank)
                            intTarget_Rank_Priority = RANK_OBJ.Get_Rank_Priority_by_Rank(RANK_OBJ.Get_DEFECT_PRIORITY_RANK_by_DEFECT_TYPE(strDEFECT_TYPE))
                            If intSource_Rank_Priority > intTarget_Rank_Priority Then
                                Call RANK_OBJ.Set_DEFECT_GRADE_by_PRIORITY(strDEFECT_TYPE, .GRADE, .PRIORITY, intIndex, .DEFECT_CODE, .Rank)
                            Else
                                If intSource_Rank_Priority = intTarget_Rank_Priority Then
                                    If .PRIORITY < RANK_OBJ.Get_DEFECT_PRIORITY_by_DEFECT_TYPE(strDEFECT_TYPE) Then
                                        Call RANK_OBJ.Set_DEFECT_GRADE_by_PRIORITY(strDEFECT_TYPE, .GRADE, .PRIORITY, intIndex, .DEFECT_CODE, .Rank)
                                    End If
                                End If
                            End If
                        End If
                    End With
                Next intIndex
        
                strPoint_Defect_Rank = RANK_OBJ.Get_DEFECT_PRIORITY_GRADE_by_DEFECT_TYPE("D")
    '            If strPoint_Defect_Rank = "" Then
    '                If pubPANEL_INFO.TFT_REPAIR_GRADE <> "" Then
    '                    strPoint_Defect_Rank = pubPANEL_INFO.TFT_REPAIR_GRADE
    '                Else
    '                    strPoint_Defect_Rank = frmMain.lblPre_Judge.Caption
    '                End If
    '            End If
    
                strNew_Grade = Get_Panel_Grade(strPoint_Defect_Rank, strDEFECT_Rank)
                intGrade_Defect_Index = RANK_OBJ.Get_GRADE_DEFECT_INDEX
                If strNew_Grade = "" Then
                    With typGRADE_DEFECT_DATA
                        Call RANK_OBJ.Get_DEFECT_DATA_by_Index(intGrade_Defect_Index, .PANELID, .DEFECT_CODE, .DEFECT_NAME, .DETAIL_DIVISION, .DATA_ADDRESS, .GATE_ADDRESS, _
                                                               .GRADE, .Rank, .COLOR, .GRAY_LEVEL, .ACCUMULATION)
                        strNew_Grade = .GRADE
                    End With
        End If
        'Lucas 2012.04.01 Ver.0.9.19==========================For Hightest Grade
        
                End If
        Else
            'Get highst grade from Rank table
            strNew_Grade = RANK_OBJ.Get_Highest_Grade
        End If
        'Lucas 2012.04.01 Ver.0.9.19==========================For Hightest Grade
               
        
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
'        strNew_Grade = Count_Change(typDEFECT_DATA(intGrade_Defect_Index).GRADE)

    End With


    frmMain.lblPost_Judge.Caption = strNew_Grade
    With frmMain.flxMES_Data
        .TextMatrix(4, 1) = strNew_Grade
        If strNew_Grade = "RD" Then
            .TextMatrix(5, 1) = "CDBTT"
        Else
            .TextMatrix(5, 1) = typDEFECT_DATA(intGrade_Defect_Index).DEFECT_CODE
        End If
    End With
    With frmMain.flxJudge_History
        intRow = .Rows - 1
        .TextMatrix(intRow, 3) = strNew_Grade
        If strNew_Grade = "RD" Then
            .TextMatrix(intRow, 4) = "CDBTT"
        Else
            .TextMatrix(intRow, 4) = typDEFECT_DATA(intGrade_Defect_Index).DEFECT_CODE
        End If
        .TextMatrix(intRow, 5) = Get_Defect_Name(.TextMatrix(intRow, 4))
        .TextMatrix(intRow, 6) = Format(TIME, "HH:MM:SS")
    End With
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Result.mdb"

    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)

        strQuery = "UPDATE PANEL_DATA SET "
        strQuery = strQuery & "PANEL_RANK='" & strDEFECT_Rank & "', "
        strQuery = strQuery & "PANEL_GRADE='" & strNew_Grade & "', "
        strQuery = strQuery & "PANEL_LOSSCODE='" & frmMain.flxJudge_History.TextMatrix(intRow, 4) & "' WHERE "
        strQuery = strQuery & "KEYID='" & RANK_OBJ.Get_Current_KEYID & "'"

        dbMyDB.Execute (strQuery)

        dbMyDB.Close
    End If
    
'    Lucas Ver.0.9.17 2012.03.29========================For QJPG delay sending
'    msec = 0
'    Call delaytime(msec)
'    Lucas Ver.0.9.17 2012.03.29========================For QJPG delay sending
    Call Send_Panel_Judge(pubPANEL_INFO.PANELID, strNew_Grade, frmMain.flxJudge_History.TextMatrix(intRow, 4), "")

    intPTN_Index = CInt(Me.lblCurrent_PTN_Index.Caption)
    EQP.Set_PATTERN_END_by_Index (intPTN_Index)

    Unload Me
    
'    Load frmSimple_Judge
'    frmSimple_Judge.Show
    
End Sub

Private Sub cmdMain_Window_Click()

    Me.Hide
    frmMain.Show
    
End Sub

Private Sub cmdSet_Address_Click()

    Dim strCommand                      As String
    Dim strDevice_State                 As String
    
    Dim intRow                          As Integer
    Dim intSize                         As Integer
    Dim intIndex                        As Integer
    Dim intPortID                       As Integer
    
    intRow = Me.flxDefect_List.Rows - 1
    
    If Left(frmMain.flxEQ_Information.TextMatrix(5, 1), 5) = "CATST" Then
        If Mid(Me.flxDefect_List.TextMatrix(intRow, 0), 2, 1) = "M" Then
            strCommand = "RRAD"
            For intIndex = 0 To 2
                intSize = cSIZE_TOTALPIXEL_MES - Len(Trim(Me.txtX_Data(intIndex).Text))
                strCommand = strCommand & Trim(Me.txtX_Data(intIndex).Text) & Space(intSize)
                intSize = cSIZE_TOTALPIXEL_MES - Len(Trim(Me.txtY_Gate(intIndex).Text))
                strCommand = strCommand & Trim(Me.txtY_Gate(intIndex).Text) & Space(intSize)
            Next intIndex
        ElseIf Mid(Me.flxDefect_List.TextMatrix(intRow, 0), 2, 1) = "L" Then
            strCommand = "RRAD"
            For intIndex = 0 To 1
                intSize = cSIZE_TOTALPIXEL_MES - Len(Trim(Me.txtX_Data(intIndex).Text))
                strCommand = strCommand & Trim(Me.txtX_Data(intIndex).Text) & Space(intSize)
                intSize = cSIZE_TOTALPIXEL_MES - Len(Trim(Me.txtY_Gate(intIndex).Text))
                strCommand = strCommand & Trim(Me.txtY_Gate(intIndex).Text) & Space(intSize)
            Next intIndex
            strCommand = strCommand & Space(cSIZE_TOTALPIXEL_MES * 2)
        Else
            strCommand = "RRAD"
            intSize = cSIZE_TOTALPIXEL_MES - Len(Trim(Me.txtX_Data(0).Text))
            strCommand = strCommand & Trim(Me.txtX_Data(0).Text) & Space(intSize)
            intSize = cSIZE_TOTALPIXEL_MES - Len(Trim(Me.txtY_Gate(0).Text))
            strCommand = strCommand & Trim(Me.txtY_Gate(0).Text) & Space(intSize)
            strCommand = strCommand & Space(cSIZE_TOTALPIXEL_MES * 4)
        End If
        Call ENV.Get_Device_Data_by_Name("API", intPortID, strDevice_State)
        If intPortID = 0 Then
            intPortID = 9
'            Call ENV.Set_Device_Info(8, "API", "PORT OPEN")
'            Call State_Change(8, "API", cDEVICE_ONLINE)
        End If
        Call API_Sequence(intPortID, strCommand)
    ElseIf Left(frmMain.flxEQ_Information.TextMatrix(5, 1), 5) = "CALOI" Then
        If Mid(Me.flxDefect_List.TextMatrix(intRow, 0), 2, 1) = "M" Then
            strCommand = "RADC"
            For intIndex = 0 To 2
                intSize = cSIZE_TOTALPIXEL_MES - Len(Trim(Me.txtX_Data(intIndex).Text))
                strCommand = strCommand & Trim(Me.txtX_Data(intIndex).Text) & Space(intSize)
                intSize = cSIZE_TOTALPIXEL_MES - Len(Trim(Me.txtY_Gate(intIndex).Text))
                strCommand = strCommand & Trim(Me.txtY_Gate(intIndex).Text) & Space(intSize)
            Next intIndex
        ElseIf Mid(Me.flxDefect_List.TextMatrix(intRow, 0), 2, 1) = "L" Then
            strCommand = "RADC"
            For intIndex = 0 To 1
                intSize = cSIZE_TOTALPIXEL_MES - Len(Trim(Me.txtX_Data(intIndex).Text))
                strCommand = strCommand & Trim(Me.txtX_Data(intIndex).Text) & Space(intSize)
                intSize = cSIZE_TOTALPIXEL_MES - Len(Trim(Me.txtY_Gate(intIndex).Text))
                strCommand = strCommand & Trim(Me.txtY_Gate(intIndex).Text) & Space(intSize)
            Next intIndex
            strCommand = strCommand & Space(cSIZE_TOTALPIXEL_MES * 2)
        Else
            strCommand = "RADC"
            intSize = cSIZE_TOTALPIXEL_MES - Len(Trim(Me.txtX_Data(0).Text))
            strCommand = strCommand & Trim(Me.txtX_Data(0).Text) & Space(intSize)
            intSize = cSIZE_TOTALPIXEL_MES - Len(Trim(Me.txtY_Gate(0).Text))
            strCommand = strCommand & Trim(Me.txtY_Gate(0).Text) & Space(intSize)
            strCommand = strCommand & Space(cSIZE_TOTALPIXEL_MES * 4)
        End If
        Call ENV.Get_Device_Data_by_Name("CALOI", intPortID, strDevice_State)
        Call BLOI_Sequence(intPortID, strCommand)
    End If
    
End Sub

Private Sub flxDefect_A_DblClick()

    Dim typPATTERN_DATA     As PATTERN_LIST_DATA
    
    Dim intRow              As Integer
    Dim intIndex            As Integer
    Dim intPTN_Index        As Integer
    
    Dim strDefect_Code      As String
    Dim strDEFECT_NAME      As String
    Dim strDEFECT_KIND      As String
    Dim strAddress_Count    As String
    
    intRow = Me.flxDefect_A.Row
    
    If intRow > 0 Then
        If Get_Defect_Type(Me.flxDefect_A.TextMatrix(intRow, 0)) <> "A" Then
            strDefect_Code = Me.flxDefect_A.TextMatrix(intRow, 0)
            strDEFECT_NAME = Me.flxDefect_A.TextMatrix(intRow, 1)
            strDEFECT_KIND = Mid(strDefect_Code, 2, 1)
            
            For intIndex = 0 To 2
                Me.txtX_Data(intIndex).Enabled = False
                Me.txtY_Gate(intIndex).Enabled = False
                Me.txtX_Data(intIndex).BackColor = vbWhite
                Me.txtY_Gate(intIndex).BackColor = vbWhite
            Next intIndex
            Me.txtX_Data(0).Enabled = True
            Me.txtY_Gate(0).Enabled = True
            
            intRow = Add_Grid(strDefect_Code, Me.flxDefect_List)
            With Me.flxDefect_List
                .TextMatrix(intRow, 1) = strDEFECT_NAME
                .TextMatrix(intRow, 8) = frmMain.flxMES_Data.TextMatrix(17, 1)
                
                intPTN_Index = CInt(Me.lblCurrent_PTN_Index.Caption)
                With typPATTERN_DATA
                    Call EQP.Get_PATTERN_LIST_by_Index(intPTN_Index, .PATTERN_CODE, .PATTERN_NAME, .DELAY_TIME, .LEVEL, .DH, .DL, .VGH, .VGL, .RESCUE_HIGH, .RESCUE_LOW, .VCOM)
                End With
                .TextMatrix(intRow, 12) = typPATTERN_DATA.PATTERN_NAME
            End With
            
            m_DEFECT_CODE = ""
            m_DEFECT_NAME = ""
            m_DEFECT_KIND = ""
            
            strAddress_Count = Get_Defect_Address_Count(strDefect_Code)
            If pubCST_INFO.PROCESS_NUM <> "3000" Then
                Select Case strAddress_Count
                Case "X":
                    Call Load_Manual_Judge
                Case "0":
                    For intIndex = 0 To 2
                        Me.txtX_Data(intIndex).BackColor = vbBlack
                        Me.txtY_Gate(intIndex).BackColor = vbBlack
                    Next intIndex
                    Call Set_Interlock
                Case "1":
                    For intIndex = 1 To 2
                        Me.txtX_Data(intIndex).BackColor = vbBlack
                        Me.txtY_Gate(intIndex).BackColor = vbBlack
                    Next intIndex
                    Call Set_Interlock
                Case "2":
                    Me.txtX_Data(2).BackColor = vbBlack
                    Me.txtY_Gate(2).BackColor = vbBlack
                    Call Set_Interlock
                Case "3":
                    Call Set_Interlock
                End Select
            End If
            If strAddress_Count = "X" Then
                Call Load_Manual_Judge
            End If
                     
        End If
    End If
    
End Sub

Private Sub flxDefect_A_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim intRow              As Integer
    
    If Button = vbRightButton Then
'        Me.cmdGrade.SetFocus
    Else
        intRow = Me.flxDefect_A.Row
        
        m_DEFECT_CODE = Me.flxDefect_A.TextMatrix(intRow, 0)
        m_DEFECT_NAME = Me.flxDefect_A.TextMatrix(intRow, 1)
        m_DEFECT_KIND = Mid(m_DEFECT_CODE, 2, 1)
    End If
    
End Sub

Private Sub flxDefect_B_DblClick()

    Dim typPATTERN_DATA     As PATTERN_LIST_DATA
    
    Dim intRow              As Integer
    Dim intIndex            As Integer
    Dim intPTN_Index        As Integer
    
    Dim strDefect_Code      As String
    Dim strDEFECT_NAME      As String
    Dim strDEFECT_KIND      As String
    Dim strAddress_Count    As String
    
    intRow = Me.flxDefect_B.Row
    
    If intRow > 0 Then
        If Get_Defect_Type(Me.flxDefect_B.TextMatrix(intRow, 0)) <> "A" Then
            strDefect_Code = Me.flxDefect_B.TextMatrix(intRow, 0)
            strDEFECT_NAME = Me.flxDefect_B.TextMatrix(intRow, 1)
            strDEFECT_KIND = Mid(strDefect_Code, 2, 1)
        
            For intIndex = 0 To 2
                Me.txtX_Data(intIndex).Enabled = False
                Me.txtY_Gate(intIndex).Enabled = False
                Me.txtX_Data(intIndex).BackColor = vbWhite
                Me.txtY_Gate(intIndex).BackColor = vbWhite
            Next intIndex
            Me.txtX_Data(0).Enabled = True
            Me.txtY_Gate(0).Enabled = True
            Me.txtX_Data(1).Enabled = True
            Me.txtY_Gate(1).Enabled = True
            
            intRow = Add_Grid(strDefect_Code, Me.flxDefect_List)
            With Me.flxDefect_List
                .TextMatrix(intRow, 1) = strDEFECT_NAME
                .TextMatrix(intRow, 8) = frmMain.flxMES_Data.TextMatrix(17, 1)
            
                intPTN_Index = CInt(Me.lblCurrent_PTN_Index.Caption)
                With typPATTERN_DATA
                    Call EQP.Get_PATTERN_LIST_by_Index(intPTN_Index, .PATTERN_CODE, .PATTERN_NAME, .DELAY_TIME, .LEVEL, .DH, .DL, .VGH, .VGL, .RESCUE_HIGH, .RESCUE_LOW, .VCOM)
                End With
                .TextMatrix(intRow, 12) = typPATTERN_DATA.PATTERN_NAME
            End With
            
            m_DEFECT_CODE = ""
            m_DEFECT_NAME = ""
            m_DEFECT_KIND = ""
        
            strAddress_Count = Get_Defect_Address_Count(strDefect_Code)
            If pubCST_INFO.PROCESS_NUM <> "3000" Then
                Select Case strAddress_Count
                Case "X":
                     Call Load_Manual_Judge
                Case "0":
                    For intIndex = 0 To 2
                        Me.txtX_Data(intIndex).BackColor = vbBlack
                        Me.txtY_Gate(intIndex).BackColor = vbBlack
                    Next intIndex
                    Call Set_Interlock
                Case "1":
                    For intIndex = 1 To 2
                        Me.txtX_Data(intIndex).BackColor = vbBlack
                        Me.txtY_Gate(intIndex).BackColor = vbBlack
                    Next intIndex
                    Call Set_Interlock
                Case "2":
                    Me.txtX_Data(2).BackColor = vbBlack
                    Me.txtY_Gate(2).BackColor = vbBlack
                    Call Set_Interlock
                Case "3":
                    Call Set_Interlock
                End Select
            End If
                     If strAddress_Count = "X" Then
                       Call Load_Manual_Judge
                     End If
        End If
    End If
    
End Sub

Private Sub flxDefect_B_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim intRow              As Integer
    
    If Button = vbRightButton Then
'        Me.cmdGrade.SetFocus
    Else
        intRow = Me.flxDefect_B.Row
        
        m_DEFECT_CODE = Me.flxDefect_B.TextMatrix(intRow, 0)
        m_DEFECT_NAME = Me.flxDefect_B.TextMatrix(intRow, 1)
        m_DEFECT_KIND = Mid(m_DEFECT_CODE, 2, 1)
    End If

End Sub

Private Sub flxDefect_C_DblClick()

    Dim typPATTERN_DATA     As PATTERN_LIST_DATA
    
    Dim intRow              As Integer
    Dim intIndex            As Integer
    Dim intPTN_Index        As Integer
    
    Dim strDefect_Code      As String
    Dim strDEFECT_NAME      As String
    Dim strDEFECT_KIND      As String
    Dim strAddress_Count    As String
    
    intRow = Me.flxDefect_C.Row
    
    If intRow > 0 Then
        If Get_Defect_Type(Me.flxDefect_C.TextMatrix(intRow, 0)) <> "A" Then
            strDefect_Code = Me.flxDefect_C.TextMatrix(intRow, 0)
            strDEFECT_NAME = Me.flxDefect_C.TextMatrix(intRow, 1)
            strDEFECT_KIND = Mid(strDefect_Code, 2, 1)
        
            For intIndex = 0 To 2
                Me.txtX_Data(intIndex).Enabled = False
                Me.txtY_Gate(intIndex).Enabled = False
                Me.txtX_Data(intIndex).BackColor = vbWhite
                Me.txtY_Gate(intIndex).BackColor = vbWhite
            Next intIndex
            Me.txtX_Data(0).Enabled = True
            Me.txtY_Gate(0).Enabled = True
        
            intRow = Add_Grid(strDefect_Code, Me.flxDefect_List)
            With Me.flxDefect_List
                .TextMatrix(intRow, 1) = strDEFECT_NAME
                .TextMatrix(intRow, 8) = frmMain.flxMES_Data.TextMatrix(17, 1)
            
                intPTN_Index = CInt(Me.lblCurrent_PTN_Index.Caption)
                With typPATTERN_DATA
                    Call EQP.Get_PATTERN_LIST_by_Index(intPTN_Index, .PATTERN_CODE, .PATTERN_NAME, .DELAY_TIME, .LEVEL, .DH, .DL, .VGH, .VGL, .RESCUE_HIGH, .RESCUE_LOW, .VCOM)
                End With
                .TextMatrix(intRow, 12) = typPATTERN_DATA.PATTERN_NAME
            End With
            
            m_DEFECT_CODE = ""
            m_DEFECT_NAME = ""
            m_DEFECT_KIND = ""
        
            strAddress_Count = Get_Defect_Address_Count(strDefect_Code)
            If pubCST_INFO.PROCESS_NUM <> "3000" Then
                Select Case strAddress_Count
                Case "X":
                     Call Load_Manual_Judge
                Case "0":
                    For intIndex = 0 To 2
                        Me.txtX_Data(intIndex).BackColor = vbBlack
                        Me.txtY_Gate(intIndex).BackColor = vbBlack
                    Next intIndex
                    Call Set_Interlock
                Case "1":
                    For intIndex = 1 To 2
                        Me.txtX_Data(intIndex).BackColor = vbBlack
                        Me.txtY_Gate(intIndex).BackColor = vbBlack
                    Next intIndex
                    Call Set_Interlock
                Case "2":
                    Me.txtX_Data(2).BackColor = vbBlack
                    Me.txtY_Gate(2).BackColor = vbBlack
                    Call Set_Interlock
                Case "3":
                    Call Set_Interlock
                End Select
            End If
                     If strAddress_Count = "X" Then
                       Call Load_Manual_Judge
                     End If
        End If
    End If
    
End Sub

Private Sub flxDefect_C_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim intRow              As Integer
    
    If Button = vbRightButton Then
'        Me.cmdGrade.SetFocus
    Else
        intRow = Me.flxDefect_C.Row
        
        m_DEFECT_CODE = Me.flxDefect_C.TextMatrix(intRow, 0)
        m_DEFECT_NAME = Me.flxDefect_C.TextMatrix(intRow, 1)
        m_DEFECT_KIND = Mid(m_DEFECT_CODE, 2, 1)
    End If

End Sub

Private Sub flxDefect_D_DblClick()

    Dim typPATTERN_DATA     As PATTERN_LIST_DATA
    
    Dim intRow              As Integer
    Dim intIndex            As Integer
    Dim intPTN_Index        As Integer
    
    Dim strDefect_Code      As String
    Dim strDEFECT_NAME      As String
    Dim strDEFECT_KIND      As String
    Dim strAddress_Count    As String
    
    intRow = Me.flxDefect_D.Row
    
    If intRow > 0 Then
        If Get_Defect_Type(Me.flxDefect_D.TextMatrix(intRow, 0)) <> "A" Then
            strDefect_Code = Me.flxDefect_D.TextMatrix(intRow, 0)
            strDEFECT_NAME = Me.flxDefect_D.TextMatrix(intRow, 1)
            strDEFECT_KIND = Mid(strDefect_Code, 2, 1)
        
            For intIndex = 0 To 2
                Me.txtX_Data(intIndex).Enabled = True
                Me.txtY_Gate(intIndex).Enabled = True
                Me.txtX_Data(intIndex).BackColor = vbWhite
                Me.txtY_Gate(intIndex).BackColor = vbWhite
            Next intIndex
            
            intRow = Add_Grid(strDefect_Code, Me.flxDefect_List)
            With Me.flxDefect_List
                .TextMatrix(intRow, 1) = strDEFECT_NAME
                .TextMatrix(intRow, 8) = frmMain.flxMES_Data.TextMatrix(17, 1)
            
                intPTN_Index = CInt(Me.lblCurrent_PTN_Index.Caption)
                With typPATTERN_DATA
                    Call EQP.Get_PATTERN_LIST_by_Index(intPTN_Index, .PATTERN_CODE, .PATTERN_NAME, .DELAY_TIME, .LEVEL, .DH, .DL, .VGH, .VGL, .RESCUE_HIGH, .RESCUE_LOW, .VCOM)
                End With
                .TextMatrix(intRow, 12) = typPATTERN_DATA.PATTERN_NAME
            End With
            
            m_DEFECT_CODE = ""
            m_DEFECT_NAME = ""
            m_DEFECT_KIND = ""
        
            strAddress_Count = Get_Defect_Address_Count(strDefect_Code)
            If pubCST_INFO.PROCESS_NUM <> "3000" Then
                Select Case strAddress_Count
                Case "X":
                     Call Load_Manual_Judge
                Case "0":
                    For intIndex = 0 To 2
                        Me.txtX_Data(intIndex).BackColor = vbBlack
                        Me.txtY_Gate(intIndex).BackColor = vbBlack
                    Next intIndex
                    Call Set_Interlock
                Case "1":
                    For intIndex = 1 To 2
                        Me.txtX_Data(intIndex).BackColor = vbBlack
                        Me.txtY_Gate(intIndex).BackColor = vbBlack
                    Next intIndex
                    Call Set_Interlock
                Case "2":
                    Me.txtX_Data(2).BackColor = vbBlack
                    Me.txtY_Gate(2).BackColor = vbBlack
                    Call Set_Interlock
                Case "3":
                    Call Set_Interlock
                End Select
            End If
                      If strAddress_Count = "X" Then
                          Call Load_Manual_Judge
                      End If
        End If
    End If
    
End Sub

Private Sub flxDefect_D_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim intRow              As Integer
    
    If Button = vbRightButton Then
'        Me.cmdGrade.SetFocus
    Else
        intRow = Me.flxDefect_D.Row
        
        m_DEFECT_CODE = Me.flxDefect_D.TextMatrix(intRow, 0)
        m_DEFECT_NAME = Me.flxDefect_D.TextMatrix(intRow, 1)
        m_DEFECT_KIND = Mid(m_DEFECT_CODE, 2, 1)
    End If

End Sub

Private Sub flxDefect_E_DblClick()

    Dim typPATTERN_DATA     As PATTERN_LIST_DATA
    
    Dim intRow              As Integer
    Dim intIndex            As Integer
    Dim intPTN_Index        As Integer
    
    Dim strDefect_Code      As String
    Dim strDEFECT_NAME      As String
    Dim strDEFECT_KIND      As String
    Dim strAddress_Count    As String
    
    intRow = Me.flxDefect_E.Row
    
    If intRow > 0 Then
        If Get_Defect_Type(Me.flxDefect_E.TextMatrix(intRow, 0)) <> "A" Then
            strDefect_Code = Me.flxDefect_E.TextMatrix(intRow, 0)
            strDEFECT_NAME = Me.flxDefect_E.TextMatrix(intRow, 1)
            strDEFECT_KIND = Mid(strDefect_Code, 2, 1)
        
            For intIndex = 0 To 2
                Me.txtX_Data(intIndex).Enabled = False
                Me.txtY_Gate(intIndex).Enabled = False
                Me.txtX_Data(intIndex).BackColor = vbWhite
                Me.txtY_Gate(intIndex).BackColor = vbWhite
            Next intIndex
            Me.txtX_Data(0).Enabled = True
            Me.txtY_Gate(0).Enabled = True
        
            intRow = Add_Grid(strDefect_Code, Me.flxDefect_List)
            With Me.flxDefect_List
                .TextMatrix(intRow, 1) = strDEFECT_NAME
                .TextMatrix(intRow, 8) = frmMain.flxMES_Data.TextMatrix(17, 1)
            
                intPTN_Index = CInt(Me.lblCurrent_PTN_Index.Caption)
                With typPATTERN_DATA
                    Call EQP.Get_PATTERN_LIST_by_Index(intPTN_Index, .PATTERN_CODE, .PATTERN_NAME, .DELAY_TIME, .LEVEL, .DH, .DL, .VGH, .VGL, .RESCUE_HIGH, .RESCUE_LOW, .VCOM)
                End With
                .TextMatrix(intRow, 12) = typPATTERN_DATA.PATTERN_NAME
            End With
            
            m_DEFECT_CODE = ""
            m_DEFECT_NAME = ""
            m_DEFECT_KIND = ""
        
            strAddress_Count = Get_Defect_Address_Count(strDefect_Code)
            If pubCST_INFO.PROCESS_NUM <> "3000" Then
                Select Case strAddress_Count
                Case "X":
                    Call Load_Manual_Judge
                Case "0":
                    For intIndex = 0 To 2
                        Me.txtX_Data(intIndex).BackColor = vbBlack
                        Me.txtY_Gate(intIndex).BackColor = vbBlack
                    Next intIndex
                    Call Set_Interlock
                Case "1":
                    For intIndex = 1 To 2
                        Me.txtX_Data(intIndex).BackColor = vbBlack
                        Me.txtY_Gate(intIndex).BackColor = vbBlack
                    Next intIndex
                    Call Set_Interlock
                Case "2":
                    Me.txtX_Data(2).BackColor = vbBlack
                    Me.txtY_Gate(2).BackColor = vbBlack
                    Call Set_Interlock
                Case "3":
                    Call Set_Interlock
                End Select
             End If
                      If strAddress_Count = "X" Then
                        Call Load_Manual_Judge
                      End If
        End If
    End If
    
End Sub

Private Sub flxDefect_E_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim intRow              As Integer
    
    If Button = vbRightButton Then
'        Me.cmdGrade.SetFocus
    Else
        intRow = Me.flxDefect_E.Row
        
        m_DEFECT_CODE = Me.flxDefect_E.TextMatrix(intRow, 0)
        m_DEFECT_NAME = Me.flxDefect_E.TextMatrix(intRow, 1)
        m_DEFECT_KIND = Mid(m_DEFECT_CODE, 2, 1)
    End If

End Sub

Private Sub flxDefect_F_DblClick()

    Dim typPATTERN_DATA     As PATTERN_LIST_DATA
    
    Dim intRow              As Integer
    Dim intIndex            As Integer
    Dim intPTN_Index        As Integer
    
    Dim strDefect_Code      As String
    Dim strDEFECT_NAME      As String
    Dim strDEFECT_KIND      As String
    Dim strAddress_Count    As String
    
    intRow = Me.flxDefect_F.Row
    
    If intRow > 0 Then
        If Get_Defect_Type(Me.flxDefect_F.TextMatrix(intRow, 0)) <> "A" Then
            strDefect_Code = Me.flxDefect_F.TextMatrix(intRow, 0)
            strDEFECT_NAME = Me.flxDefect_F.TextMatrix(intRow, 1)
            strDEFECT_KIND = Mid(strDefect_Code, 2, 1)
        
            For intIndex = 0 To 2
                Me.txtX_Data(intIndex).Enabled = False
                Me.txtY_Gate(intIndex).Enabled = False
                Me.txtX_Data(intIndex).BackColor = vbWhite
                Me.txtY_Gate(intIndex).BackColor = vbWhite
            Next intIndex
            Me.txtX_Data(0).Enabled = True
            Me.txtY_Gate(0).Enabled = True
        
            intRow = Add_Grid(strDefect_Code, Me.flxDefect_List)
            With Me.flxDefect_List
                .TextMatrix(intRow, 1) = strDEFECT_NAME
                .TextMatrix(intRow, 8) = frmMain.flxMES_Data.TextMatrix(17, 1)
            
                intPTN_Index = CInt(Me.lblCurrent_PTN_Index.Caption)
                With typPATTERN_DATA
                    Call EQP.Get_PATTERN_LIST_by_Index(intPTN_Index, .PATTERN_CODE, .PATTERN_NAME, .DELAY_TIME, .LEVEL, .DH, .DL, .VGH, .VGL, .RESCUE_HIGH, .RESCUE_LOW, .VCOM)
                End With
                .TextMatrix(intRow, 12) = typPATTERN_DATA.PATTERN_NAME
            End With
            
            m_DEFECT_CODE = ""
            m_DEFECT_NAME = ""
            m_DEFECT_KIND = ""
        
            strAddress_Count = Get_Defect_Address_Count(strDefect_Code)
            If pubCST_INFO.PROCESS_NUM <> "3000" Then
                Select Case strAddress_Count
                Case "X":
                     Call Load_Manual_Judge
                Case "0":
                    For intIndex = 0 To 2
                        Me.txtX_Data(intIndex).BackColor = vbBlack
                        Me.txtY_Gate(intIndex).BackColor = vbBlack
                    Next intIndex
                    Call Set_Interlock
                Case "1":
                    For intIndex = 1 To 2
                        Me.txtX_Data(intIndex).BackColor = vbBlack
                        Me.txtY_Gate(intIndex).BackColor = vbBlack
                    Next intIndex
                    Call Set_Interlock
                Case "2":
                    Me.txtX_Data(2).BackColor = vbBlack
                    Me.txtY_Gate(2).BackColor = vbBlack
                    Call Set_Interlock
                Case "3":
                    Call Set_Interlock
                End Select
            End If
                      If strAddress_Count = "X" Then
                        Call Load_Manual_Judge
                      End If
        End If
    End If
    
End Sub

Private Sub flxDefect_F_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim intRow              As Integer
    
    If Button = vbRightButton Then
'        Me.cmdGrade.SetFocus
    Else
        intRow = Me.flxDefect_F.Row
        
        m_DEFECT_CODE = Me.flxDefect_F.TextMatrix(intRow, 0)
        m_DEFECT_NAME = Me.flxDefect_F.TextMatrix(intRow, 1)
        m_DEFECT_KIND = Mid(m_DEFECT_CODE, 2, 1)
    End If

End Sub

Private Sub flxDefect_G_DblClick()

    Dim typPATTERN_DATA     As PATTERN_LIST_DATA
    
    Dim intRow              As Integer
    Dim intIndex            As Integer
    Dim intPTN_Index        As Integer
    
    Dim strDefect_Code      As String
    Dim strDEFECT_NAME      As String
    Dim strDEFECT_KIND      As String
    Dim strAddress_Count    As String
    
    intRow = Me.flxDefect_G.Row
    
    If intRow > 0 Then
        If Get_Defect_Type(Me.flxDefect_G.TextMatrix(intRow, 0)) <> "A" Then
            strDefect_Code = Me.flxDefect_G.TextMatrix(intRow, 0)
            strDEFECT_NAME = Me.flxDefect_G.TextMatrix(intRow, 1)
            strDEFECT_KIND = Mid(strDefect_Code, 2, 1)
        
            For intIndex = 0 To 2
                Me.txtX_Data(intIndex).Enabled = False
                Me.txtY_Gate(intIndex).Enabled = False
                Me.txtX_Data(intIndex).BackColor = vbWhite
                Me.txtY_Gate(intIndex).BackColor = vbWhite
            Next intIndex
            Me.txtX_Data(0).Enabled = True
            Me.txtY_Gate(0).Enabled = True
        
            intRow = Add_Grid(strDefect_Code, Me.flxDefect_List)
            With Me.flxDefect_List
                .TextMatrix(intRow, 1) = strDEFECT_NAME
                .TextMatrix(intRow, 8) = frmMain.flxMES_Data.TextMatrix(17, 1)
            
                intPTN_Index = CInt(Me.lblCurrent_PTN_Index.Caption)
                With typPATTERN_DATA
                    Call EQP.Get_PATTERN_LIST_by_Index(intPTN_Index, .PATTERN_CODE, .PATTERN_NAME, .DELAY_TIME, .LEVEL, .DH, .DL, .VGH, .VGL, .RESCUE_HIGH, .RESCUE_LOW, .VCOM)
                End With
                .TextMatrix(intRow, 12) = typPATTERN_DATA.PATTERN_NAME
            End With
            
            m_DEFECT_CODE = ""
            m_DEFECT_NAME = ""
            m_DEFECT_KIND = ""
        
            strAddress_Count = Get_Defect_Address_Count(strDefect_Code)
            If pubCST_INFO.PROCESS_NUM <> "3000" Then
                Select Case strAddress_Count
                Case "X":
                    Call Load_Manual_Judge
                Case "0":
                    For intIndex = 0 To 2
                        Me.txtX_Data(intIndex).BackColor = vbBlack
                        Me.txtY_Gate(intIndex).BackColor = vbBlack
                    Next intIndex
                    Call Set_Interlock
                Case "1":
                    For intIndex = 1 To 2
                        Me.txtX_Data(intIndex).BackColor = vbBlack
                        Me.txtY_Gate(intIndex).BackColor = vbBlack
                    Next intIndex
                    Call Set_Interlock
                Case "2":
                    Me.txtX_Data(2).BackColor = vbBlack
                    Me.txtY_Gate(2).BackColor = vbBlack
                    Call Set_Interlock
                Case "3":
                    Call Set_Interlock
                End Select
            End If
                    If strAddress_Count = "X" Then
                       Call Load_Manual_Judge
                    End If
        End If
    End If
    
End Sub

Private Sub flxDefect_G_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim intRow              As Integer
    
    If Button = vbRightButton Then
'        Me.cmdGrade.SetFocus
    Else
        intRow = Me.flxDefect_G.Row
        
        m_DEFECT_CODE = Me.flxDefect_G.TextMatrix(intRow, 0)
        m_DEFECT_NAME = Me.flxDefect_G.TextMatrix(intRow, 1)
        m_DEFECT_KIND = Mid(m_DEFECT_CODE, 2, 1)
    End If

End Sub

Private Sub flxDefect_H_DblClick()

    Dim typPATTERN_DATA     As PATTERN_LIST_DATA
    
    Dim intRow              As Integer
    Dim intIndex            As Integer
    Dim intPTN_Index        As Integer
    
    Dim strDefect_Code      As String
    Dim strDEFECT_NAME      As String
    Dim strDEFECT_KIND      As String
    Dim strAddress_Count    As String
    
    intRow = Me.flxDefect_H.Row
    
    If intRow > 0 Then
        If Get_Defect_Type(Me.flxDefect_H.TextMatrix(intRow, 0)) <> "A" Then
            strDefect_Code = Me.flxDefect_H.TextMatrix(intRow, 0)
            strDEFECT_NAME = Me.flxDefect_H.TextMatrix(intRow, 1)
            strDEFECT_KIND = Mid(strDefect_Code, 2, 1)
        
            For intIndex = 0 To 2
                Me.txtX_Data(intIndex).Enabled = False
                Me.txtY_Gate(intIndex).Enabled = False
                Me.txtX_Data(intIndex).BackColor = vbWhite
                Me.txtY_Gate(intIndex).BackColor = vbWhite
            Next intIndex
            Me.txtX_Data(0).Enabled = True
            Me.txtY_Gate(0).Enabled = True
        
            intRow = Add_Grid(strDefect_Code, Me.flxDefect_List)
            With Me.flxDefect_List
                .TextMatrix(intRow, 1) = strDEFECT_NAME
                .TextMatrix(intRow, 8) = frmMain.flxMES_Data.TextMatrix(17, 1)
            
                intPTN_Index = CInt(Me.lblCurrent_PTN_Index.Caption)
                With typPATTERN_DATA
                    Call EQP.Get_PATTERN_LIST_by_Index(intPTN_Index, .PATTERN_CODE, .PATTERN_NAME, .DELAY_TIME, .LEVEL, .DH, .DL, .VGH, .VGL, .RESCUE_HIGH, .RESCUE_LOW, .VCOM)
                End With
                .TextMatrix(intRow, 12) = typPATTERN_DATA.PATTERN_NAME
            End With
            
            m_DEFECT_CODE = ""
            m_DEFECT_NAME = ""
            m_DEFECT_KIND = ""
        
            strAddress_Count = Get_Defect_Address_Count(strDefect_Code)
            If pubCST_INFO.PROCESS_NUM <> "3000" Then
                Select Case strAddress_Count
                Case "X":
                     Call Load_Manual_Judge
                Case "0":
                    For intIndex = 0 To 2
                        Me.txtX_Data(intIndex).BackColor = vbBlack
                        Me.txtY_Gate(intIndex).BackColor = vbBlack
                    Next intIndex
                    Call Set_Interlock
                Case "1":
                    For intIndex = 1 To 2
                        Me.txtX_Data(intIndex).BackColor = vbBlack
                        Me.txtY_Gate(intIndex).BackColor = vbBlack
                    Next intIndex
                    Call Set_Interlock
                Case "2":
                    Me.txtX_Data(2).BackColor = vbBlack
                    Me.txtY_Gate(2).BackColor = vbBlack
                    Call Set_Interlock
                Case "3":
                    Call Set_Interlock
                End Select
            End If
           If strAddress_Count = "X" Then
              Call Load_Manual_Judge
           End If
        End If
    End If
    
End Sub

Private Sub flxDefect_H_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim intRow              As Integer
    
    If Button = vbRightButton Then
'        Me.cmdGrade.SetFocus
    Else
        intRow = Me.flxDefect_H.Row
        
        m_DEFECT_CODE = Me.flxDefect_H.TextMatrix(intRow, 0)
        m_DEFECT_NAME = Me.flxDefect_H.TextMatrix(intRow, 1)
        m_DEFECT_KIND = Mid(m_DEFECT_CODE, 2, 1)
    End If

End Sub

Private Sub flxDefect_I_DblClick()

    Dim typPATTERN_DATA     As PATTERN_LIST_DATA
    
    Dim intRow              As Integer
    Dim intIndex            As Integer
    Dim intPTN_Index        As Integer
    
    Dim strDefect_Code      As String
    Dim strDEFECT_NAME      As String
    Dim strDEFECT_KIND      As String
    Dim strAddress_Count    As String
    
    intRow = Me.flxDefect_I.Row
    
    If intRow > 0 Then
        If Get_Defect_Type(Me.flxDefect_I.TextMatrix(intRow, 0)) <> "A" Then
            strDefect_Code = Me.flxDefect_I.TextMatrix(intRow, 0)
            strDEFECT_NAME = Me.flxDefect_I.TextMatrix(intRow, 1)
            strDEFECT_KIND = Mid(strDefect_Code, 2, 1)
        
            For intIndex = 0 To 2
                Me.txtX_Data(intIndex).Enabled = False
                Me.txtY_Gate(intIndex).Enabled = False
                Me.txtX_Data(intIndex).BackColor = vbWhite
                Me.txtY_Gate(intIndex).BackColor = vbWhite
            Next intIndex
            Me.txtX_Data(0).Enabled = True
            Me.txtY_Gate(0).Enabled = True
        
            intRow = Add_Grid(strDefect_Code, Me.flxDefect_List)
            With Me.flxDefect_List
                .TextMatrix(intRow, 1) = strDEFECT_NAME
                .TextMatrix(intRow, 8) = frmMain.flxMES_Data.TextMatrix(17, 1)
            
                intPTN_Index = CInt(Me.lblCurrent_PTN_Index.Caption)
                With typPATTERN_DATA
                    Call EQP.Get_PATTERN_LIST_by_Index(intPTN_Index, .PATTERN_CODE, .PATTERN_NAME, .DELAY_TIME, .LEVEL, .DH, .DL, .VGH, .VGL, .RESCUE_HIGH, .RESCUE_LOW, .VCOM)
                End With
                .TextMatrix(intRow, 12) = typPATTERN_DATA.PATTERN_NAME
            End With
            
            m_DEFECT_CODE = ""
            m_DEFECT_NAME = ""
            m_DEFECT_KIND = ""
                    
            strAddress_Count = Get_Defect_Address_Count(strDefect_Code)
            If pubCST_INFO.PROCESS_NUM <> "3000" Then
                Select Case strAddress_Count
                Case "X":
                     Call Load_Manual_Judge
                Case "0":
                    For intIndex = 0 To 2
                        Me.txtX_Data(intIndex).BackColor = vbBlack
                        Me.txtY_Gate(intIndex).BackColor = vbBlack
                    Next intIndex
                    Call Set_Interlock
                Case "1":
                    For intIndex = 1 To 2
                        Me.txtX_Data(intIndex).BackColor = vbBlack
                        Me.txtY_Gate(intIndex).BackColor = vbBlack
                    Next intIndex
                    Call Set_Interlock
                Case "2":
                    Me.txtX_Data(2).BackColor = vbBlack
                    Me.txtY_Gate(2).BackColor = vbBlack
                    Call Set_Interlock
                Case "3":
                    Call Set_Interlock
                End Select
            End If
            If strAddress_Count = "X" Then
              Call Load_Manual_Judge
            End If
        End If
    End If
    
End Sub

Private Sub flxDefect_I_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim intRow              As Integer
    
    If Button = vbRightButton Then
'        Me.cmdGrade.SetFocus
    Else
        intRow = Me.flxDefect_I.Row
        
        m_DEFECT_CODE = Me.flxDefect_I.TextMatrix(intRow, 0)
        m_DEFECT_NAME = Me.flxDefect_I.TextMatrix(intRow, 1)
        m_DEFECT_KIND = Mid(m_DEFECT_CODE, 2, 1)
    End If

End Sub

Private Sub flxDefect_List_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

'    If Me.cmdGrade.Enabled = True Then
'        If Button = vbRightButton Then
'            frmJudge.cmdGrade.SetFocus
'        End If
'    End If

End Sub

Private Sub flxPG_Data_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

'    If Me.cmdGrade.Enabled = True Then
'        If Button = vbRightButton Then
'            frmJudge.cmdGrade.SetFocus
'        End If
'    End If

End Sub

Private Sub Form_Activate()

    frmMain.cmdJudge.Enabled = True
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If Me.picCurrent_Pattern.Enabled = True Then
        If KeyCode = vbKeySpace Then
            Call picCurrent_Pattern_Click
        End If
    End If
    
End Sub

Private Sub Form_Load()

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    
    Dim intDefect_Count             As Integer
    Dim intRecord_Count             As Integer
    Dim intRecord_Index             As Integer
    Dim intRow                      As Integer
    Dim intPortNo                   As Integer
    
    Call Init_Grid
    Call Init_Form
    Call Fill_Data
    
    m_CURRENT_DEFECT_INDEX = 1

    intDefect_Count = RANK_OBJ.Get_DEFECT_DATA_COUNT
    
'    If m_CURRENT_DEFECT_INDEX >= intDefect_Count Then
'        Me.cmdGrade.Enabled = True
'    Else
'        Me.cmdGrade.Enabled = False
'    End If
    
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Parameter.mdb"
    
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
        strQuery = "SELECT * FROM PATTERN_LIST"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        
        If lstRecord.EOF = False Then
            lstRecord.MoveLast
            intRecord_Count = lstRecord.RecordCount
            m_PATTERN_COUNT = intRecord_Count
            If intRecord_Count > 0 Then
                ReDim PATTERN_LIST(intRecord_Count)
                lstRecord.MoveFirst
                intRecord_Index = 0
                While lstRecord.EOF = False
                    intRecord_Index = intRecord_Index + 1
                    With PATTERN_LIST(intRecord_Index)
                        .PATTERN_CODE = lstRecord.Fields("PATTERN_CODE")
                        .PATTERN_NAME = lstRecord.Fields("PATTERN_NAME")
                        .DELAY_TIME = lstRecord.Fields("DELAY_TIME")
                        .LEVEL = lstRecord.Fields("LEVEL")
                        .DH = lstRecord.Fields("DH")
                        .DL = lstRecord.Fields("DL")
                        .VGH = lstRecord.Fields("VGH")
                        .VGL = lstRecord.Fields("VGL")
                        .VCOM = lstRecord.Fields("VCOM")
                    End With
                    lstRecord.MoveNext
                Wend
            End If
        End If
        lstRecord.Close
        
        dbMyDB.Close
    End If
    
    Me.flxPG_Data.Rows = 1
    If intRecord_Count > 0 Then
        For intRecord_Index = 1 To intRecord_Count
            intRow = Add_Grid("", Me.flxPG_Data)
            With Me.flxPG_Data
                .Row = intRow
                .Col = 0
'                If Dir(App.PATH & "\Env\Standard_Info\" & PATTERN_LIST(intRecord_Index).PATTERN_NAME & ".jpg", vbNormal) <> "" Then
'                    .CellPicture = LoadPicture(App.PATH & "\Env\Standard_Info\" & PATTERN_LIST(intRecord_Index).PATTERN_NAME & ".jpg")
'                End If
                .TextMatrix(intRow, 1) = PATTERN_LIST(intRecord_Index).PATTERN_CODE
                .TextMatrix(intRow, 2) = PATTERN_LIST(intRecord_Index).PATTERN_NAME
                .TextMatrix(intRow, 3) = PATTERN_LIST(intRecord_Index).DELAY_TIME
            End With
        Next intRecord_Index
        Me.lblCurrent_PTN_Index.Caption = "0"
    End If
    
    Me.Height = 11535
    Me.Width = 20370
    
    Me.KeyPreview = True
    
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

'    If Button = vbRightButton Then
'        Me.cmdGrade.SetFocus
'    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    frmMain.cmdJudge.Enabled = False
    
    Call SaveLog("frmJudge_Form_Unload", "Manual Judge window unload.")
    
End Sub

Private Sub Frame1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

'    If Me.cmdGrade.Enabled = True Then
'        If Button = vbRightButton Then
'            Me.cmdGrade.SetFocus
'        End If
'    End If
    
End Sub

Private Sub Frame2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

'    If Me.cmdGrade.Enabled = True Then
'        If Button = vbRightButton Then
'            frmJudge.cmdGrade.SetFocus
'        End If
'    End If

End Sub

Private Sub Frame3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

'    If Me.cmdGrade.Enabled = True Then
'        If Button = vbRightButton Then
'            frmJudge.cmdGrade.SetFocus
'        End If
'    End If

End Sub

Private Sub Frame4_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

'    If Me.cmdGrade.Enabled = True Then
'        If Button = vbRightButton Then
'            frmJudge.cmdGrade.SetFocus
'        End If
'    End If

End Sub

Private Sub Frame5_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

'    If Me.cmdGrade.Enabled = True Then
'        If Button = vbRightButton Then
'            frmJudge.cmdGrade.SetFocus
'        End If
'    End If

End Sub

Private Sub Frame6_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

'    If Me.cmdGrade.Enabled = True Then
'        If Button = vbRightButton Then
'            frmJudge.cmdGrade.SetFocus
'        End If
'    End If

End Sub

Private Sub Frame7_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

'    If Me.cmdGrade.Enabled = True Then
'        If Button = vbRightButton Then
'            frmJudge.cmdGrade.SetFocus
'        End If
'    End If

End Sub

Private Sub Frame8_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

'    If Me.cmdGrade.Enabled = True Then
'        If Button = vbRightButton Then
'            frmJudge.cmdGrade.SetFocus
'        End If
'    End If
    
End Sub

Private Sub Frame9_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

'    If Me.cmdGrade.Enabled = True Then
'        If Button = vbRightButton Then
'            frmJudge.cmdGrade.SetFocus
'        End If
'    End If

End Sub

Private Sub picCurrent_Pattern_Click()

    Dim intPortID                       As Integer
    Dim intPTN_Index                    As Integer
    
'    If m_CURRENT_PATTERN_INDEX = m_PATTERN_COUNT Then
'        m_CURRENT_PATTERN_INDEX = 1
'    Else
'        m_CURRENT_PATTERN_INDEX = m_CURRENT_PATTERN_INDEX + 1
'    End If
'
'    Select Case UCase(PATTERN_LIST(m_CURRENT_PATTERN_INDEX).PATTERN_CODE)
'    Case "R":
'        Me.picCurrent_Pattern.BackColor = vbGreen
'        'PG Interface
'    Case "G":
'        Me.picCurrent_Pattern.BackColor = vbBlue
'        'PG Interface
'    Case "B":
'        Me.picCurrent_Pattern.BackColor = vbBlack
'        'PG Interface
'    Case "BLACK:"
'        Me.picCurrent_Pattern.BackColor = vbRed
'    Case "WHITE":
'    End Select
'    Me.tmrPattern_Delay.Interval = PATTERN_LIST(m_CURRENT_PATTERN_INDEX).DELAY_TIME * 1000
'    Me.picCurrent_Pattern.Enabled = False
'    Me.tmrPattern_Delay.Enabled = True
    Me.picCurrent_Pattern.Enabled = False
    If CInt(Me.lblCurrent_PTN_Index.Caption) < Me.flxPG_Data.Rows - 1 Then
        intPTN_Index = CInt(Me.lblCurrent_PTN_Index.Caption)
        EQP.Set_PATTERN_END_by_Index (intPTN_Index)
        intPortID = EQP.Get_PG_PortID
        Call QUEUE.Put_Send_Command(intPortID, "QPCC")
    Else
        intPTN_Index = CInt(Me.lblCurrent_PTN_Index.Caption)
        EQP.Set_PATTERN_END_by_Index (intPTN_Index)
        Me.lblCurrent_PTN_Index.Caption = "0"
        intPortID = EQP.Get_PG_PortID
        Call QUEUE.Put_Send_Command(intPortID, "QPCC")
        Me.cmdGrade.Enabled = True
    End If
    
End Sub

Private Sub tmrPattern_Delay_Timer()

    Me.tmrPattern_Delay.Enabled = False
    Me.picCurrent_Pattern.Enabled = True
    
End Sub
'Lucas Ver.0.9.17 2012.03.29========================For QJPG delay sending
Public Sub delaytime(TTime As Long)
On Error Resume Next
Dim Tstart As Single
Tstart = Timer
While (Timer - Tstart) < TTime
DoEvents
Wend
End Sub
'Lucas Ver.0.9.17 2012.03.29========================For QJPG delay sending

Private Sub Init_Form()

    Dim dbMyDB                          As Database
    
    Dim lstRecord                       As Recordset
    
    Dim typRANK_DATA                    As RANK_DATA_STRUCTURE
    
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
        
        strQuery = "SELECT * FROM DEFECT_LIST WHERE "
        strQuery = strQuery & "DEFECT_KIND = 'D'"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
    
        If lstRecord.EOF = False Then
            lstRecord.MoveFirst
            While lstRecord.EOF = False
                intRow = Add_Grid(lstRecord.Fields("DEFECT_CODE"), Me.flxDefect_A)
                Me.flxDefect_A.TextMatrix(intRow, 1) = lstRecord.Fields("DEFECT_NAME")
                typRANK_DATA = Get_DEFECT_DATA_by_CODE(lstRecord.Fields("DEFECT_CODE"))
                Me.flxDefect_A.Row = intRow
                Me.flxDefect_A.Col = 0
                If typRANK_DATA.DEFECT_TYPE = "A" Then
                    Me.flxDefect_A.CellForeColor = vbBlue
                Else
                    Me.flxDefect_A.CellForeColor = vbBlack
                End If
                lstRecord.MoveNext
            Wend
        End If
        lstRecord.Close
        
        strQuery = "SELECT * FROM DEFECT_LIST WHERE "
        strQuery = strQuery & "DEFECT_KIND = 'L'"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
    
        If lstRecord.EOF = False Then
            lstRecord.MoveFirst
            While lstRecord.EOF = False
                intRow = Add_Grid(lstRecord.Fields("DEFECT_CODE"), Me.flxDefect_B)
                Me.flxDefect_B.TextMatrix(intRow, 1) = lstRecord.Fields("DEFECT_NAME")
                typRANK_DATA = Get_DEFECT_DATA_by_CODE(lstRecord.Fields("DEFECT_CODE"))
                Me.flxDefect_B.Row = intRow
                Me.flxDefect_B.Col = 0
                If typRANK_DATA.DEFECT_TYPE = "A" Then
                    Me.flxDefect_B.CellForeColor = vbBlue
                Else
                    Me.flxDefect_B.CellForeColor = vbBlack
                End If
                lstRecord.MoveNext
            Wend
        End If
        lstRecord.Close
    
        strQuery = "SELECT * FROM DEFECT_LIST WHERE "
        strQuery = strQuery & "DEFECT_KIND = 'G'"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
    
        If lstRecord.EOF = False Then
            lstRecord.MoveFirst
            While lstRecord.EOF = False
                intRow = Add_Grid(lstRecord.Fields("DEFECT_CODE"), Me.flxDefect_C)
                Me.flxDefect_C.TextMatrix(intRow, 1) = lstRecord.Fields("DEFECT_NAME")
                typRANK_DATA = Get_DEFECT_DATA_by_CODE(lstRecord.Fields("DEFECT_CODE"))
                Me.flxDefect_C.Row = intRow
                Me.flxDefect_C.Col = 0
                If typRANK_DATA.DEFECT_TYPE = "A" Then
                    Me.flxDefect_C.CellForeColor = vbBlue
                Else
                    Me.flxDefect_C.CellForeColor = vbBlack
                End If
                lstRecord.MoveNext
            Wend
        End If
        lstRecord.Close
    
        strQuery = "SELECT * FROM DEFECT_LIST WHERE "
        strQuery = strQuery & "DEFECT_KIND = 'M'"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
    
        If lstRecord.EOF = False Then
            lstRecord.MoveFirst
            While lstRecord.EOF = False
                intRow = Add_Grid(lstRecord.Fields("DEFECT_CODE"), Me.flxDefect_D)
                Me.flxDefect_D.TextMatrix(intRow, 1) = lstRecord.Fields("DEFECT_NAME")
                typRANK_DATA = Get_DEFECT_DATA_by_CODE(lstRecord.Fields("DEFECT_CODE"))
                Me.flxDefect_D.Row = intRow
                Me.flxDefect_D.Col = 0
                If typRANK_DATA.DEFECT_TYPE = "A" Then
                    Me.flxDefect_D.CellForeColor = vbBlue
                Else
                    Me.flxDefect_D.CellForeColor = vbBlack
                End If
                lstRecord.MoveNext
            Wend
        End If
        lstRecord.Close
    
        strQuery = "SELECT * FROM DEFECT_LIST WHERE "
        strQuery = strQuery & "DEFECT_KIND = 'F'"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
    
        If lstRecord.EOF = False Then
            lstRecord.MoveFirst
            While lstRecord.EOF = False
                intRow = Add_Grid(lstRecord.Fields("DEFECT_CODE"), Me.flxDefect_E)
                Me.flxDefect_E.TextMatrix(intRow, 1) = lstRecord.Fields("DEFECT_NAME")
                typRANK_DATA = Get_DEFECT_DATA_by_CODE(lstRecord.Fields("DEFECT_CODE"))
                Me.flxDefect_E.Row = intRow
                Me.flxDefect_E.Col = 0
                If typRANK_DATA.DEFECT_TYPE = "A" Then
                    Me.flxDefect_E.CellForeColor = vbBlue
                Else
                    Me.flxDefect_E.CellForeColor = vbBlack
                End If
                lstRecord.MoveNext
            Wend
        End If
        lstRecord.Close
    
        strQuery = "SELECT * FROM DEFECT_LIST WHERE "
        strQuery = strQuery & "DEFECT_KIND = 'P'"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
    
        If lstRecord.EOF = False Then
            lstRecord.MoveFirst
            While lstRecord.EOF = False
                intRow = Add_Grid(lstRecord.Fields("DEFECT_CODE"), Me.flxDefect_F)
                Me.flxDefect_F.TextMatrix(intRow, 1) = lstRecord.Fields("DEFECT_NAME")
                typRANK_DATA = Get_DEFECT_DATA_by_CODE(lstRecord.Fields("DEFECT_CODE"))
                Me.flxDefect_F.Row = intRow
                Me.flxDefect_F.Col = 0
                If typRANK_DATA.DEFECT_TYPE = "A" Then
                    Me.flxDefect_F.CellForeColor = vbBlue
                Else
                    Me.flxDefect_F.CellForeColor = vbBlack
                End If
                lstRecord.MoveNext
            Wend
        End If
        lstRecord.Close
        
        strQuery = "SELECT * FROM DEFECT_LIST WHERE "
        strQuery = strQuery & "DEFECT_KIND = 'A'"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
    
        If lstRecord.EOF = False Then
            lstRecord.MoveFirst
            While lstRecord.EOF = False
                intRow = Add_Grid(lstRecord.Fields("DEFECT_CODE"), Me.flxDefect_G)
                Me.flxDefect_G.TextMatrix(intRow, 1) = lstRecord.Fields("DEFECT_NAME")
                typRANK_DATA = Get_DEFECT_DATA_by_CODE(lstRecord.Fields("DEFECT_CODE"))
                Me.flxDefect_G.Row = intRow
                Me.flxDefect_G.Col = 0
                If typRANK_DATA.DEFECT_TYPE = "A" Then
                    Me.flxDefect_G.CellForeColor = vbBlue
                Else
                    Me.flxDefect_G.CellForeColor = vbBlack
                End If
                lstRecord.MoveNext
            Wend
        End If
        lstRecord.Close
        
        strQuery = "SELECT * FROM DEFECT_LIST WHERE "
        strQuery = strQuery & "DEFECT_KIND = 'C'"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
    
        If lstRecord.EOF = False Then
            lstRecord.MoveFirst
            While lstRecord.EOF = False
                intRow = Add_Grid(lstRecord.Fields("DEFECT_CODE"), Me.flxDefect_H)
                Me.flxDefect_H.TextMatrix(intRow, 1) = lstRecord.Fields("DEFECT_NAME")
                typRANK_DATA = Get_DEFECT_DATA_by_CODE(lstRecord.Fields("DEFECT_CODE"))
                Me.flxDefect_H.Row = intRow
                Me.flxDefect_H.Col = 0
                If typRANK_DATA.DEFECT_TYPE = "A" Then
                    Me.flxDefect_H.CellForeColor = vbBlue
                Else
                    Me.flxDefect_H.CellForeColor = vbBlack
                End If
                lstRecord.MoveNext
            Wend
        End If
        lstRecord.Close
        
        strQuery = "SELECT * FROM DEFECT_LIST WHERE "
        strQuery = strQuery & "DEFECT_KIND = 'O'"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
    
        If lstRecord.EOF = False Then
            lstRecord.MoveFirst
            While lstRecord.EOF = False
                intRow = Add_Grid(lstRecord.Fields("DEFECT_CODE"), Me.flxDefect_I)
                Me.flxDefect_I.TextMatrix(intRow, 1) = lstRecord.Fields("DEFECT_NAME")
                typRANK_DATA = Get_DEFECT_DATA_by_CODE(lstRecord.Fields("DEFECT_CODE"))
                Me.flxDefect_I.Row = intRow
                Me.flxDefect_I.Col = 0
                If typRANK_DATA.DEFECT_TYPE = "A" Then
                    Me.flxDefect_I.CellForeColor = vbBlue
                Else
                    Me.flxDefect_I.CellForeColor = vbBlack
                End If
                lstRecord.MoveNext
            Wend
        End If
        lstRecord.Close
        
        Me.cmbUseful_Defect.Clear
        
        strQuery = "SELECT * FROM USEFUL_DEFECT"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        
        If lstRecord.EOF = False Then
            lstRecord.MoveFirst
            While lstRecord.EOF = False
                Me.cmbUseful_Defect.AddItem lstRecord.Fields("DEFECT_CODE") & " " & lstRecord.Fields("DEFECT_NAME")
                lstRecord.MoveNext
            Wend
        End If
        lstRecord.Close
        If Me.cmbUseful_Defect.ListCount > 0 Then
            Me.cmbUseful_Defect.Text = Me.cmbUseful_Defect.List(0)
        End If
        
        dbMyDB.Close
    End If
    
    Me.imgPG_Image.Left = 0
    Me.imgPG_Image.Top = 0
    Me.imgPG_Image.Width = Me.picCurrent_Pattern.Width
    Me.imgPG_Image.Height = Me.picCurrent_Pattern.Height
    
    Exit Sub
    
ErrorHandler:

    ErrMsg = Err.Number & " - " & Err.Description
    
    Call SaveLog("frmJudge_Init_Form", ErrMsg)

End Sub

Private Sub Init_Grid()

    Dim intRow              As Integer
    Dim intCol              As Integer
    
    With Me.flxDefect_A
        For intRow = 0 To .Rows - 1
            For intCol = 0 To .Cols - 1
                .Row = intRow
                .Col = intCol
                .CellAlignment = flexAlignCenterCenter
                .RowHeight(intRow) = 350
            Next intCol
        Next intRow
        
        .ColWidth(0) = 1200
        .TextMatrix(0, 0) = "D/F CODE"
        .ColWidth(1) = 1800
        .TextMatrix(0, 1) = "DEFECT NAME"
    End With

    With Me.flxDefect_B
        For intRow = 0 To .Rows - 1
            For intCol = 0 To .Cols - 1
                .Row = intRow
                .Col = intCol
                .CellAlignment = flexAlignCenterCenter
                .RowHeight(intRow) = 350
            Next intCol
        Next intRow
        
        .ColWidth(0) = 1200
        .TextMatrix(0, 0) = "D/F CODE"
        .ColWidth(1) = 1800
        .TextMatrix(0, 1) = "DEFECT NAME"
    End With

    With Me.flxDefect_C
        For intRow = 0 To .Rows - 1
            For intCol = 0 To .Cols - 1
                .Row = intRow
                .Col = intCol
                .CellAlignment = flexAlignCenterCenter
                .RowHeight(intRow) = 350
            Next intCol
        Next intRow
        
        .ColWidth(0) = 1200
        .TextMatrix(0, 0) = "D/F CODE"
        .ColWidth(1) = 1800
        .TextMatrix(0, 1) = "DEFECT NAME"
    End With

    With Me.flxDefect_D
        For intRow = 0 To .Rows - 1
            For intCol = 0 To .Cols - 1
                .Row = intRow
                .Col = intCol
                .CellAlignment = flexAlignCenterCenter
                .RowHeight(intRow) = 350
            Next intCol
        Next intRow
        
        .ColWidth(0) = 1200
        .TextMatrix(0, 0) = "D/F CODE"
        .ColWidth(1) = 1800
        .TextMatrix(0, 1) = "DEFECT NAME"
    End With

    With Me.flxDefect_E
        For intRow = 0 To .Rows - 1
            For intCol = 0 To .Cols - 1
                .Row = intRow
                .Col = intCol
                .CellAlignment = flexAlignCenterCenter
                .RowHeight(intRow) = 350
            Next intCol
        Next intRow
        
        .ColWidth(0) = 1200
        .TextMatrix(0, 0) = "D/F CODE"
        .ColWidth(1) = 1800
        .TextMatrix(0, 1) = "DEFECT NAME"
    End With

    With Me.flxDefect_F
        For intRow = 0 To .Rows - 1
            For intCol = 0 To .Cols - 1
                .Row = intRow
                .Col = intCol
                .CellAlignment = flexAlignCenterCenter
                .RowHeight(intRow) = 350
            Next intCol
        Next intRow
        
        .ColWidth(0) = 1200
        .TextMatrix(0, 0) = "D/F CODE"
        .ColWidth(1) = 1800
        .TextMatrix(0, 1) = "DEFECT NAME"
    End With

    With Me.flxDefect_G
        For intRow = 0 To .Rows - 1
            For intCol = 0 To .Cols - 1
                .Row = intRow
                .Col = intCol
                .CellAlignment = flexAlignCenterCenter
                .RowHeight(intRow) = 350
            Next intCol
        Next intRow
        
        .ColWidth(0) = 1200
        .TextMatrix(0, 0) = "D/F CODE"
        .ColWidth(1) = 1800
        .TextMatrix(0, 1) = "DEFECT NAME"
    End With

    With Me.flxDefect_H
        For intRow = 0 To .Rows - 1
            For intCol = 0 To .Cols - 1
                .Row = intRow
                .Col = intCol
                .CellAlignment = flexAlignCenterCenter
                .RowHeight(intRow) = 350
            Next intCol
        Next intRow
        
        .ColWidth(0) = 1200
        .TextMatrix(0, 0) = "D/F CODE"
        .ColWidth(1) = 1800
        .TextMatrix(0, 1) = "DEFECT NAME"
    End With

    With Me.flxDefect_I
        For intRow = 0 To .Rows - 1
            For intCol = 0 To .Cols - 1
                .Row = intRow
                .Col = intCol
                .CellAlignment = flexAlignCenterCenter
                .RowHeight(intRow) = 350
            Next intCol
        Next intRow
        
        .ColWidth(0) = 1200
        .TextMatrix(0, 0) = "D/F CODE"
        .ColWidth(1) = 1800
        .TextMatrix(0, 1) = "DEFECT NAME"
    End With

    With Me.flxDefect_List
        For intRow = 0 To .Rows - 1
            For intCol = 0 To .Cols - 1
                .Row = intRow
                .Col = intCol
                .CellAlignment = flexAlignCenterCenter
                .RowHeight(intRow) = 350
            Next intCol
        Next intRow
        
        .ColWidth(0) = 1200
        .TextMatrix(0, 0) = "D/F CODE"
        .ColWidth(1) = 1800
        .TextMatrix(0, 1) = "DEFECT NAME"
        .ColWidth(2) = 1200
        .TextMatrix(0, 2) = "DATA1"
        .ColWidth(3) = 1200
        .TextMatrix(0, 3) = "GATE1"
        .ColWidth(4) = 1200
        .TextMatrix(0, 4) = "DATA2"
        .ColWidth(5) = 1200
        .TextMatrix(0, 5) = "GATE2"
        .ColWidth(6) = 1200
        .TextMatrix(0, 6) = "DATA3"
        .ColWidth(7) = 1200
        .TextMatrix(0, 7) = "GATE3"
        .ColWidth(8) = 1800
        .TextMatrix(0, 8) = "PANELID"
        .ColWidth(9) = 600
        .TextMatrix(0, 9) = "RANK"
        .ColWidth(10) = 1000
        .TextMatrix(0, 10) = "VALUE"
        .ColWidth(11) = 1500
        .TextMatrix(0, 11) = "D/F DETAIL"
        .ColWidth(12) = 1200
        .TextMatrix(0, 12) = "COLOR"
        .ColWidth(13) = 1200
        .TextMatrix(0, 13) = "GRAY L/V"
        .ColWidth(14) = 800
        .TextMatrix(0, 14) = "GRADE"
    End With

    With Me.flxPG_Data
        For intRow = 0 To .Rows - 1
            For intCol = 0 To .Cols - 1
                .Row = intRow
                .Col = intCol
                .CellAlignment = flexAlignCenterCenter
                .RowHeight(intRow) = 350
            Next intCol
        Next intRow
        
        .ColWidth(0) = 600
        .TextMatrix(0, 0) = ""
        .ColWidth(1) = 1000
        .TextMatrix(0, 1) = "CODE"
        .ColWidth(2) = 1000
        .TextMatrix(0, 2) = "NAME"
        .ColWidth(3) = 1000
        .TextMatrix(0, 3) = "DELAY"
    End With
    
End Sub

Private Sub Fill_Data()

    Dim typDEFECT_DATA              As DEFECT_DATA_STRUCTURE
    
    Dim intIndex                    As Integer
    
    With typDEFECT_DATA
        If RANK_OBJ.Get_DEFECT_DATA_by_Index(1, .PANELID, .DEFECT_CODE, .DEFECT_NAME, .DETAIL_DIVISION, .DATA_ADDRESS, .GATE_ADDRESS, _
                                             .GRADE, .Rank, .COLOR, .GRAY_LEVEL, .ACCUMULATION) = True Then
            If Mid(.DEFECT_CODE, 2, 1) = "M" Then
                For intIndex = 1 To 3
                    Me.txtX_Data(intIndex - 1).Text = .DATA_ADDRESS(intIndex)
                    Me.txtY_Gate(intIndex - 1).Text = .GATE_ADDRESS(intIndex)
                Next intIndex
            Else
                Me.txtX_Data(0).Text = .DATA_ADDRESS(1)
                Me.txtY_Gate(0).Text = .GATE_ADDRESS(1)
            End If
        Else
            Call SaveLog("frmJudge_Fill_Data", "DEFECT DATA loading fail. Index : 1")
        End If
    End With
    
End Sub

Private Function Add_Grid(ByVal pDEFECT_CODE As String, pGRID As Object) As Integer

    Dim objFlexGrid     As Object
    
    Dim intRow          As Integer
    Dim intCol          As Integer
    
    With pGRID
        intRow = .Rows
        .AddItem pDEFECT_CODE
        Call RANK_OBJ.Set_Select_DEFECTCODE(pDEFECT_CODE)
        .RowHeight(intRow) = 350
        For intCol = 0 To .Cols - 1
            .Row = intRow
            .Col = intCol
            .CellAlignment = flexAlignCenterCenter
        Next intCol
        
        Add_Grid = .Rows - 1
    End With

End Function

Private Sub Set_Interlock()

    Dim intIndex                        As Integer
    
    With Me
        .flxDefect_A.Enabled = False
        .flxDefect_B.Enabled = False
        .flxDefect_C.Enabled = False
        .flxDefect_D.Enabled = False
        .flxDefect_E.Enabled = False
        .flxDefect_F.Enabled = False
        .flxDefect_G.Enabled = False
        .flxDefect_H.Enabled = False
        .flxDefect_I.Enabled = False
        .cmdGrade.Enabled = True        '2012. 04. 24
    End With
    
    For intIndex = 0 To 2
        Me.txtX_Data(intIndex).Text = ""
        Me.txtY_Gate(intIndex).Text = ""
    Next intIndex
    
End Sub

Public Function Get_Current_Defect_Index() As Integer

    Get_Current_Defect_Index = m_CURRENT_DEFECT_INDEX
    
End Function

Public Function Get_Defect_Type(ByVal pDEFECT_CODE As String) As String

    Dim typRANK_DATA                As RANK_DATA_STRUCTURE
    Dim typGRADE_DATA()             As GRADE_DATA_STRUCTURE
    
    Dim intGrade_Count              As Integer
    
    Call Get_Rank_Data(pubCST_INFO.PROCESS_NUM, typRANK_DATA, typGRADE_DATA, pDEFECT_CODE, intGrade_Count)

    Get_Defect_Type = typRANK_DATA.DEFECT_TYPE
    
End Function

Public Sub Load_Manual_Judge()

    Dim typRANK_DATA                    As RANK_DATA_STRUCTURE
    Dim typGRADE_DATA()                 As GRADE_DATA_STRUCTURE
    Dim typPATTERN_LIST                 As PATTERN_LIST_DATA
    
    Dim strDATA_ADDRESS(1 To 3)         As String
    Dim strGATE_ADDRESS(1 To 3)         As String
        
    Dim strDefect_Code                  As String
    
    Dim intDefect_Count                 As Integer
    Dim intIndex                        As Integer
    Dim intCol                          As Integer
    Dim intRow                          As Integer
    Dim intGrade_Count                  As Integer
        '============Leo 2012.05.22 Add Rank Level Start
    Dim intRankLevel                 As Integer
    '============Leo 2012.05.22 Add Rank Level end

    intRow = frmJudge.flxDefect_List.Rows - 1
    strDefect_Code = frmJudge.flxDefect_List.TextMatrix(intRow, 0)
   
    
    For intIndex = 1 To 3
        strDATA_ADDRESS(intIndex) = Space(5)
        strGATE_ADDRESS(intIndex) = Space(5)
        frmJudge.txtX_Data(intIndex - 1).Text = strDATA_ADDRESS(intIndex)
        frmJudge.txtY_Gate(intIndex - 1).Text = strGATE_ADDRESS(intIndex)
    Next intIndex
    
    intIndex = 0
    For intCol = 2 To 7 Step 2
        intIndex = intIndex + 1
        With frmJudge.flxDefect_List
            .TextMatrix(intRow, intCol) = strDATA_ADDRESS(intIndex)
            .TextMatrix(intRow, intCol + 1) = strGATE_ADDRESS(intIndex)
        End With
    Next intCol
    
    'Check Defect Type
    Call Get_Rank_Data(pubCST_INFO.PROCESS_NUM, typRANK_DATA, typGRADE_DATA, strDefect_Code, intGrade_Count)

    Load frmManual_Judge
    
    With frmManual_Judge
         '============Leo 2012.05.22 Add Rank Level Start
                For intRankLevel = 0 To UBound(RankLevel)
                    If (Trim(typRANK_DATA.Rank(intRankLevel)) <> "0") And (Trim(typRANK_DATA.Rank(intRankLevel)) <> "-") Then
                        .lblGrade(intRankLevel).Caption = RankLevel(intRankLevel)
                        .optSpec_Value(intRankLevel).Caption = typRANK_DATA.Rank(intRankLevel)
                        .lblGrade(intRankLevel).Visible = True
                        .optSpec_Value(intRankLevel).Visible = True
                    End If
                Next intRankLevel
                                        
'                If (Trim(typRANK_DATA.RANK_Y) <> "0") And (Trim(typRANK_DATA.RANK_Y) <> "-") Then
'                    .lblGrade(0).Caption = "Y"
'                    .optSpec_Value(0).Caption = typRANK_DATA.RANK_Y
'                    .lblGrade(0).Visible = True
'                    .optSpec_Value(0).Visible = True
'                End If
'
'                If (Trim(typRANK_DATA.RANK_L) <> "0") And (Trim(typRANK_DATA.RANK_L) <> "-") Then
'                    .lblGrade(1).Caption = "L"
'                    .optSpec_Value(1).Caption = typRANK_DATA.RANK_L
'                    .lblGrade(1).Visible = True
'                    .optSpec_Value(1).Visible = True
'                End If
'
'                If (Trim(typRANK_DATA.RANK_K) <> "0") And (Trim(typRANK_DATA.RANK_K) <> "-") Then
'                    .lblGrade(2).Caption = "K"
'                    .optSpec_Value(2).Caption = typRANK_DATA.RANK_K
'                    .lblGrade(2).Visible = True
'                    .optSpec_Value(2).Visible = True
'                End If
'
'                If (Trim(typRANK_DATA.RANK_C) <> "0") And (Trim(typRANK_DATA.RANK_C) <> "-") Then
'                    .lblGrade(3).Caption = "C"
'                    .optSpec_Value(3).Caption = typRANK_DATA.RANK_C
'                    .lblGrade(3).Visible = True
'                    .optSpec_Value(3).Visible = True
'                End If
'
'                If (Trim(typRANK_DATA.RANK_S) <> "0") And (Trim(typRANK_DATA.RANK_S) <> "-") Then
'                    .lblGrade(4).Caption = "S"
'                    .optSpec_Value(4).Caption = typRANK_DATA.RANK_S
'                    .lblGrade(4).Visible = True
'                    .optSpec_Value(4).Visible = True
'                End If
 '============Leo 2012.05.22 Add Rank Level End
        .lblDefect_Code.Caption = strDefect_Code
        .lblDefect_Name.Text = frmJudge.flxDefect_List.TextMatrix(intRow, 1)
        .lstData_Address.Clear
        .lstGate_Address.Clear
        
        For intIndex = 1 To 3
            .lstData_Address.AddItem strDATA_ADDRESS(intIndex)
            .lstGate_Address.AddItem strGATE_ADDRESS(intIndex)
        Next intIndex
    End With
    
    frmManual_Judge.Show
    
    frmJudge.flxDefect_List.TextMatrix(frmJudge.flxDefect_List.Rows - 1, 11) = typRANK_DATA.DETAIL_DIVISION
    intIndex = CInt(frmJudge.lblCurrent_PTN_Index.Caption)
    
    If intIndex > 0 Then
        With typPATTERN_LIST
            Call EQP.Get_PATTERN_LIST_by_Index(intIndex, .PATTERN_CODE, .PATTERN_NAME, .DELAY_TIME, .LEVEL, .DH, .DL, .VGH, .VGL, .RESCUE_HIGH, .RESCUE_LOW, .VCOM)
        frmJudge.flxDefect_List.TextMatrix(frmJudge.flxDefect_List.Rows - 1, 12) = .PATTERN_NAME
        frmJudge.flxDefect_List.TextMatrix(frmJudge.flxDefect_List.Rows - 1, 13) = .LEVEL
        End With
    End If

End Sub
'2012. 04. 24
Private Sub txtX_Data_Change(Index As Integer)

    If (Me.txtX_Data(0).Text <> "") And (Me.txtY_Gate(0).Text <> "") Then
        If (IsNumeric(Me.txtX_Data(0).Text) = True) And (IsNumeric(Me.txtY_Gate(0).Text) = True) Then
            Me.cmdGrade.Enabled = True
        End If
    End If

End Sub
'
'Private Sub txtX_Data_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'
'    If Me.picCurrent_Pattern.Enabled = True Then
'        If KeyCode = vbKeySpace Then
'            Call picCurrent_Pattern_Click
'            Me.txtX_Data(Index).Text = Trim(Me.txtX_Data(Index).Text)
'        Else
'            Me.txtX_Data(Index).Text = Me.txtX_Data(Index).Text & Chr(KeyCode)
'            SendKeys "{End}"
'        End If
'    End If
'
'End Sub

'Private Sub txtY_Gate_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'
'    If Me.picCurrent_Pattern.Enabled = True Then
'        If KeyCode = vbKeySpace Then
'            Call picCurrent_Pattern_Click
'            Me.txtY_Gate(Index).Text = Trim(Me.txtY_Gate(Index).Text)
'        Else
'            Me.txtX_Data(Index).Text = Me.txtX_Data(Index).Text & Chr(KeyCode)
'            SendKeys "{End}"
'        End If
'    End If
'
'End Sub
'2012. 04. 24
Private Sub txtY_Gate_Change(Index As Integer)

    If (Me.txtX_Data(0).Text <> "") And (Me.txtY_Gate(0).Text <> "") Then
        If (IsNumeric(Me.txtX_Data(0).Text) = True) And (IsNumeric(Me.txtY_Gate(0).Text) = True) Then
            Me.cmdGrade.Enabled = True
        End If
    End If
    
End Sub
