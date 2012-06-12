VERSION 5.00
Begin VB.Form frmManual_Judge 
   Caption         =   "Manual Judge"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16815
   LinkTopic       =   "Form1"
   ScaleHeight     =   7275
   ScaleWidth      =   16815
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Frame Frame4 
      Height          =   1875
      Left            =   0
      TabIndex        =   30
      Top             =   5400
      Width           =   16815
      Begin VB.OptionButton optSpec_Value 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Index           =   11
         Left            =   13200
         TabIndex        =   38
         Top             =   840
         Width           =   2835
      End
      Begin VB.OptionButton optSpec_Value 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Index           =   10
         Left            =   9570
         TabIndex        =   36
         Top             =   840
         Width           =   2835
      End
      Begin VB.OptionButton optSpec_Value 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Index           =   9
         Left            =   6000
         TabIndex        =   34
         Top             =   840
         Width           =   2835
      End
      Begin VB.OptionButton optSpec_Value 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Index           =   8
         Left            =   2910
         TabIndex        =   32
         Top             =   840
         Width           =   2835
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "SPEC."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   330
         TabIndex        =   40
         Top             =   1140
         Width           =   1080
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "GRADE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   300
         TabIndex        =   39
         Top             =   510
         Width           =   1200
      End
      Begin VB.Label lblGrade 
         AutoSize        =   -1  'True
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   11
         Left            =   13200
         TabIndex        =   37
         Top             =   510
         Width           =   240
      End
      Begin VB.Label lblGrade 
         AutoSize        =   -1  'True
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   10
         Left            =   9570
         TabIndex        =   35
         Top             =   510
         Width           =   240
      End
      Begin VB.Label lblGrade 
         AutoSize        =   -1  'True
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   6000
         TabIndex        =   33
         Top             =   510
         Width           =   240
      End
      Begin VB.Label lblGrade 
         AutoSize        =   -1  'True
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   2910
         TabIndex        =   31
         Top             =   510
         Width           =   240
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1875
      Left            =   0
      TabIndex        =   19
      Top             =   3510
      Width           =   16815
      Begin VB.OptionButton optSpec_Value 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Index           =   7
         Left            =   13200
         TabIndex        =   27
         Top             =   840
         Width           =   2835
      End
      Begin VB.OptionButton optSpec_Value 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Index           =   6
         Left            =   9570
         TabIndex        =   25
         Top             =   840
         Width           =   2835
      End
      Begin VB.OptionButton optSpec_Value 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Index           =   5
         Left            =   6000
         TabIndex        =   23
         Top             =   840
         Width           =   2835
      End
      Begin VB.OptionButton optSpec_Value 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Index           =   4
         Left            =   2910
         TabIndex        =   21
         Top             =   840
         Width           =   2835
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "SPEC."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   330
         TabIndex        =   29
         Top             =   1140
         Width           =   1080
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "GRADE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   300
         TabIndex        =   28
         Top             =   510
         Width           =   1200
      End
      Begin VB.Label lblGrade 
         AutoSize        =   -1  'True
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   13200
         TabIndex        =   26
         Top             =   510
         Width           =   240
      End
      Begin VB.Label lblGrade 
         AutoSize        =   -1  'True
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   9570
         TabIndex        =   24
         Top             =   510
         Width           =   240
      End
      Begin VB.Label lblGrade 
         AutoSize        =   -1  'True
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   6000
         TabIndex        =   22
         Top             =   510
         Width           =   240
      End
      Begin VB.Label lblGrade 
         AutoSize        =   -1  'True
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   2910
         TabIndex        =   20
         Top             =   510
         Width           =   240
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1875
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   1620
      Width           =   16815
      Begin VB.OptionButton optSpec_Value 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Index           =   3
         Left            =   13200
         TabIndex        =   17
         Top             =   840
         Width           =   2835
      End
      Begin VB.OptionButton optSpec_Value 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Index           =   2
         Left            =   9570
         TabIndex        =   16
         Top             =   840
         Width           =   2835
      End
      Begin VB.OptionButton optSpec_Value 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Index           =   1
         Left            =   6000
         TabIndex        =   15
         Top             =   840
         Width           =   2835
      End
      Begin VB.OptionButton optSpec_Value 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Index           =   0
         Left            =   2910
         TabIndex        =   14
         Top             =   840
         Width           =   2835
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "SPEC."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   330
         TabIndex        =   18
         Top             =   1110
         Width           =   1080
      End
      Begin VB.Label lblGrade 
         AutoSize        =   -1  'True
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   13170
         TabIndex        =   13
         Top             =   540
         Width           =   240
      End
      Begin VB.Label lblGrade 
         AutoSize        =   -1  'True
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   9570
         TabIndex        =   12
         Top             =   510
         Width           =   240
      End
      Begin VB.Label lblGrade 
         AutoSize        =   -1  'True
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   6000
         TabIndex        =   11
         Top             =   510
         Width           =   240
      End
      Begin VB.Label lblGrade 
         AutoSize        =   -1  'True
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   2910
         TabIndex        =   10
         Top             =   510
         Width           =   240
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "GRADE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   300
         TabIndex        =   9
         Top             =   510
         Width           =   1200
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1605
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16815
      Begin VB.TextBox lblDefect_Name 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   8460
         TabIndex        =   41
         Top             =   450
         Width           =   2895
      End
      Begin VB.ListBox lstGate_Address 
         Height          =   600
         ItemData        =   "frmManual_Judge.frx":0000
         Left            =   14250
         List            =   "frmManual_Judge.frx":0002
         TabIndex        =   7
         Top             =   690
         Width           =   2385
      End
      Begin VB.ListBox lstData_Address 
         Height          =   600
         ItemData        =   "frmManual_Judge.frx":0004
         Left            =   11670
         List            =   "frmManual_Judge.frx":0006
         TabIndex        =   6
         Top             =   690
         Width           =   2385
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "GATE ADDRESS"
         Height          =   180
         Left            =   14250
         TabIndex        =   5
         Top             =   450
         Width           =   1395
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "DATA ADDRESS"
         Height          =   180
         Left            =   11670
         TabIndex        =   4
         Top             =   450
         Width           =   1380
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "DEFECT NAME"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5910
         TabIndex        =   3
         Top             =   510
         Width           =   2475
      End
      Begin VB.Label lblDefect_Code 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2730
         TabIndex        =   2
         Top             =   450
         Width           =   2895
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DEFECT CODE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   1
         Top             =   510
         Width           =   2475
      End
   End
End
Attribute VB_Name = "frmManual_Judge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Dim intIndex                 As Integer
    
    For intIndex = 0 To 11
        Me.lblGrade(intIndex).Visible = False
        Me.optSpec_Value(intIndex).Visible = False
    Next intIndex
    
End Sub

Private Sub optSpec_Value_Click(Index As Integer)

    Dim intRow                  As Integer
    
    intRow = frmJudge.flxDefect_List.Rows - 1
    
    With frmJudge.flxDefect_List
        .TextMatrix(intRow, 9) = Me.lblGrade(Index).Caption
        .TextMatrix(intRow, 10) = Me.optSpec_Value(Index).Caption
    End With
    
    Unload Me
    
End Sub
