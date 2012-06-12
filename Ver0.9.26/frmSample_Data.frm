VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSample_Data 
   Caption         =   "MES Data Input"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16860
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   16860
   StartUpPosition =   2  '화면 가운데
   Begin VB.Frame Frame4 
      Height          =   7305
      Left            =   12630
      TabIndex        =   22
      Top             =   0
      Width           =   4185
      Begin VB.TextBox txtShare_Value 
         Alignment       =   2  '가운데 맞춤
         Height          =   285
         Left            =   870
         TabIndex        =   28
         Top             =   6630
         Width           =   1755
      End
      Begin MSFlexGridLib.MSFlexGrid flxShare_Data 
         Height          =   5475
         Left            =   180
         TabIndex        =   24
         Top             =   540
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   9657
         _Version        =   393216
      End
      Begin VB.Label lblShare_Title 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  '단일 고정
         Height          =   270
         Left            =   870
         TabIndex        =   27
         Top             =   6210
         Width           =   1775
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "DATA"
         Height          =   180
         Left            =   240
         TabIndex        =   26
         Top             =   6660
         Width           =   480
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "NAME"
         Height          =   180
         Left            =   240
         TabIndex        =   25
         Top             =   6240
         Width           =   540
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "SHARE_DATA"
         Height          =   180
         Left            =   270
         TabIndex        =   23
         Top             =   270
         Width           =   1170
      End
   End
   Begin VB.Frame Frame3 
      Height          =   7305
      Left            =   8430
      TabIndex        =   15
      Top             =   0
      Width           =   4185
      Begin VB.TextBox txtJOB_Value 
         Alignment       =   2  '가운데 맞춤
         Height          =   285
         Left            =   870
         TabIndex        =   21
         Top             =   6630
         Width           =   1755
      End
      Begin MSFlexGridLib.MSFlexGrid flxJob_Data 
         Height          =   5475
         Left            =   180
         TabIndex        =   17
         Top             =   540
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   9657
         _Version        =   393216
      End
      Begin VB.Label lblJOB_Title 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  '단일 고정
         Height          =   270
         Left            =   870
         TabIndex        =   20
         Top             =   6210
         Width           =   1775
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "DATA"
         Height          =   180
         Left            =   240
         TabIndex        =   19
         Top             =   6660
         Width           =   480
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "NAME"
         Height          =   180
         Left            =   240
         TabIndex        =   18
         Top             =   6240
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "JOB_DATA"
         Height          =   180
         Left            =   270
         TabIndex        =   16
         Top             =   270
         Width           =   915
      End
   End
   Begin VB.Frame Frame2 
      Height          =   7305
      Left            =   4230
      TabIndex        =   8
      Top             =   0
      Width           =   4185
      Begin VB.TextBox txtPanel_Value 
         Alignment       =   2  '가운데 맞춤
         Height          =   270
         Left            =   870
         TabIndex        =   14
         Top             =   6630
         Width           =   1785
      End
      Begin MSFlexGridLib.MSFlexGrid flxPanel_Data 
         Height          =   5475
         Left            =   180
         TabIndex        =   10
         Top             =   540
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   9657
         _Version        =   393216
      End
      Begin VB.Label lblPanel_Title 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  '단일 고정
         Height          =   270
         Left            =   870
         TabIndex        =   13
         Top             =   6210
         Width           =   1775
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "DATA"
         Height          =   180
         Left            =   240
         TabIndex        =   12
         Top             =   6660
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "NAME"
         Height          =   180
         Left            =   240
         TabIndex        =   11
         Top             =   6240
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "PANEL_DATA"
         Height          =   180
         Left            =   270
         TabIndex        =   9
         Top             =   300
         Width           =   1170
      End
   End
   Begin VB.Frame Frame1 
      Height          =   7305
      Left            =   30
      TabIndex        =   1
      Top             =   0
      Width           =   4185
      Begin VB.TextBox txtCST_Value 
         Alignment       =   2  '가운데 맞춤
         Height          =   285
         Left            =   870
         TabIndex        =   7
         Top             =   6660
         Width           =   1755
      End
      Begin MSFlexGridLib.MSFlexGrid flxCST_DATA 
         Height          =   5475
         Left            =   180
         TabIndex        =   3
         Top             =   540
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   9657
         _Version        =   393216
         Rows            =   12
      End
      Begin VB.Label lblCST_TITLE 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  '단일 고정
         Height          =   270
         Left            =   870
         TabIndex        =   6
         Top             =   6210
         Width           =   1755
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "DATA"
         Height          =   180
         Left            =   240
         TabIndex        =   5
         Top             =   6720
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "NAME"
         Height          =   180
         Left            =   240
         TabIndex        =   4
         Top             =   6270
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CST DATA"
         Height          =   180
         Left            =   270
         TabIndex        =   2
         Top             =   270
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdCreate_Data 
      Caption         =   "CREATE"
      Height          =   525
      Left            =   7680
      TabIndex        =   0
      Top             =   7590
      Width           =   1245
   End
End
Attribute VB_Name = "frmSample_Data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

