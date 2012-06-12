VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Main"
   ClientHeight    =   10644
   ClientLeft      =   168
   ClientTop       =   -4452
   ClientWidth     =   20244
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10644
   ScaleWidth      =   20244
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   20244
      _ExtentX        =   35708
      _ExtentY        =   1058
      ButtonWidth     =   2381
      ButtonHeight    =   926
      Appearance      =   1
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   7
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "HISTORY"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "VERSION"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "PTN INFO"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "SYSTEM SET."
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "RANK"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "View Log"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "EXIT"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame6 
      Height          =   4185
      Left            =   9060
      TabIndex        =   35
      Top             =   6060
      Width           =   7515
      Begin VB.PictureBox picGraph_Base 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   90
         ScaleHeight     =   3804
         ScaleWidth      =   7284
         TabIndex        =   36
         Top             =   240
         Width           =   7335
         Begin VB.Line linY 
            BorderColor     =   &H0000FFFF&
            X1              =   270
            X2              =   270
            Y1              =   330
            Y2              =   3060
         End
         Begin VB.Line linX 
            BorderColor     =   &H0000FFFF&
            X1              =   270
            X2              =   7140
            Y1              =   3060
            Y2              =   3060
         End
         Begin VB.Label lblTime 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   60
            Top             =   3150
            Width           =   90
         End
         Begin VB.Label lblTime 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   180
            Index           =   1
            Left            =   510
            TabIndex        =   59
            Top             =   3150
            Width           =   90
         End
         Begin VB.Label lblTime 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   180
            Index           =   2
            Left            =   780
            TabIndex        =   58
            Top             =   3150
            Width           =   90
         End
         Begin VB.Label lblTime 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   180
            Index           =   3
            Left            =   1020
            TabIndex        =   57
            Top             =   3150
            Width           =   90
         End
         Begin VB.Label lblTime 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   180
            Index           =   4
            Left            =   1260
            TabIndex        =   56
            Top             =   3150
            Width           =   90
         End
         Begin VB.Label lblTime 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "5"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   180
            Index           =   5
            Left            =   1500
            TabIndex        =   55
            Top             =   3150
            Width           =   90
         End
         Begin VB.Label lblTime 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "6"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   180
            Index           =   6
            Left            =   1740
            TabIndex        =   54
            Top             =   3150
            Width           =   90
         End
         Begin VB.Label lblTime 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "7"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   180
            Index           =   7
            Left            =   1980
            TabIndex        =   53
            Top             =   3150
            Width           =   90
         End
         Begin VB.Label lblTime 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "8"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   180
            Index           =   8
            Left            =   2220
            TabIndex        =   52
            Top             =   3150
            Width           =   90
         End
         Begin VB.Label lblTime 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "9"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   180
            Index           =   9
            Left            =   2460
            TabIndex        =   51
            Top             =   3150
            Width           =   90
         End
         Begin VB.Label lblTime 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "10"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   180
            Index           =   10
            Left            =   2700
            TabIndex        =   50
            Top             =   3150
            Width           =   180
         End
         Begin VB.Label lblTime 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "11"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   180
            Index           =   11
            Left            =   3030
            TabIndex        =   49
            Top             =   3150
            Width           =   180
         End
         Begin VB.Label lblTime 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "12"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   180
            Index           =   12
            Left            =   3360
            TabIndex        =   48
            Top             =   3150
            Width           =   180
         End
         Begin VB.Label lblTime 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "13"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   180
            Index           =   13
            Left            =   3690
            TabIndex        =   47
            Top             =   3150
            Width           =   180
         End
         Begin VB.Label lblTime 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "14"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   180
            Index           =   14
            Left            =   4020
            TabIndex        =   46
            Top             =   3150
            Width           =   180
         End
         Begin VB.Label lblTime 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "15"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   180
            Index           =   15
            Left            =   4350
            TabIndex        =   45
            Top             =   3150
            Width           =   180
         End
         Begin VB.Label lblTime 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "16"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   180
            Index           =   16
            Left            =   4680
            TabIndex        =   44
            Top             =   3150
            Width           =   180
         End
         Begin VB.Label lblTime 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "17"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   180
            Index           =   17
            Left            =   5010
            TabIndex        =   43
            Top             =   3150
            Width           =   180
         End
         Begin VB.Label lblTime 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "18"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   180
            Index           =   18
            Left            =   5340
            TabIndex        =   42
            Top             =   3150
            Width           =   180
         End
         Begin VB.Label lblTime 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "19"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   180
            Index           =   19
            Left            =   5670
            TabIndex        =   41
            Top             =   3150
            Width           =   180
         End
         Begin VB.Label lblTime 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "20"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   180
            Index           =   20
            Left            =   6000
            TabIndex        =   40
            Top             =   3150
            Width           =   180
         End
         Begin VB.Label lblTime 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "21"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   180
            Index           =   21
            Left            =   6330
            TabIndex        =   39
            Top             =   3150
            Width           =   180
         End
         Begin VB.Label lblTime 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "22"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   180
            Index           =   22
            Left            =   6630
            TabIndex        =   38
            Top             =   3150
            Width           =   180
         End
         Begin VB.Label lblTime 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "23"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   180
            Index           =   23
            Left            =   6960
            TabIndex        =   37
            Top             =   3150
            Width           =   180
         End
         Begin VB.Line linProduct_Count_Hour 
            BorderColor     =   &H0080FF80&
            Index           =   0
            Visible         =   0   'False
            X1              =   -60
            X2              =   330
            Y1              =   2970
            Y2              =   2850
         End
         Begin VB.Line linProduct_Count_Hour 
            BorderColor     =   &H0080FF80&
            Index           =   1
            Visible         =   0   'False
            X1              =   0
            X2              =   390
            Y1              =   120
            Y2              =   0
         End
         Begin VB.Line linProduct_Count_Hour 
            BorderColor     =   &H0080FF80&
            Index           =   2
            Visible         =   0   'False
            X1              =   0
            X2              =   390
            Y1              =   120
            Y2              =   0
         End
         Begin VB.Line linProduct_Count_Hour 
            BorderColor     =   &H0080FF80&
            Index           =   3
            Visible         =   0   'False
            X1              =   0
            X2              =   390
            Y1              =   120
            Y2              =   0
         End
         Begin VB.Line linProduct_Count_Hour 
            BorderColor     =   &H0080FF80&
            Index           =   4
            Visible         =   0   'False
            X1              =   0
            X2              =   390
            Y1              =   120
            Y2              =   0
         End
         Begin VB.Line linProduct_Count_Hour 
            BorderColor     =   &H0080FF80&
            Index           =   5
            Visible         =   0   'False
            X1              =   0
            X2              =   390
            Y1              =   120
            Y2              =   0
         End
         Begin VB.Line linProduct_Count_Hour 
            BorderColor     =   &H0080FF80&
            Index           =   6
            Visible         =   0   'False
            X1              =   0
            X2              =   390
            Y1              =   120
            Y2              =   0
         End
         Begin VB.Line linProduct_Count_Hour 
            BorderColor     =   &H0080FF80&
            Index           =   7
            Visible         =   0   'False
            X1              =   0
            X2              =   390
            Y1              =   120
            Y2              =   0
         End
         Begin VB.Line linProduct_Count_Hour 
            BorderColor     =   &H0080FF80&
            Index           =   8
            Visible         =   0   'False
            X1              =   0
            X2              =   390
            Y1              =   120
            Y2              =   0
         End
         Begin VB.Line linProduct_Count_Hour 
            BorderColor     =   &H0080FF80&
            Index           =   9
            Visible         =   0   'False
            X1              =   0
            X2              =   390
            Y1              =   120
            Y2              =   0
         End
         Begin VB.Line linProduct_Count_Hour 
            BorderColor     =   &H0080FF80&
            Index           =   10
            Visible         =   0   'False
            X1              =   0
            X2              =   390
            Y1              =   120
            Y2              =   0
         End
         Begin VB.Line linProduct_Count_Hour 
            BorderColor     =   &H0080FF80&
            Index           =   11
            Visible         =   0   'False
            X1              =   0
            X2              =   390
            Y1              =   120
            Y2              =   0
         End
         Begin VB.Line linProduct_Count_Hour 
            BorderColor     =   &H0080FF80&
            Index           =   12
            Visible         =   0   'False
            X1              =   0
            X2              =   390
            Y1              =   120
            Y2              =   0
         End
         Begin VB.Line linProduct_Count_Hour 
            BorderColor     =   &H0080FF80&
            Index           =   13
            Visible         =   0   'False
            X1              =   0
            X2              =   390
            Y1              =   120
            Y2              =   0
         End
         Begin VB.Line linProduct_Count_Hour 
            BorderColor     =   &H0080FF80&
            Index           =   14
            Visible         =   0   'False
            X1              =   0
            X2              =   390
            Y1              =   120
            Y2              =   0
         End
         Begin VB.Line linProduct_Count_Hour 
            BorderColor     =   &H0080FF80&
            Index           =   15
            Visible         =   0   'False
            X1              =   0
            X2              =   390
            Y1              =   120
            Y2              =   0
         End
         Begin VB.Line linProduct_Count_Hour 
            BorderColor     =   &H0080FF80&
            Index           =   16
            Visible         =   0   'False
            X1              =   0
            X2              =   390
            Y1              =   120
            Y2              =   0
         End
         Begin VB.Line linProduct_Count_Hour 
            BorderColor     =   &H0080FF80&
            Index           =   17
            Visible         =   0   'False
            X1              =   0
            X2              =   390
            Y1              =   120
            Y2              =   0
         End
         Begin VB.Line linProduct_Count_Hour 
            BorderColor     =   &H0080FF80&
            Index           =   18
            Visible         =   0   'False
            X1              =   0
            X2              =   390
            Y1              =   120
            Y2              =   0
         End
         Begin VB.Line linProduct_Count_Hour 
            BorderColor     =   &H0080FF80&
            Index           =   19
            Visible         =   0   'False
            X1              =   0
            X2              =   390
            Y1              =   120
            Y2              =   0
         End
         Begin VB.Line linProduct_Count_Hour 
            BorderColor     =   &H0080FF80&
            Index           =   20
            Visible         =   0   'False
            X1              =   0
            X2              =   390
            Y1              =   120
            Y2              =   0
         End
         Begin VB.Line linProduct_Count_Hour 
            BorderColor     =   &H0080FF80&
            Index           =   21
            Visible         =   0   'False
            X1              =   0
            X2              =   390
            Y1              =   120
            Y2              =   0
         End
         Begin VB.Line linProduct_Count_Hour 
            BorderColor     =   &H0080FF80&
            Index           =   22
            Visible         =   0   'False
            X1              =   0
            X2              =   390
            Y1              =   120
            Y2              =   0
         End
         Begin VB.Line linProduct_Count_Hour 
            BorderColor     =   &H0080FF80&
            Index           =   23
            Visible         =   0   'False
            X1              =   0
            X2              =   390
            Y1              =   120
            Y2              =   0
         End
      End
   End
   Begin VB.Timer tmrBulletin_Board 
      Interval        =   3000
      Left            =   10200
      Top             =   3750
   End
   Begin VB.FileListBox fleRank_File 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1584
      Left            =   18960
      TabIndex        =   34
      Top             =   8340
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Timer tmrCommand 
      Interval        =   200
      Left            =   9480
      Top             =   3960
   End
   Begin VB.Timer tmrLog 
      Interval        =   60000
      Left            =   8640
      Top             =   3840
   End
   Begin VB.Timer tmrTimeOut 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   7920
      Top             =   3600
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   408
      Left            =   0
      TabIndex        =   19
      Top             =   10236
      Width           =   20244
      _ExtentX        =   35708
      _ExtentY        =   720
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   3246
            MinWidth        =   3246
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   6774
            MinWidth        =   6774
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   23283
            MinWidth        =   23283
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   3598
            MinWidth        =   3598
            TextSave        =   "11:15"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5355
      Index           =   7
      Left            =   0
      TabIndex        =   17
      Top             =   5070
      Width           =   3135
      Begin MSFlexGridLib.MSFlexGrid flxAPI_Information 
         Height          =   2235
         Left            =   60
         TabIndex        =   30
         Top             =   3000
         Width           =   3015
         _ExtentX        =   5313
         _ExtentY        =   3937
         _Version        =   393216
         Rows            =   6
         FixedRows       =   0
      End
      Begin MSFlexGridLib.MSFlexGrid flxEQ_Information 
         Height          =   2235
         Left            =   60
         TabIndex        =   20
         Top             =   420
         Width           =   3015
         _ExtentX        =   5313
         _ExtentY        =   3937
         _Version        =   393216
         Rows            =   6
         FixedRows       =   0
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "API INFORMATION"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   90
         TabIndex        =   32
         Top             =   2760
         Width           =   1830
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "EQ INFORMATION"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   31
         Top             =   210
         Width           =   1785
      End
   End
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Index           =   5
      Left            =   3150
      TabIndex        =   16
      Top             =   6060
      Width           =   5895
      Begin MSFlexGridLib.MSFlexGrid flxRUN_Info 
         Height          =   2205
         Left            =   60
         TabIndex        =   29
         Top             =   150
         Width           =   5775
         _ExtentX        =   10181
         _ExtentY        =   3874
         _Version        =   393216
         Rows            =   6
         FixedRows       =   0
      End
   End
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   4
      Left            =   3150
      TabIndex        =   15
      Top             =   5310
      Width           =   13425
      Begin VB.Label lblPre_Loss_Code 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   10890
         TabIndex        =   28
         Top             =   210
         Width           =   1995
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "PRE PROCESS LOSS CODE"
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
         Left            =   8280
         TabIndex        =   27
         Top             =   300
         Width           =   2895
      End
      Begin VB.Label lblPost_Judge 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6330
         TabIndex        =   26
         Top             =   210
         Width           =   1635
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "POST PANEL JUDGE"
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
         Left            =   4110
         TabIndex        =   25
         Top             =   300
         Width           =   2145
      End
      Begin VB.Label lblPre_Judge 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2160
         TabIndex        =   24
         Top             =   210
         Width           =   1635
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "PRE PANEL JUDGE"
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
         Left            =   120
         TabIndex        =   23
         Top             =   300
         Width           =   1980
      End
   End
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1755
      Index           =   3
      Left            =   5640
      TabIndex        =   14
      Top             =   8490
      Width           =   3405
      Begin VB.CommandButton cmdJudge 
         Caption         =   "Manual Judge"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   20.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   33
         Top             =   210
         Width           =   3135
      End
   End
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9615
      Index           =   0
      Left            =   16590
      TabIndex        =   13
      Top             =   630
      Width           =   3765
      Begin MSFlexGridLib.MSFlexGrid flxMES_Data 
         Height          =   9405
         Left            =   30
         TabIndex        =   21
         Top             =   150
         Width           =   3645
         _ExtentX        =   6414
         _ExtentY        =   16574
         _Version        =   393216
         Rows            =   70
         FixedRows       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2745
      Left            =   3150
      TabIndex        =   12
      Top             =   2550
      Width           =   13455
      Begin MSFlexGridLib.MSFlexGrid flxJudge_History 
         Height          =   2475
         Left            =   60
         TabIndex        =   22
         Top             =   180
         Width           =   13335
         _ExtentX        =   23516
         _ExtentY        =   4360
         _Version        =   393216
         Rows            =   1
         Cols            =   7
         SelectionMode   =   1
      End
      Begin MSCommLib.MSComm MSComm 
         Index           =   1
         Left            =   600
         Top             =   120
         _ExtentX        =   995
         _ExtentY        =   995
         _Version        =   393216
         CommPort        =   2
         DTREnable       =   -1  'True
         InputLen        =   1
         OutBufferSize   =   1024
         RThreshold      =   1
         BaudRate        =   19200
         SThreshold      =   1
      End
      Begin MSCommLib.MSComm MSComm 
         Index           =   2
         Left            =   1170
         Top             =   120
         _ExtentX        =   995
         _ExtentY        =   995
         _Version        =   393216
         CommPort        =   3
         DTREnable       =   -1  'True
         InputLen        =   1
         OutBufferSize   =   1024
         RThreshold      =   1
         BaudRate        =   19200
         SThreshold      =   1
      End
      Begin MSCommLib.MSComm MSComm 
         Index           =   3
         Left            =   1740
         Top             =   120
         _ExtentX        =   995
         _ExtentY        =   995
         _Version        =   393216
         CommPort        =   4
         DTREnable       =   -1  'True
         Handshaking     =   1
         InputLen        =   1024
         OutBufferSize   =   1024
         RThreshold      =   1
         BaudRate        =   19200
         SThreshold      =   1
      End
      Begin MSCommLib.MSComm MSComm 
         Index           =   4
         Left            =   2310
         Top             =   120
         _ExtentX        =   995
         _ExtentY        =   995
         _Version        =   393216
         CommPort        =   5
         DTREnable       =   -1  'True
         InputLen        =   1
         OutBufferSize   =   1024
         RThreshold      =   1
         BaudRate        =   19200
         SThreshold      =   1
      End
      Begin MSCommLib.MSComm MSComm 
         Index           =   5
         Left            =   2880
         Top             =   120
         _ExtentX        =   995
         _ExtentY        =   995
         _Version        =   393216
         CommPort        =   6
         DTREnable       =   -1  'True
         InputLen        =   1
         OutBufferSize   =   1024
         RThreshold      =   1
         BaudRate        =   19200
         SThreshold      =   1
      End
      Begin MSCommLib.MSComm MSComm 
         Index           =   6
         Left            =   3450
         Top             =   120
         _ExtentX        =   995
         _ExtentY        =   995
         _Version        =   393216
         CommPort        =   7
         DTREnable       =   -1  'True
         InputLen        =   1
         OutBufferSize   =   1024
         RThreshold      =   1
         BaudRate        =   19200
         SThreshold      =   1
      End
      Begin MSCommLib.MSComm MSComm 
         Index           =   7
         Left            =   4020
         Top             =   120
         _ExtentX        =   995
         _ExtentY        =   995
         _Version        =   393216
         CommPort        =   8
         DTREnable       =   -1  'True
         Handshaking     =   1
         InputLen        =   1
         OutBufferSize   =   1024
         RThreshold      =   1
         BaudRate        =   19200
         SThreshold      =   1
      End
      Begin MSCommLib.MSComm MSComm 
         Index           =   0
         Left            =   30
         Top             =   120
         _ExtentX        =   995
         _ExtentY        =   995
         _Version        =   393216
         DTREnable       =   -1  'True
         InputLen        =   1
         OutBufferSize   =   1024
         RThreshold      =   1
         BaudRate        =   19200
         SThreshold      =   1
      End
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1755
      Left            =   3150
      TabIndex        =   10
      Top             =   8490
      Width           =   2475
      Begin VB.CommandButton cmdForced_Unload 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Manual Unload"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1395
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   2205
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1845
      Left            =   0
      TabIndex        =   7
      Top             =   3210
      Width           =   3135
      Begin MSFlexGridLib.MSFlexGrid flxPre_Align_PanelID 
         Height          =   825
         Left            =   60
         TabIndex        =   8
         Top             =   960
         Width           =   3015
         _ExtentX        =   5313
         _ExtentY        =   1461
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
      End
      Begin MSFlexGridLib.MSFlexGrid flxAlign_PanelID 
         Height          =   825
         Left            =   60
         TabIndex        =   9
         Top             =   150
         Width           =   3015
         _ExtentX        =   5313
         _ExtentY        =   1439
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Index           =   2
      Left            =   0
      TabIndex        =   5
      Top             =   1410
      Width           =   3135
      Begin MSFlexGridLib.MSFlexGrid flxStatus 
         Height          =   1515
         Left            =   60
         TabIndex        =   6
         Top             =   180
         Width           =   3015
         _ExtentX        =   5313
         _ExtentY        =   2667
         _Version        =   393216
         Rows            =   4
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Index           =   1
      Left            =   3150
      TabIndex        =   1
      Top             =   630
      Width           =   13425
      Begin VB.ListBox lstMessage 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.6
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   1416
         ItemData        =   "frmMain.frx":08CA
         Left            =   60
         List            =   "frmMain.frx":08CC
         TabIndex        =   2
         Top             =   210
         Width           =   13245
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   630
      Width           =   3135
      Begin VB.Label lblUser 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.6
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   4
         Top             =   240
         Width           =   2265
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "USER"
         BeginProperty Font 
            Name            =   "ËÎÌå"
            Size            =   9.6
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   330
         Width           =   525
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuOff_Line_Request 
         Caption         =   "Device Off Line Request"
      End
      Begin VB.Menu mnuGrade_Test 
         Caption         =   "For Grade Test"
         Begin VB.Menu mnuOn_Line 
            Caption         =   "On Line Connect"
         End
         Begin VB.Menu mnuBefore_Block_Contact 
            Caption         =   "Before Block Contact"
         End
         Begin VB.Menu mnu_After_Block_Contact 
            Caption         =   "After Block Contact"
         End
         Begin VB.Menu RBBU 
            Caption         =   "RBBU"
         End
         Begin VB.Menu mnuAddress_Info_Normal 
            Caption         =   "Address_Info (Normal)"
         End
         Begin VB.Menu mnuAddress_Info_Mura 
            Caption         =   "Address_Info (Mura)"
         End
      End
      Begin VB.Menu mnuUser_Regist 
         Caption         =   "User Regist"
      End
   End
   Begin VB.Menu mnuPG_Command 
      Caption         =   "PG Command"
      Visible         =   0   'False
      Begin VB.Menu mnuPG_Online_Request 
         Caption         =   "On Line Request"
      End
      Begin VB.Menu mnuSetting_Modify 
         Caption         =   "Setting Modify"
      End
      Begin VB.Menu mnuPG_Power_On 
         Caption         =   "PG Power On"
      End
      Begin VB.Menu mnuPG_Power_Off 
         Caption         =   "PG Power Off"
      End
   End
   Begin VB.Menu mnuAPI_Command 
      Caption         =   "API Command"
      Visible         =   0   'False
      Begin VB.Menu mnuAPI_On_Line_Request 
         Caption         =   "On Line Request"
      End
      Begin VB.Menu mnuMES_Data_Send 
         Caption         =   "MES Data Send"
      End
      Begin VB.Menu mnuAPI_EQP_Statue_Request 
         Caption         =   "API Status Request"
      End
   End
   Begin VB.Menu mnuEQ_Command 
      Caption         =   "EQ Command"
      Visible         =   0   'False
      Begin VB.Menu mnuEQP_Status_Request 
         Caption         =   "EQP Status Request"
      End
      Begin VB.Menu mnuEQ_Buzz_Send 
         Caption         =   "Buzz/Message Send"
      End
      Begin VB.Menu mnuEQ_Signal_On 
         Caption         =   "Signal On"
      End
      Begin VB.Menu mnuEQ_Signal_Off 
         Caption         =   "Signal Off"
      End
   End
   Begin VB.Menu mnuMES_Command 
      Caption         =   "MES Command"
      Visible         =   0   'False
      Begin VB.Menu mnuLoad_MES_Data 
         Caption         =   "Load MES Data"
      End
      Begin VB.Menu mnuData_Input 
         Caption         =   "Data Input"
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit"
   End
   Begin VB.Menu mnuEQP_Menu 
      Caption         =   ""
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_NEW_DATE                      As Boolean
Dim m_PRODUCT_COUNT_RESET           As Boolean

Private Sub cmdForced_Unload_Click()

    Dim strCommand                  As String
    Dim strDeviceState              As String
    Dim strSortFlag                 As String
    
    Dim intLength                   As Integer
    Dim intPortNo                   As Integer
    
    If MsgBox("Are you sure panel unload by manual?", vbYesNo, "Panel manual unload") = vbYes Then
        Call ENV.Get_Device_Data_by_Name(ENV.Get_Current_Prober_Name, intPortNo, strDeviceState)
        
        If intPortNo > 0 Then
            strCommand = "QSPO"
            
            intLength = cSIZE_PANELID - Len(Trim(Me.flxAlign_PanelID.TextMatrix(1, 0)))
            strCommand = strCommand & Trim(Me.flxAlign_PanelID.TextMatrix(1, 0)) & Space(intLength) & "M"
            
            Call QUEUE.Put_Send_Command(intPortNo, strCommand)
        End If
    End If
    
End Sub

Private Sub cmdJudge_Click()

    frmJudge.Show
    
End Sub

Private Sub flxMES_Data_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        Me.PopupMenu Me.mnuMES_Command
    End If
    
End Sub

Private Sub flxStatus_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim strTime                     As String
    
    Dim intRow                      As Integer
    
    intRow = Me.flxStatus.Row
    
    If Button = vbRightButton Then
        Select Case intRow
        Case 1:
            Me.PopupMenu frmMain.mnuEQ_Command
            If InStr(Me.flxEQ_Information.TextMatrix(5, 1), "LOI") = 0 Then
                Me.mnuEQ_Signal_On.Visible = False
                Me.mnuEQ_Signal_Off.Visible = False
            Else
                Me.mnuEQ_Signal_On.Visible = True
                Me.mnuEQ_Signal_Off.Visible = True
            End If
        Case 2:
            Me.PopupMenu frmMain.mnuAPI_Command
        Case 3:
            Me.PopupMenu frmMain.mnuPG_Command
        End Select
    End If
    
End Sub

Private Sub Form_Load()

    Dim intPortNo           As Integer
    
On Error GoTo ErrorHandler

    Call Init_Form
    Call Init_Grid
    Call Fill_Grid
        
    Me.MSComm(0).PortOpen = True
    m_NEW_DATE = False
    m_PRODUCT_COUNT_RESET = False
    Call EQP.Set_DEFECT_UPLOAD(True)
    
    Me.Left = 0
    Me.Top = 0
    Me.Height = 11535
    Me.Width = 20490
    
    With Me.Toolbar1
        .Buttons(1).Enabled = False
        .Buttons(2).Enabled = False
        .Buttons(3).Enabled = False
        .Buttons(4).Enabled = False
        .Buttons(5).Enabled = False
        .Buttons(6).Enabled = False
        .Buttons(7).Enabled = True
    End With
    Me.mnuTools.Enabled = False
    
    Me.Caption = "JPS Program" & " - Version : " & App.Major & "." & App.Minor & "." & App.Revision
    Me.StatusBar.Panels(1).Text = "JPS : " & App.Major & "." & App.Minor & "." & App.Revision
    
    For intPortNo = 0 To 7
        If ENV.Get_Port_Use(intPortNo + 1) = True Then
            If Me.MSComm(intPortNo).PortOpen = False Then
                Me.MSComm(intPortNo).PortOpen = True
                Call SaveLog("Form_Load", intPortNo + 1 & " port open")
            End If
        End If
    Next intPortNo
    
    Exit Sub
    
ErrorHandler:

    If intPortNo >= 0 Then
        Load frmSystem_Parameter
        frmSystem_Parameter.tabParameter.Tab = 1
        frmSystem_Parameter.Show
        Call Show_Message("Port state error", "Port" & intPortNo + 1 & " is not available.")
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If MsgBox("Do you want exit program?", vbYesNo, "Confirm") = vbYes Then
        Unload Me
    Else
        Cancel = 1
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Dim strLog                  As String
    
    Dim intPortNo               As Integer
    
    For intPortNo = 0 To 7
        If ENV.Get_Port_Use(intPortNo + 1) = True Then
            If Me.MSComm(intPortNo).PortOpen = True Then
                Me.MSComm(intPortNo).PortOpen = False
                Call SaveLog("Form_Unload", intPortNo + 1 & " port close")
            End If
        End If
    Next intPortNo
    
    strLog = "[Main " & Format(DATE, "YYYY-MM-DD") & " " & Format(TIME, "hh:mm:ss") & "] JPS program end."
    Call Write_Log(App.PATH & "\Log\," & Format(DATE, "YYYYMMDD") & "_" & Format(TIME, "hh") & ".Log," & strLog)
    
    End
    
End Sub

Private Sub lblUser_DblClick()

    Dim strMode_State               As String
    
'    Select Case Me.flxEQ_Information.TextMatrix(3, 1)
'    Case "Operator":
'        strMode_State = "ON"
'    Case "Auto and RJS":
'        strMode_State = "IA"
'    Case "Full Auto":
'        strMode_State = "FA"
'    Case "EQ Pass":
'        strMode_State = "EP"
'    End Select
'    If (strMode_State = "ON") Or (Me.lblUser.Caption = "") Then
        Load frmLogin
        frmLogin.Show
'    End If
    
End Sub

Private Sub lstSystem_Log_DblClick()

    Dim strFileName                 As String
    Dim strPath                     As String
    
    Dim varResult                   As Variant
    
    strPath = App.PATH & "\Log\"
    strFileName = Format(DATE, "YYYYMMDD") & "_" & Format(TIME, "HH") & ".Log"
    If Dir(strPath & strFileName, vbNormal) <> "" Then
        varResult = Shell("NotePad.exe " & strPath & strFileName, vbNormalFocus)
    Else
        strFileName = Format(DATE - (1 / 24), "YYYYMMDD") & "_" & Format(TIME - (1 / 24), "HH") & ".Log"
        If Dir(strPath & strFileName, vbNormal) <> "" Then
            varResult = Shell("NotePad.exe " & strPath & strFileName, vbNormalFocus)
        End If
    End If
    
End Sub

Private Sub mnu_After_Block_Contact_Click()

    Dim strPath                         As String
    Dim strFileName                     As String
    Dim strCommand                      As String
    
    Dim intFileNum                      As Integer
    
    strPath = App.PATH & "\Env\"
    strFileName = "RABC.txt"
    If Dir(strPath & strFileName, vbNormal) <> "" Then
        intFileNum = FreeFile
        Open strPath & strFileName For Input As intFileNum
        
        While Not EOF(intFileNum)
            Line Input #intFileNum, strCommand
        Wend
        
        Close intFileNum
    End If
    
    Call QUEUE.Put_Receive_Command(4, strCommand)
'    If strCommand <> "" Then
'        If (Left(strCommand, 1) = cSTX) And (Right(strCommand, 1) = cETX) Then
'            strCommand = Mid(strCommand, 2, Len(strCommand) - 2)
'        End If
'        Call BTST_Sequence(4, strCommand)
'    End If

End Sub

Private Sub mnuAddress_Info_Mura_Click()

    Dim strCommand                      As String
    
    strCommand = "RRAD003420002301100004500027900253"
    
    Call API_Sequence(9, strCommand)

End Sub

Private Sub mnuAddress_Info_Normal_Click()

    Dim strCommand                      As String
    
    strCommand = "RRAD0122000411                    "
    
    Call API_Sequence(9, strCommand)

End Sub

Private Sub mnuAPI_EQP_Statue_Request_Click()

    Dim intPortNo                   As Integer
    
    Dim strState                    As String
    
    Call ENV.Get_Device_Data_by_Name("API", intPortNo, strState)
    
    If intPortNo > 0 Then
        Call QUEUE.Put_Send_Command(intPortNo, "QEQS")
    Else
        Call SaveLog("mnuAPI_EQP_Status_Request_Click", "Port Number : 0")
    End If
    
End Sub

Private Sub mnuAPI_On_Line_Request_Click()

    Dim strTime                     As String
    Dim strCommand                  As String
    Dim strStatus                   As String
    Dim strDevice_Name              As String
    Dim strMode_State               As String
    
    Dim intPortNo                   As Integer
    
    strTime = Format(DATE, "YYYYMMDD") & Format(TIME, "HHMMSS")
    With frmMain.flxEQ_Information
        Select Case .TextMatrix(3, 1)
        Case "Operator":
            strMode_State = "ON"
        Case "Auto and RJS":
            strMode_State = "IA"
        Case "Full Auto":
            strMode_State = "FA"
        Case "EQ Pass":
            strMode_State = "EP"
        End Select
    
        strCommand = "QONA" & strTime & .TextMatrix(1, 1) & .TextMatrix(2, 1) & strMode_State & "CAAPI" & Right(.TextMatrix(5, 1), 3) & Me.lblUser.Caption
        'Lucas----If PFCD is Space.Then it didn't Send Command
     If .TextMatrix(2, 1) <> "            " Then
        Call ENV.Get_Device_Data_by_Name("API", intPortNo, strStatus)
        If intPortNo = 0 Then
            For intPortNo = 1 To 8
                If ENV.Get_Port_Use(intPortNo) = True Then
                    Call ENV.Get_Device_Data_by_PortID(intPortNo, strDevice_Name, strStatus)
                    If strDevice_Name = "" Then
                        If Me.MSComm(intPortNo - 1).PortOpen = True Then
                            Call QUEUE.Put_Send_Command(intPortNo, strCommand)
                        End If
                    End If
                End If
            Next intPortNo
'Lucas 2012.01.05 Ver.0.9.2 ----For CAAPI online
'        Else
'            If .TextMatrix(2, 1) <> "            " Then
'            Call QUEUE.Put_Send_Command(intPortNo, strCommand)
'            End If
        End If
     End If
    End With
    
    With Me.flxAPI_Information
        .TextMatrix(0, 1) = strTime
        .TextMatrix(1, 1) = Me.flxEQ_Information.TextMatrix(1, 1)
        .TextMatrix(2, 1) = Me.flxEQ_Information.TextMatrix(2, 1)
'        .TextMatrix(3, 1) = Me.flxEQ_Information.TextMatrix(3, 1)
        .TextMatrix(5, 1) = "CAAPI"
    End With
    
End Sub

Private Sub mnuBefore_Block_Contact_Click()

    Dim strPath                         As String
    Dim strFileName                     As String
    Dim strCommand                      As String
    
    Dim intFileNum                      As Integer
    
    strPath = App.PATH & "\Env\"
    strFileName = "RBBC.txt"
    If Dir(strPath & strFileName, vbNormal) <> "" Then
        intFileNum = FreeFile
        Open strPath & strFileName For Input As intFileNum
        
        While Not EOF(intFileNum)
            Line Input #intFileNum, strCommand
        Wend
        
        Close intFileNum
    End If
    
    Call QUEUE.Put_Receive_Command(4, strCommand)
'    If strCommand <> "" Then
'        If (Left(strCommand, 1) = cSTX) And (Right(strCommand, 1) = cETX) Then
'            strCommand = Mid(strCommand, 2, Len(strCommand) - 2)
'        End If
'        Call BTST_Sequence(4, strCommand)
'    End If

End Sub

Private Sub mnuData_Input_Click()

    Dim intRow                      As Integer
    Dim intSize                     As Integer
    
    Dim bolLoad_Input               As Boolean
    
    intRow = Me.flxMES_Data.Row
    
    If intRow > 0 Then
        bolLoad_Input = False
        Select Case intRow
        Case 1:
            bolLoad_Input = True
            intSize = cSIZE_PFCD
        Case 2:
            bolLoad_Input = True
            intSize = cSIZE_OWNER_MES
        Case 3:
            bolLoad_Input = True
            intSize = cSIZE_PROCESSNUM_MES
        Case 17:
            bolLoad_Input = True
            intSize = cSIZE_PANELID
        End Select
        
        If bolLoad_Input = True Then
            Load frmMes_Data_Input
            frmMes_Data_Input.lblRow.Caption = intRow
            frmMes_Data_Input.lblTitle.Caption = Me.flxMES_Data.TextMatrix(intRow, 0)
            frmMes_Data_Input.txtMes_Data.MaxLength = intSize
            frmMes_Data_Input.Show
        End If
    End If

End Sub

Private Sub mnuEQ_Buzz_Send_Click()

    Load frmSend_Message
    frmSend_Message.Show
    
End Sub

Private Sub mnuEQ_Signal_Off_Click()

    Dim intPortNo           As Integer
    
    Dim strStatus           As String
    
    Call ENV.Get_Device_Data_by_Name(Left(frmMain.flxEQ_Information.TextMatrix(5, 1), 5), intPortNo, strStatus)
    
    If intPortNo > 0 Then
        Call QUEUE.Put_Send_Command(intPortNo, "QSOF")
    Else
        Call SaveLog("mnuEQ_Signal_Off_Click", "Port Number : 0")
    End If
    
End Sub

Private Sub mnuEQ_Signal_On_Click()

    Dim intPortNo           As Integer
    
    Dim strStatus           As String
    
    Call ENV.Get_Device_Data_by_Name(Left(frmMain.flxEQ_Information.TextMatrix(5, 1), 5), intPortNo, strStatus)
    
    If intPortNo > 0 Then
        Call QUEUE.Put_Send_Command(intPortNo, "QSON")
    Else
        Call SaveLog("mnuEQ_Signal_On_Click", "Port Number : 0")
    End If
    
End Sub

Private Sub mnuEQP_Status_Request_Click()

    Dim intPortNo           As Integer
    
    Dim strStatus           As String
    
    Call ENV.Get_Device_Data_by_Name(Left(frmMain.flxEQ_Information.TextMatrix(5, 1), 5), intPortNo, strStatus)
    
    If intPortNo > 0 Then
        Call QUEUE.Put_Send_Command(intPortNo, "QEQS")
    Else
        Call SaveLog("mnuEQP_Status_Request_Click", "Port Number : 0")
    End If
    
End Sub

Private Sub mnuExit_Click()

    Unload Me
    
End Sub

Private Sub Init_Form()

    Dim typCOUNT_CHANGE     As COUNT_CHANGE_DATA
    
    Dim strPath             As String
    Dim strFileName         As String
    Dim strTemp             As String
    
    Dim intFileNum          As Integer
    Dim intPos              As Integer
    
    strPath = App.PATH & "\Env\"
    strFileName = "Auto_Grade.cfg"
    
    If Dir(strPath & strFileName, vbNormal) <> "" Then
        intFileNum = FreeFile
        
        Open strPath & strFileName For Input As intFileNum
        
        While Not EOF(intFileNum)
            Line Input #intFileNum, strTemp
            intPos = InStr(strTemp, "=")
            If intPos > 0 Then
                With typCOUNT_CHANGE
                    Select Case Left(strTemp, intPos - 1)
                    Case "FINAL RANK":
                        .FINAL_GRADE = Mid(strTemp, intPos + 1)
                    Case "CHANGE GRADE":
                        .NEW_GRADE = Mid(strTemp, intPos + 1)
                    Case "COUNT":
                        .Count = CInt(Mid(strTemp, intPos + 1))
                    Case "CURRENT COUNT":
                        .CURRENT_COUNT = CInt(Mid(strTemp, intPos + 1))
                    End Select
                End With
            End If
        Wend
        
        Close intFileNum

        frmMain.StatusBar.Panels(2).Text = "Remained Grade Count : " & typCOUNT_CHANGE.CURRENT_COUNT
    Else
        frmMain.StatusBar.Panels(2).Text = "Remained Grade Count : no data"
    End If
'    strPath = App.Path & "\Env\"
'    strFileName = "GRADE_LIST.dat"
'
'    If Dir(strPath & strFileName, vbNormal) <> "" Then
'        intFileNum = FreeFile
'        Open strPath & strFileName For Input As intFileNum
'
'        While Not EOF(intFileNum)
'            Line Input #intFileNum, strTemp
'            Me.cmbGrade_List.AddItem strTemp
'        Wend
'
'        Close intFileNum
'
'        Me.cmbGrade_List.Text = Me.cmbGrade_List.List(0)
'    End If
    
 '============Leo 2012.05.22 Add Rank Level
  Call RANK_OBJ.Get_Rank_Levels
End Sub

Private Sub Init_Grid()

    Dim intRow              As Integer
    Dim intCol              As Integer
    
    With Me.flxStatus
        For intRow = 0 To .Rows - 1
            For intCol = 0 To .Cols - 1
                .Row = intRow
                .Col = intCol
                .CellAlignment = flexAlignCenterCenter
                .RowHeight(intRow) = 350
                If (intRow <> 0) And (intCol = 1) Then
                    .CellBackColor = vbRed
                    .CellForeColor = vbBlack
                End If
            Next intCol
        Next intRow
        
        .ColWidth(0) = 1400
        .TextMatrix(0, 0) = "DEVICE"
        .ColWidth(1) = 1500
        .TextMatrix(0, 1) = "STATUS"
        
        .TextMatrix(1, 0) = "PROBER"
        .TextMatrix(2, 0) = "API"
        .TextMatrix(3, 0) = "PG"
        
        .TextMatrix(1, 1) = "OFF LINE"
        .TextMatrix(2, 1) = "OFF LINE"
        .TextMatrix(3, 1) = "OFF LINE"
    End With
    
    With Me.flxPre_Align_PanelID
        For intRow = 0 To .Rows - 1
            .Row = intRow
            .Col = 0
            .CellAlignment = flexAlignCenterCenter
            .RowHeight(intRow) = 350
        Next intRow
        
        .ColWidth(0) = 2900
        
        .TextMatrix(0, 0) = "Pre-Alignment PANEL"
    End With
    
    With Me.flxAlign_PanelID
        For intRow = 0 To .Rows - 1
            .Row = intRow
            .Col = 0
            .CellAlignment = flexAlignCenterCenter
            .RowHeight(intRow) = 350
        Next intRow
        
        .ColWidth(0) = 2900
        
        .TextMatrix(0, 0) = "Alignment PANEL"
    End With
    
    With Me.flxEQ_Information
        For intRow = 0 To .Rows - 1
            .Row = intRow
            For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellAlignment = flexAlignCenterCenter
                .RowHeight(intRow) = 350
            Next intCol
        Next intRow
        
        .ColWidth(0) = 1400
        .ColWidth(1) = 1500
        
        .TextMatrix(0, 0) = "ON TIME"
        .TextMatrix(1, 0) = "DRIVE TYPE"
        .TextMatrix(2, 0) = "PFCD"
        .TextMatrix(3, 0) = "RUN MODE"
        .TextMatrix(4, 0) = "RUN STATE"
        .TextMatrix(5, 0) = "EQP NAME"
    End With
    
    With Me.flxAPI_Information
        For intRow = 0 To .Rows - 1
            .Row = intRow
            For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellAlignment = flexAlignCenterCenter
                .RowHeight(intRow) = 350
            Next intCol
        Next intRow
        
        .ColWidth(0) = 1400
        .ColWidth(1) = 1500
        
        .TextMatrix(0, 0) = "ON TIME"
        .TextMatrix(1, 0) = "DRIVE TYPE"
        .TextMatrix(2, 0) = "PFCD"
        .TextMatrix(3, 0) = "RUN MODE"
        .TextMatrix(4, 0) = "RUN STATE"
        .TextMatrix(5, 0) = "EQP NAME"
    End With
    
    With Me.flxMES_Data
        For intRow = 0 To .Rows - 1
            For intCol = 0 To .Cols - 1
                .Row = intRow
                .Col = intCol
                .CellAlignment = flexAlignCenterCenter
                .RowHeight(intRow) = 270
            Next intCol
        Next intRow
        
        .Col = 0
        .ColWidth(0) = 1500
        .ColWidth(1) = 1800
    
        .TextMatrix(0, 0) = "CST ID"
        .TextMatrix(1, 0) = "PFCD"
        .Row = 1
        .CellForeColor = vbBlue
        .CellFontBold = True
        .TextMatrix(2, 0) = "OWNER"
        .Row = 2
        .CellForeColor = vbBlue
        .CellFontBold = True
        .TextMatrix(3, 0) = "Process Num"
        .Row = 3
        .CellForeColor = vbBlue
        .CellFontBold = True
        .TextMatrix(4, 0) = "Port ID"
        .TextMatrix(5, 0) = "Port Type"
        .TextMatrix(6, 0) = "Destination Fab"
        .TextMatrix(7, 0) = "PanelCount"
        .TextMatrix(8, 0) = "RMANO"
        .TextMatrix(9, 0) = "OQCNO"
        .TextMatrix(10, 0) = "Source Fab"
        .TextMatrix(11, 0) = "CST SPARE1"
        .TextMatrix(12, 0) = "CST SPARE2"
        .TextMatrix(13, 0) = "CST SPARE3"
        .TextMatrix(14, 0) = "CST SPARE4"
        .TextMatrix(15, 0) = "CST SPARE5"
        .TextMatrix(16, 0) = "Slot Num"
        .TextMatrix(17, 0) = "Panel ID"
        .Row = 17
        .CellForeColor = vbBlue
        .CellFontBold = True
        .TextMatrix(18, 0) = "LightON Panel Grade"
        .TextMatrix(19, 0) = "LightON Reason Code"
        .TextMatrix(20, 0) = "Cell Rescue Flag"
        .TextMatrix(21, 0) = "Cell Repair Grade"
        .TextMatrix(22, 0) = "TFT Repair Grade"
        .TextMatrix(23, 0) = "CF Panel ID"
        .TextMatrix(24, 0) = "CF O/X Information"
        .TextMatrix(25, 0) = "Panel Owner Type"
        .TextMatrix(26, 0) = "Abnormal CF"
        .TextMatrix(27, 0) = "Abnormal TFT"
        .TextMatrix(28, 0) = "Abnormal LCD"
        .TextMatrix(29, 0) = "Group ID"
        .TextMatrix(30, 0) = "Repair Rework Count"
        .TextMatrix(31, 0) = "Carbonization FLAG"
        .TextMatrix(32, 0) = "Carbonization Grade"
        .TextMatrix(33, 0) = "Carbonization R/W Count"
        .TextMatrix(34, 0) = "Polarizer R/W Count"
        .TextMatrix(35, 0) = "X TOTAL PIXEL"
        .TextMatrix(36, 0) = "Y TOTAL PIXEL"
        .TextMatrix(37, 0) = "X one Pixel Length"
        .TextMatrix(38, 0) = "Y one Pixel Length"
        .TextMatrix(39, 0) = "LCD Qtap LotGroupID"
        .TextMatrix(40, 0) = "SK flag"
        .TextMatrix(41, 0) = "CF R Defect code"
        .TextMatrix(42, 0) = "ODF AK Flag"
        .TextMatrix(43, 0) = "BPAM rework flag"
        .TextMatrix(44, 0) = "LCD Bright dot flag"
        .TextMatrix(45, 0) = "CFPS hight err flag"
        .TextMatrix(46, 0) = "PI Insp. NG flag"
        .TextMatrix(47, 0) = "PI over bake flag"
        .TextMatrix(48, 0) = "PI over Q time flag"
        .TextMatrix(49, 0) = "ODF over bake flag"
        .TextMatrix(50, 0) = "ODF over Qtime flag"
        .TextMatrix(51, 0) = "HVA over bake flag"
        .TextMatrix(52, 0) = "HVA over Qtime flag"
        .TextMatrix(53, 0) = "Seal Insp. flag"
        .TextMatrix(54, 0) = "ODF checker flag"
        .TextMatrix(55, 0) = "ODF Door open flag"
        .TextMatrix(56, 0) = "LOT1 Operation mode"
        .TextMatrix(57, 0) = "LOT2 Operation mode"
        .TextMatrix(58, 0) = "Product ID"
        .TextMatrix(59, 0) = "OWNER ID"
        .TextMatrix(60, 0) = "Panel SPARE1"
        .TextMatrix(61, 0) = "Panel SPARE2"
        .TextMatrix(62, 0) = "Panel SPARE3"
        .TextMatrix(63, 0) = "Panel SPARE4"
        .TextMatrix(64, 0) = "Panel SPARE5"
        .TextMatrix(65, 0) = "Panel SPARE6"
        .TextMatrix(66, 0) = "Panel SPARE7"
        .TextMatrix(67, 0) = "Panel SPARE8"
        .TextMatrix(68, 0) = "Panel SPARE9"
        .TextMatrix(69, 0) = "Panel SPARE10"
    End With
    
    With Me.flxJudge_History
        For intRow = 0 To .Rows - 1
            For intCol = 0 To .Cols - 1
                .Row = intRow
                .Col = intCol
                .CellAlignment = flexAlignCenterCenter
                .RowHeight(intRow) = 350
            Next intCol
        Next intRow
                
        For intCol = 0 To .Cols - 1
            If intCol = 0 Then
                .ColWidth(intCol) = 2000
            Else
                .ColWidth(intCol) = 1800
            End If
        Next intCol
        
        .TextMatrix(0, 0) = "PANEL ID"
        .TextMatrix(0, 1) = "USER INFO"
        .TextMatrix(0, 2) = "PROCESS No."
        .TextMatrix(0, 3) = "PANEL GRADE."
        .TextMatrix(0, 4) = "LOSS CODE"
        .TextMatrix(0, 5) = "DEFECT NAME"
        .TextMatrix(0, 6) = "JUDGE TIME"
    End With
    
    With Me.flxRUN_Info
        For intRow = 0 To .Rows - 1
            For intCol = 0 To .Cols - 1
                .Row = intRow
                .Col = intCol
                .CellAlignment = flexAlignCenterCenter
                .RowHeight(intRow) = 350
            Next intCol
        Next intRow
        
        .ColWidth(0) = 3000
        .ColWidth(1) = 2500
        
        .TextMatrix(0, 0) = "PRODUCTION TARGET"
        .TextMatrix(1, 0) = "CURRENT PRODUCT COUNT"
        .TextMatrix(2, 0) = "TOTAL PORODUCT COUNT"
        .TextMatrix(3, 0) = "PRODUCT COUNT PER HOUR"
        .TextMatrix(4, 0) = "AVERAGE TACT TIME"
        .TextMatrix(5, 0) = "S/W START TIME"
    End With
'
'    With Me.flxLoss_Code
'        For intRow = 0 To .Rows - 1
'            For intCol = 0 To .Cols - 1
'                .Row = intRow
'                .Col = intCol
'                .CellAlignment = flexAlignCenterCenter
'                .RowHeight(intRow) = 350
'            Next intCol
'        Next intRow
'
'        .ColWidth(0) = 1500
'        .ColWidth(1) = 3000
'
'        .TextMatrix(0, 0) = "LOSS CODE"
'        .TextMatrix(0, 1) = "DESCRIPTION"
'    End With
    
End Sub

Private Sub Fill_Grid()

    Call Set_RUN_Data
    
End Sub

Private Sub mnuLoad_MES_Data_Click()

    Dim typCST_INFO             As CST_INFO_ELEMENTS
    Dim typPANEL_INFO           As PANEL_INFO_ELEMENTS
    
    Dim strPath                 As String
    Dim strFileName             As String
    Dim strCST_Info_Length      As String
    Dim strPanel_Info_Length    As String
    Dim strCommand              As String
    Dim strSUB_Command          As String
    Dim strMES_DATA_Command     As String
    Dim strMsg                  As String
    
    Dim intFileNum              As Integer
    Dim intIndex                As Integer
    
    strPath = App.PATH & "\Env\"
    strFileName = "RBBC.txt"
    
    If Dir(strPath & strFileName, vbNormal) <> "" Then
        intFileNum = FreeFile
        
        Open strPath & strFileName For Input As intFileNum
        
        Line Input #intFileNum, strCommand
        
        strCommand = Mid(strCommand, 6)
        strCommand = Mid(strCommand, cSIZE_PANELID + (cSIZE_FLAG * 3) + 1)
                    
        strCST_Info_Length = Mid(strCommand, 1, cSIZE_INFO_LENGTH)
        strMES_DATA_Command = strCST_Info_Length
        strCommand = Mid(strCommand, cSIZE_INFO_LENGTH + 1)
        strSUB_Command = Mid(strCommand, 1, CInt(strCST_Info_Length))
        strMES_DATA_Command = strMES_DATA_Command & strSUB_Command
        strCommand = Mid(strCommand, CInt(strCST_Info_Length) + 1)
        Call Decode_CST_Information_Elements(strSUB_Command, typCST_INFO)
        
        strPanel_Info_Length = Mid(strCommand, 1, cSIZE_INFO_LENGTH)
        strMES_DATA_Command = strMES_DATA_Command & strPanel_Info_Length
        strCommand = Mid(strCommand, cSIZE_INFO_LENGTH + 1)
        strSUB_Command = Mid(strCommand, 1, CInt(strPanel_Info_Length))
        strMES_DATA_Command = strMES_DATA_Command & strSUB_Command
        strCommand = Mid(strCommand, CInt(strPanel_Info_Length) + 1)
        Call Decode_PANEL_Information_Elements(strSUB_Command, typPANEL_INFO, typCST_INFO.PFCD)
            
        Close intFileNum
        
        With typCST_INFO
            Me.flxMES_Data.TextMatrix(0, 1) = .CSTID
            Me.flxMES_Data.TextMatrix(1, 1) = .PFCD
            Me.flxMES_Data.TextMatrix(2, 1) = .OWNER
            Me.flxMES_Data.TextMatrix(3, 1) = .PROCESS_NUM
            Me.flxMES_Data.TextMatrix(4, 1) = .PORTID
            Me.flxMES_Data.TextMatrix(5, 1) = .PORT_TYPE
            Me.flxMES_Data.TextMatrix(6, 1) = .DESTINATION_FAB
            Me.flxMES_Data.TextMatrix(7, 1) = .PANEL_COUNT
            Me.flxMES_Data.TextMatrix(8, 1) = .RMANO
            Me.flxMES_Data.TextMatrix(9, 1) = .OQCNO
            Me.flxMES_Data.TextMatrix(10, 1) = .SOURCE_FAB
            For intIndex = 1 To 5
                Me.flxMES_Data.TextMatrix(10 + intIndex, 1) = .CST_SPARE(intIndex)
            Next intIndex
        End With
        
        With typPANEL_INFO
            Me.flxMES_Data.TextMatrix(16, 1) = .SLOT_NUM
            Me.flxMES_Data.TextMatrix(17, 1) = .PANELID
            Me.flxMES_Data.TextMatrix(18, 1) = .LIGHT_ON_PANEL_GRADE
            Me.flxMES_Data.TextMatrix(19, 1) = .LIGHT_ON_REASON_CODE
            Me.flxMES_Data.TextMatrix(20, 1) = .CELL_LINE_RESCUE_FLAG
            Me.flxMES_Data.TextMatrix(21, 1) = .CELL_REPAIR_JUDGE_GRADE
            Me.flxMES_Data.TextMatrix(22, 1) = .TFT_REPAIR_GRADE
            Me.flxMES_Data.TextMatrix(23, 1) = .CF_PANELID
            Me.flxMES_Data.TextMatrix(24, 1) = .CF_PANEL_OX_INFORMATION
            Me.flxMES_Data.TextMatrix(25, 1) = .PANEL_OWNER_TYPE
            Me.flxMES_Data.TextMatrix(26, 1) = .ABNORMAL_CF
            Me.flxMES_Data.TextMatrix(27, 1) = .ABNORMAL_TFT
            Me.flxMES_Data.TextMatrix(28, 1) = .ABNORMAL_LCD
            Me.flxMES_Data.TextMatrix(29, 1) = .GROUP_ID
            Me.flxMES_Data.TextMatrix(30, 1) = .REPAIR_REWORK_COUNT
            Me.flxMES_Data.TextMatrix(31, 1) = .CARBONIZATION_FLAG
            Me.flxMES_Data.TextMatrix(32, 1) = .CARBONIZATION_GRADE
            Me.flxMES_Data.TextMatrix(33, 1) = .CARBONIZATION_REWORK_COUNT
            Me.flxMES_Data.TextMatrix(34, 1) = .POLARIZER_REWORK_COUNT
            Me.flxMES_Data.TextMatrix(35, 1) = .X_TOTAL_PIXEL
            Me.flxMES_Data.TextMatrix(36, 1) = .Y_TOTAL_PIXEL
            Me.flxMES_Data.TextMatrix(37, 1) = .X_ONE_PIXEL_LENGTH
            Me.flxMES_Data.TextMatrix(38, 1) = .Y_ONE_PIXEL_LENGTH
            Me.flxMES_Data.TextMatrix(39, 1) = .LCD_Q_TAP_LOT_GROUPID
            Me.flxMES_Data.TextMatrix(40, 1) = .SK_FLAG
            Me.flxMES_Data.TextMatrix(41, 1) = .CF_R_DEFECT_CODE
            Me.flxMES_Data.TextMatrix(42, 1) = .ODK_AK_FLAG
            Me.flxMES_Data.TextMatrix(43, 1) = .BPAM_REWORK_FLAG
            Me.flxMES_Data.TextMatrix(44, 1) = .LCD_BRIGHT_DOT_FLAG
            Me.flxMES_Data.TextMatrix(45, 1) = .CF_PS_HEIGHT_ERR_FLAG
            Me.flxMES_Data.TextMatrix(46, 1) = .PI_INSPECTION_NG_FLAG
            Me.flxMES_Data.TextMatrix(47, 1) = .PI_OVER_BAKE_FLAG
            Me.flxMES_Data.TextMatrix(48, 1) = .PI_OVER_Q_TIME_FLAG
            Me.flxMES_Data.TextMatrix(49, 1) = .ODF_OVER_BAKE_FLAG
            Me.flxMES_Data.TextMatrix(50, 1) = .ODF_OVER_Q_TIME_FLAG
            Me.flxMES_Data.TextMatrix(51, 1) = .HVA_OVER_BAKE_FLAG
            Me.flxMES_Data.TextMatrix(52, 1) = .HVA_OVER_Q_TIME_FLAG
            Me.flxMES_Data.TextMatrix(53, 1) = .SEAL_INSPECTION_FLAG
            Me.flxMES_Data.TextMatrix(54, 1) = .ODF_CHECKER_FLAG
            Me.flxMES_Data.TextMatrix(55, 1) = .ODF_DOOR_OPEN_FLAG
            Me.flxMES_Data.TextMatrix(56, 1) = .LOT1_OPERATION_MODE
            Me.flxMES_Data.TextMatrix(57, 1) = .LOT2_OPERATION_MODE
            Me.flxMES_Data.TextMatrix(58, 1) = .PRODUCTID
            Me.flxMES_Data.TextMatrix(59, 1) = .OWNERID
            For intIndex = 1 To 10
                Me.flxMES_Data.TextMatrix(59 + intIndex, 1) = Space(25)
            Next intIndex
        End With
    End If
    
End Sub

Private Sub mnuMES_Data_Send_Click()

    Dim strStatus                       As String
    Dim strCommand                      As String
    
    Dim intPortNo                       As Integer
    Dim strMES_Exist                    As String
    Dim strJOB_Exist                    As String
    Dim strSHARE_Exist                  As String
    Dim strMES_DATA                     As String
    Dim strJOB_DATA                     As String
    Dim strSHARE_DATA                   As String
    Dim strDevice_State                 As String
    
    If frmMain.flxEQ_Information.TextMatrix(4, 1) = "ROC" Then
    
                    'Auto Mode
    Call ENV.Get_Device_Data_by_Name("API", intPortNo, strDevice_State)

                    If intPortNo > 0 Then
                        'QDAC & Time & Panel ID & Owner & Process Number & PFCD & MES Data & Job Data
                        Call EQP.Get_MES_Data_for_API(strMES_Exist, strJOB_Exist, strSHARE_Exist, strMES_DATA, strJOB_DATA, strSHARE_DATA)
                        strCommand = Format(DATE, "YYYYMMDD") & Format(TIME, "HHMMSS") & pubPANEL_INFO.PANELID & pubCST_INFO.OWNER & pubCST_INFO.PROCESS_NUM
                        strCommand = strCommand & pubCST_INFO.PFCD & strMES_DATA & "070" & strJOB_DATA & "207" & strSHARE_DATA
                        Call QUEUE.Put_Send_Command(intPortNo, "QDAC" & strCommand)
                        Call EQP.Set_QDAC_COMMAND("QDAC" & strCommand)
                    End If
    Else
        Call Show_Message("Data error", "QDAC command does not exist.")
    End If
    
End Sub

Private Sub mnuOff_Line_Request_Click()

    Load frmOffLine_Request
    If Me.flxStatus.TextMatrix(1, 1) = cDEVICE_ONLINE Then
        frmOffLine_Request.cmdEQ_OffLine.Enabled = True
    Else
        frmOffLine_Request.cmdEQ_OffLine.Enabled = False
    End If
    If Me.flxStatus.TextMatrix(2, 1) = cDEVICE_ONLINE Then
        frmOffLine_Request.cmdAPI_OffLine.Enabled = True
    Else
        frmOffLine_Request.cmdAPI_OffLine.Enabled = False
    End If
    
    frmOffLine_Request.Show
    
End Sub

Private Sub mnuOn_Line_Click()

    Dim strPath                         As String
    Dim strFileName                     As String
    Dim strCommand                      As String
    
    Dim intFileNum                      As Integer
    
    strPath = App.PATH & "\Env\"
    strFileName = "RONT.txt"
    If Dir(strPath & strFileName, vbNormal) <> "" Then
        intFileNum = FreeFile
        Open strPath & strFileName For Input As intFileNum
        
        While Not EOF(intFileNum)
            Line Input #intFileNum, strCommand
        Wend
        
        Close intFileNum
    End If
    Call QUEUE.Put_Receive_Command(4, strCommand)
'    If strCommand <> "" Then
'        If (Left(strCommand, 1) = cSTX) And (Right(strCommand, 1) = cETX) Then
'            strCommand = Mid(strCommand, 2, Len(strCommand) - 2)
'        End If
'        Call BTST_Sequence(4, strCommand)
'    End If
    
End Sub

Private Sub mnuPG_Online_Request_Click()

    Dim intPortID                       As Integer
    
    Dim strState                        As String
    Dim strDevice_Name                  As String
    Dim strPFCD                         As String
    
    Call ENV.Get_Device_Data_by_Name("PG", intPortID, strState)
    
    If Me.flxEQ_Information.TextMatrix(2, 1) <> "" Then
        If Len(Me.flxEQ_Information.TextMatrix(2, 1)) > 5 Then
            strPFCD = Mid(Me.flxEQ_Information.TextMatrix(2, 1), 3, 5)
        Else
            strPFCD = Space(5)
        End If
    Else
        strPFCD = Space(5)
    End If
        
    If intPortID <> 0 Then
        Call QUEUE.Put_Send_Command(intPortID, "QDRF" & strPFCD)
'        Call QUEUE.Put_Send_Command(intPortID, "QDRF")
    Else
        For intPortID = 1 To 8
            If ENV.Get_Port_Use(intPortID) = True Then
                Call ENV.Get_Device_Data_by_PortID(intPortID, strDevice_Name, strState)
                If strDevice_Name = "" Then
                    If ENV.Get_Port_Use(intPortID) = True Then
                        If Me.MSComm(intPortID - 1).PortOpen = True Then
                            Call QUEUE.Put_Send_Command(intPortID, "QDRF" & strPFCD)
'                            Call QUEUE.Put_Send_Command(intPortID, "QDRF")
                        End If
                    End If
                End If
            End If
        Next intPortID
    End If
    
End Sub

Private Sub mnuPG_Power_Off_Click()

    Dim intPortID                       As Integer
    
    intPortID = EQP.Get_PG_PortID
    Call QUEUE.Put_Send_Command(intPortID, "QPPF")
    
End Sub

Private Sub mnuPG_Power_On_Click()

    Dim intPortID                       As Integer
    
    intPortID = EQP.Get_PG_PortID
    Load frmJudge
    frmJudge.lblCurrent_PTN_Index.Caption = "0"
    
    Call QUEUE.Put_Send_Command(intPortID, "QPPO")
    
End Sub

Private Sub mnuSetting_Modify_Click()

    Dim intPortID                       As Integer
    
    intPortID = EQP.Get_PG_PortID
    Call QUEUE.Put_Send_Command(intPortID, "QSMY" & Mid(frmMain.flxEQ_Information.TextMatrix(2, 1), 3, 5) & frmMain.flxMES_Data.TextMatrix(3, 1))
    
End Sub

Private Sub mnuUser_Regist_Click()

    Select Case ENV.Get_Current_User_Level
    Case "S":
        Load frmUser_Regist
        frmUser_Regist.Show
    Case "E":
        Load frmUser_Regist
        frmUser_Regist.Show
    Case "P":
        Call Show_Message("User confirm", "Can't access user regist menu.")
    Case "T":
        Call Show_Message("User confirm", "Can't access user regist menu.")
    Case Else
        Call Show_Message("User confirm", "Access denied.")
    End Select
    
End Sub

Private Sub MSComm_OnComm(Index As Integer)

    Dim bolReadEnd                      As Boolean
    
    Dim intPortNo                       As Integer
    Dim intResult                       As Integer
    Dim intMsgLen                       As Integer
        
    Dim strInputMsg                     As String
    Dim strDevice_Name                  As String
    Dim strDevice_State                 As String
    Dim strCommand                      As String
    
    intPortNo = Index + 1
    
    Select Case MSComm(Index).CommEvent
        'Error
        Case comBreak:          'Received stop signal
            Call Show_Message("Comm error", "PORT : " & Index & " Reveced stop signal")
        Case comFrame:          'Frame error
            Call Show_Message("Comm error", "PORT : " & Index & " Frame error")
        Case comOverrun:        'Data loss
            Call Show_Message("Comm error", "PORT : " & Index & " Data loss")
        Case comRxOver:         'Receive buffer overflow
            Call Show_Message("Comm error", "PORT : " & Index & " Receive buffer overflow")
        Case comRxParity:       'Parity error
            Call Show_Message("Comm error", "PORT : " & Index & " Parity error")
        Case comTxFull:         'Send buffer full
            Call Show_Message("Comm error", "PORT : " & Index & " Send buffer full")
        Case comDCB:            'Unexpected error within DCB search
            Call Show_Message("Comm error", "PORT : " & Index & " Unexpected error within DBC search")
        'Event
        Case comEvCD:           'CD line change
            Call ENV.Get_Device_Data_by_PortID(intPortNo, strDevice_Name, strDevice_State)
            Select Case strDevice_Name
            Case "API":
                If strDevice_State = cDEVICE_ONLINE Then
                    strCommand = "POFA"
                    Call API_Sequence(intPortNo, strCommand)
'                    strCommand = "API state change to OFF-LINE"
'                    Call Show_Message("Comm notice", strCommand)
                End If
            Case "PG":
                If strDevice_State = cDEVICE_ONLINE Then
                    strCommand = "ROFG"
                    Call PG_Sequence(intPortNo, strCommand)
'                    strCommand = "PG state change to OFF-LINE"
'                    Call Show_Message("Comm notice", strCommand)
                End If
            Case "CATST":
                If strDevice_State = cDEVICE_ONLINE Then
                    strCommand = "ROFT"
                    Call BTST_Sequence(intPortNo, strCommand)
'                    strCommand = "CATST state change to OFF-LINE"
'                    Call Show_Message("Comm notice", strCommand)
                End If
            Case "CALOI":
                If strDevice_State = cDEVICE_ONLINE Then
                    strCommand = "ROFI"
                    Call BLOI_Sequence(intPortNo, strCommand)
'                    strCommand = "CALOI state change to OFF-LINE"
'                    Call Show_Message("Comm notice", strCommand)
                End If
            End Select
        Case comEvCTS:          'CTS line change
        Case comEvDSR:          'DSR line change
        Case comEvRing:         'Detected call
            Call Show_Message("Comm notice", "PORT : " & Index & " Detected call")
        Case comEvReceive:      'Received event
            'If QUEUE.Get_TimeOut(intPortNo) <> 0 Then
'                intResult = QUEUE.Set_TimeOut(intPortNo, 0)
                If intResult = cCOMMAND_QUEUE_INVALID_PORT_NO Then
                    Call SaveLog("MSComm_OnComm", "Invalid port number. Port No : " & intPortNo)
                Else
                    strInputMsg = Me.MSComm(Index).Input
                    While strInputMsg <> ""
                        If Left(strInputMsg, 1) = cSTX Then
                            Call QUEUE.Set_Input_Command(intPortNo, Left(strInputMsg, 1))
                        ElseIf Left(strInputMsg, 1) = cETX Then
                            Call QUEUE.Set_Input_Command(intPortNo, (QUEUE.Get_Input_Command(intPortNo) & Left(strInputMsg, 1)))
                            Select Case QUEUE.Put_Receive_Command(intPortNo, QUEUE.Get_Input_Command(intPortNo))
                            Case cCOMMAND_QUEUE_NORMALCY:
                                Call SaveLog("MSComm_OnComm", "Port No. : " & intPortNo & ", Receive Command : " & QUEUE.Get_Input_Command(intPortNo))
                                Call QUEUE.Set_Input_Command(intPortNo, "")
                            Case cCOMMAND_QUEUE_FULL:
                                    Call SaveLog("MSComm_OnComm", "Port No. : " & intPortNo & ", Received command queue is full.")
                            End Select
                        Else
                            If Asc(Left(strInputMsg, 1)) = 0 Then
                                Call QUEUE.Set_Input_Command(intPortNo, (QUEUE.Get_Input_Command(intPortNo) & " "))
                            Else
                                Call QUEUE.Set_Input_Command(intPortNo, (QUEUE.Get_Input_Command(intPortNo) & Left(strInputMsg, 1)))
                            End If
                        End If
                        If Len(strInputMsg) > 1 Then
                            strInputMsg = Mid(strInputMsg, 2)
                        Else
                            strInputMsg = ""
                        End If
                    Wend
'                    bolReadEnd = False
'                    While bolReadEnd = False
'                        strInputMsg = Me.MSComm(Index).Input
'                        If strInputMsg <> "" Then
'                            Call QUEUE.Set_Input_Command(intPortNo, (QUEUE.Get_Input_Command(intPortNo) & strInputMsg))
'                            If strInputMsg = cETX Then
'                                Select Case QUEUE.Put_Receive_Command(intPortNo, QUEUE.Get_Input_Command(intPortNo))
'                                Case cCOMMAND_QUEUE_NORMALCY:
'                                    Call SaveLog("MSComm_OnComm", "Port No. : " & intPortNo & " Receive Command : " & QUEUE.Get_Input_Command(intPortNo))
'                                    Call QUEUE.Set_Input_Command(intPortNo, "")
'                                Case cCOMMAND_QUEUE_FULL:
'                                        Call SaveLog("MSComm_OnComm", "Port No. : " & intPortNo & " Received command queue is full.")
'                                End Select
'                                bolReadEnd = True
'                            End If
'                        Else
'                            bolReadEnd = True
'                        End If
'                    Wend
                End If
            'End If
        Case comEvSend:         'Send event
        Case comEvEOF:          'End of file
    End Select
                
End Sub

Private Sub RBBU_Click()

    Dim strPath                         As String
    Dim strFileName                     As String
    Dim strCommand                      As String
    
    Dim intFileNum                      As Integer
    
    strPath = App.PATH & "\Env\"
    strFileName = "RBBU.txt"
    If Dir(strPath & strFileName, vbNormal) <> "" Then
        intFileNum = FreeFile
        Open strPath & strFileName For Input As intFileNum
        
        While Not EOF(intFileNum)
            Line Input #intFileNum, strCommand
        Wend
        
        Close intFileNum
    End If
    
    Call QUEUE.Put_Receive_Command(4, strCommand)
'    If strCommand <> "" Then
'        If (Left(strCommand, 1) = cSTX) And (Right(strCommand, 1) = cETX) Then
'            strCommand = Mid(strCommand, 2, Len(strCommand) - 2)
'        End If
'        Call BTST_Sequence(4, strCommand)
'    End If

End Sub

Private Sub tmrBulletin_Board_Timer()

    Dim strMessage              As String
    
    Dim intCount                As Integer
    Dim intIndex                As Integer
    
    If ENV.Get_NOTICE_UPDATE = True Then
        Me.lstMessage.Clear
        intCount = ENV.Get_NOTICE_Count
        If intCount > 0 Then
            For intIndex = 1 To intCount
                strMessage = ENV.Get_NOTICE_MESSAGE_by_Index(intIndex)
                Me.lstMessage.AddItem strMessage
            Next intIndex
        End If
        Call ENV.Set_NOTICE_UPDATE(False)
    Else
        If Me.lstMessage.ListCount > 1 Then
            strMessage = Me.lstMessage.List(0)
            Me.lstMessage.RemoveItem (0)
            Me.lstMessage.AddItem strMessage
        End If
    End If
    
    If Left(Format(TIME, "HHMMSS"), 4) = "0730" Then
        If m_NEW_DATE = False Then
            'New date process program here
            Me.flxJudge_History.Rows = 1
            m_NEW_DATE = True
        End If
    Else
        If m_NEW_DATE = True Then
            m_NEW_DATE = False
        End If
    End If
    
End Sub

Private Sub tmrCommand_Timer()

    Dim intPortNo               As Integer
    Dim intLength               As Integer
    Dim intResult               As Integer
    
    Dim strCommand              As String
    Dim strDevice_Name          As String
    Dim strDevice_State         As String
    Dim strInputMsg             As String
    Dim strMode_State           As String
    Dim strFileName             As String
    Dim strLocalPath            As String
    Dim strLog                  As String
    Dim strDB_Path              As String
    Dim strDB_FileName          As String
    Dim strDB_New_FileName      As String
    
    Dim bolReadEnd              As Boolean
    
    Me.tmrCommand.Enabled = False
    
'==========================================================================================================
'
'  Modify Date : 2011. 12. 19
'  Modify by K.H. KIM
'  Content
'    - If exist log data in queue memory get that data and jump to write function
'
'
'  Start of modify
'
'==========================================================================================================
    
    If QUEUE.Get_Log_Data(strLog) = cCOMMAND_QUEUE_NORMALCY Then
        Call Write_Log(strLog)
    End If
    
'===========================================================================================================
'
'  End of modify
'
'===========================================================================================================
    
    For intPortNo = 1 To 8
        Select Case QUEUE.Get_Send_Command(intPortNo, strCommand)
        Case cCOMMAND_QUEUE_NORMALCY:
            If Me.MSComm(intPortNo - 1).PortOpen = True Then
                If Len(strCommand) >= 6 Then
                    Me.MSComm(intPortNo - 1).Output = strCommand
                    Call SaveLog("tmrCommand_Timer", "Port No. : " & intPortNo & ", Send Command : " & strCommand)
                End If
                Select Case Mid(strCommand, 2, 4)
                Case "YABC":
'==========================================================================================================
'
'  Modify Date : 2011. 12. 28
'  Modify by K.H. KIM
'  Content
'    - Function call position change from modCATST_Sequence to tmrCommand_Timer
'
'
'  Start of modify
'
'==========================================================================================================
                    If Left(frmMain.flxEQ_Information.TextMatrix(5, 1), 5) = "CATST" Then
                        Call Decode_CATST_After_Block_Contact(EQP.Get_RABC_Command)
                    Else
                        Call Decode_CALOI_After_Block_Contact(EQP.Get_RABC_Command)
                    End If
                Case "YBBU":
                    
                    Call Decode_Before_Block_Uncontact(intPortNo)

'==========================================================================================================
'
'  Modify Date : 2011. 12. 13
'  Modify by K.H. KIM
'  Content
'    - Standard information file download condition change for increase processing time
'
'
'  Start of modify
'
'==========================================================================================================
                    Call ENV.Reset_Download_Flag
                    Call Get_File_From_Host("control.csv", "table")
                    Call Read_Control
                
                    If (ENV.Get_Download_Flag = "E") Or (ENV.Get_Download_Flag = "") Then
                        strDB_Path = App.PATH & "\DB\"
                        strDB_FileName = "STANDARD_INFO_Temp.mdb"
                        strDB_New_FileName = "STANDARD_INFO.mdb"
                        
                        If Dir(strDB_Path & strDB_New_FileName, vbNormal) <> "" Then
                            Kill strDB_Path & strDB_New_FileName
                        End If
                        FileCopy strDB_Path & strDB_FileName, strDB_Path & strDB_New_FileName
                        Call Standard_Files_Download
                    End If
                    
'===========================================================================================================
'
'  End of modify
'
'===========================================================================================================
                    Call RANK_OBJ.Set_Current_KEYID("")
                    Call RANK_OBJ.Set_Current_Grade("")
                    Call RANK_OBJ.Set_Select_DEFECTCODE("")
                    
'                    Call Get_File_From_Host("Public notice.txt", "Table")
'                    Call Read_Notice_File
'                    Call Get_File_From_Host("Auto alarm.csv", "Table")
'                    Call Decode_Auto_Alarm
                    
                    If ENV.Get_Data_Change = True Then
                        Call Compress_DB
                    End If
                Case "YBBC":
'==========================================================================================================
'
'  Modify Date : 2011. 12. 13
'  Modify by K.H. KIM
'  Content
'    - Function call position change from modCATST_Sequence to tmrCommand_Timer
'
'
'  Start of modify
'
'==========================================================================================================
                    If Left(frmMain.flxEQ_Information.TextMatrix(5, 1), 5) = "CATST" Then
                        intResult = Decode_CATST_Before_Block_Contact(EQP.Get_RBBC_Command)
                        Select Case intResult
                        Case 0:
                        Case 1:
                            Call QUEUE.Put_Send_Command(intPortNo, "QBAM0005CST_MES_DATA length error.")
                        Case 2:
                            Call QUEUE.Put_Send_Command(intPortNo, "QBAM0006PANEL_MES_DATA length error.")
                        Case 3:
                            Call QUEUE.Put_Send_Command(intPortNo, "QBAM0007JOB_MES_DATA length error.")
                        Case 4:
                            Call QUEUE.Put_Send_Command(intPortNo, "QBAM0008SHARE_MES_DATA length error.")
                        End Select
                    Else
                        intResult = Decode_CALOI_Before_Block_Contact(EQP.Get_RBBC_Command)
                        Select Case intResult
                        Case 0:
                        Case 1:
                            Call QUEUE.Put_Send_Command(intPortNo, "QBAM0005CST_MES_DATA length error.")
                        Case 2:
                            Call QUEUE.Put_Send_Command(intPortNo, "QBAM0006PANEL_MES_DATA length error.")
                        Case 3:
                            Call QUEUE.Put_Send_Command(intPortNo, "QBAM0007JOB_MES_DATA length error.")
                        Case 4:
                            Call QUEUE.Put_Send_Command(intPortNo, "QBAM0008SHARE_MES_DATA length error.")
                        End Select
                    End If
                End Select
'===========================================================================================================
'
'  End of modify
'
'===========================================================================================================
            Else
                'Device check, port open again and send command
            End If
        Case cCOMMAND_QUEUE_EMPTY:
        End Select
        
        Select Case QUEUE.Get_Receive_Command(intPortNo, strCommand)
        Case cCOMMAND_QUEUE_NORMALCY:
            Call SaveLog("tmrCommand_Timer", "Command Queue Message : " & strCommand)
            intLength = Len(strCommand)
            strCommand = Mid(strCommand, 2, intLength - 2)
            Select Case Left(strCommand, 4)
            Case "RONT":            'BTST
                Call ENV.Set_Device_Info(intPortNo, "CATST", "PORT OPEN")
            Case "RONI":            'BLOI
                Call ENV.Set_Device_Info(intPortNo, "CALOI", "PORT OPEN")
            Case "PONA":            'API
                Call ENV.Set_Device_Info(intPortNo, "API", "PORT OPEN")
            Case "PDRF", "RONG":            'PG
                Call ENV.Set_Device_Info(intPortNo, "PG", "PORT OPEN")
                Call EQP.Set_PG_PortID(intPortNo)
            End Select
            
            Call ENV.Get_Device_Data_by_PortID(intPortNo, strDevice_Name, strDevice_State)
            Select Case strDevice_Name
            Case "API":
                Call API_Sequence(intPortNo, strCommand)
            Case "PG":
                Call PG_Sequence(intPortNo, strCommand)
            Case "CATST":
                Call BTST_Sequence(intPortNo, strCommand)
            Case "CALOI":
                Call BLOI_Sequence(intPortNo, strCommand)
            End Select
        Case cCOMMAND_QUEUE_EMPTY:
        End Select
    Next intPortNo
    
    Me.tmrCommand.Enabled = True
    
End Sub

Private Sub tmrLog_Timer()

    If (Format(TIME, "HHMM") = "0730") Or (Format(TIME, "HHMM") = "1930") Then
        If m_PRODUCT_COUNT_RESET = False Then
            m_PRODUCT_COUNT_RESET = True
            Me.flxRUN_Info.TextMatrix(2, 1) = "0"
        End If
    Else
        If m_PRODUCT_COUNT_RESET = True Then
            m_PRODUCT_COUNT_RESET = False
        End If
    End If
    
    If Format(TIME, "HHMM") = "0000" Then
        If ENV.Get_Data_Change = False Then
            Call ENV.Set_Data_Change(True)
        End If
    Else
        If ENV.Get_Data_Change = True Then
            Call ENV.Set_Data_Change(False)
        End If
    End If
    
    Call Reset_Auto_Alarm
    
End Sub

'==========================================================================================================
'
'  Modify Date : 2012. 01. 02
'  Modify by K.H. KIM
'  Content
'    - Reset auto alarm data
'
'
'  Modify Date : 2012. 03. 26
'  Modify by K.H. KIM
'  Content
'    - Reset condition change
'
'==========================================================================================================
Private Sub Reset_Auto_Alarm()

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim typAUTO_ALARM()             As AUTO_ALARM_DATA
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    
    Dim intRecord_Count             As Integer
    Dim intRecord_Index             As Integer
    
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Auto_Alarm.mdb"
    
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
        strQuery = "SELECT * FROM AUTO_ALARM_DATA"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        
        If lstRecord.EOF = False Then
            lstRecord.MoveLast
            intRecord_Count = lstRecord.RecordCount
            ReDim typAUTO_ALARM(intRecord_Count)
            intRecord_Index = 0
            lstRecord.MoveFirst
            While lstRecord.EOF = False
                intRecord_Index = intRecord_Index + 1
                With typAUTO_ALARM(intRecord_Index)
                    .PROCESS_NUM = lstRecord.Fields("PROCESS_NUM")
                    .PFCD = lstRecord.Fields("PFCD")
                    .DEFECT_CODE = lstRecord.Fields("DEFECT_CODE")
                    .Rank = lstRecord.Fields("RANK")
                    .COUNT_TIME = lstRecord.Fields("COUNT_TIME")
                    .Count = lstRecord.Fields("COUNT")
                    .ALARM_TEXT = lstRecord.Fields("ALARM_TEXT")
                    .CURRENT_COUNT = lstRecord.Fields("CURRENT_COUNT")
                    .EXPIRY_DATE = lstRecord.Fields("EXPIRY_DATE")
                    .EXPIRY_TIME = lstRecord.Fields("EXPIRY_TIME")
                End With
                lstRecord.MoveNext
            Wend
        End If
        lstRecord.Close
        
        If intRecord_Count > 0 Then
            For intRecord_Index = 1 To intRecord_Count
                With typAUTO_ALARM(intRecord_Index)
                    strQuery = ""
                    If .EXPIRY_DATE < CLng(Format(DATE, "YYYYMMDD")) Then
                        .EXPIRY_DATE = CLng(Format(DATE + ((1 / 24 / 60) * .COUNT_TIME), "YYYYMMDD"))      '2012.03.26 Modified by K.H.KIM
                        .EXPIRY_TIME = CLng(Format(TIME + ((1 / 24 / 60) * .COUNT_TIME), "HHMMSS"))         '2012.03.26 Modified by K.H.KIM
                        strQuery = "UPDATE AUTO_ALARM_DATA SET "
                        strQuery = strQuery & "CURRENT_COUNT=0, "
                        strQuery = strQuery & "EXPIRY_DATE=" & .EXPIRY_DATE & ", "
                        strQuery = strQuery & "EXPIRY_TIME=" & .EXPIRY_TIME & " WHERE "
                        strQuery = strQuery & "PROCESS_NUM='" & .PROCESS_NUM & "' AND "
                        strQuery = strQuery & "PFCD='" & .PFCD & "' AND "
                        strQuery = strQuery & "DEFECT_CODE='" & .DEFECT_CODE & "'"
'                        strQuery = strQuery & "RANK='" & .RANK & "'"
                    ElseIf (.EXPIRY_DATE = CLng(Format(DATE, "YYYYMMDD"))) And (.EXPIRY_TIME < CLng(Format(TIME, "HHMMSS"))) Then
                        .EXPIRY_DATE = CLng(Format(DATE + ((1 / 24 / 60) * .COUNT_TIME), "YYYYMMDD"))       '2012.03.26 Modified by K.H.KIM
                        .EXPIRY_TIME = CLng(Format(TIME + ((1 / 24 / 60) * .COUNT_TIME), "HHMMSS"))         '2012.03.26 Modified by K.H.KIM
                        strQuery = "UPDATE AUTO_ALARM_DATA SET "
                        strQuery = strQuery & "CURRENT_COUNT=0, "
                        strQuery = strQuery & "EXPIRY_DATE=" & .EXPIRY_DATE & ", "
                        strQuery = strQuery & "EXPIRY_TIME=" & .EXPIRY_TIME & " WHERE "
                        strQuery = strQuery & "PROCESS_NUM='" & .PROCESS_NUM & "' AND "
                        strQuery = strQuery & "PFCD='" & .PFCD & "' AND "
                        strQuery = strQuery & "DEFECT_CODE='" & .DEFECT_CODE & "'"
'                        strQuery = strQuery & "RANK='" & .RANK & "'"
                    Else
                        strQuery = ""
                    End If
                    If strQuery <> "" Then
                        dbMyDB.Execute (strQuery)
                    End If
                End With
            Next intRecord_Index
        End If
        
        dbMyDB.Close
    End If
    
End Sub

Private Sub tmrTimeOut_Timer()

    Dim intPortNo               As Integer
    Dim intMax_Retry_Count      As Integer
    Dim intResult               As Integer
    
    Dim strLastCommand          As String
    
    Dim dblTimeOut              As Double
    Dim dblCurrent_Time         As Double
    
    intMax_Retry_Count = ENV.Get_Max_Retry_Count
    dblCurrent_Time = CDbl(Format(DATE, "YYYYMMDD") & Format(TIME, "HHMMSS"))
    For intPortNo = 1 To 8
        dblTimeOut = QUEUE.Get_TimeOut(intPortNo)
        If dblTimeOut > 0 Then
            If dblTimeOut <= dblCurrent_Time Then
                If QUEUE.Get_Retry_Count(intPortNo) < intMax_Retry_Count Then
                    Call QUEUE.Increase_Retry_Count(intPortNo)
                    intResult = QUEUE.Get_Last_Command(intPortNo, strLastCommand)
                    intResult = QUEUE.Put_Send_Command(intPortNo, strLastCommand)
                Else
                    Call Show_Message("Error occurred", "Communication fail. Last command : " & strLastCommand)
                    Call SaveLog("tmrTimeOut_Timer", "Communication fail. Last command : " & strLastCommand)
                    Call QUEUE.Reset_Last_Command(intPortNo)
                End If
            End If
        End If
    Next intPortNo
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)

On Error GoTo ErrorHandler

    Dim bytResult           As Byte
    
    Select Case UCase(Button)
    Case "HISTORY":
        Load frmJudge_History
        frmJudge_History.Show
    Case "VERSION":
        Load frmVersion
        frmVersion.Show
    Case "PTN INFO.":
    Case "SYSTEM SET.":
        Load frmSystem_Parameter
        frmSystem_Parameter.tabParameter.Tab = 0
        frmSystem_Parameter.Show
    Case "RANK":
        Load frmRank_Interface
        frmRank_Interface.Show
    Case "VIEW LOG":
        frmSystem_Log.Show
    Case "EXIT":
        Unload Me
    End Select
    
    Exit Sub
    
ErrorHandler:
    
End Sub

Private Sub Compress_DB()

    Dim strDB_Path                      As String
    Dim strDB_FileName                  As String
    
    strDB_Path = App.PATH & "\DB\"
    
    strDB_FileName = "Parameter.mdb"
    DBEngine.CompactDatabase strDB_Path & strDB_FileName, strDB_Path & "Parameter_Temp.mdb", dbLangChineseSimplified
    Kill strDB_Path & strDB_FileName
    Name strDB_Path & "Parameter_Temp.mdb" As strDB_Path & strDB_FileName
    
    strDB_FileName = "Result.mdb"
    DBEngine.CompactDatabase strDB_Path & strDB_FileName, strDB_Path & "Result_Temp.mdb", dbLangChineseSimplified
    Kill strDB_Path & strDB_FileName
    Name strDB_Path & "Result_Temp.mdb" As strDB_Path & strDB_FileName
    
End Sub
