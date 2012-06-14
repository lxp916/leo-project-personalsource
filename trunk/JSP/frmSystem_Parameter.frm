VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmSystem_Parameter 
   Caption         =   "System_Parameter"
   ClientHeight    =   6648
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   10356
   LinkTopic       =   "Form1"
   ScaleHeight     =   6648
   ScaleWidth      =   10356
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Frame Frame1 
      Height          =   6645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10335
      Begin VB.TextBox txtEdit 
         Height          =   264
         Left            =   240
         TabIndex        =   63
         Top             =   6120
         Visible         =   0   'False
         Width           =   732
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   525
         Left            =   4440
         TabIndex        =   5
         Top             =   6000
         Width           =   1245
      End
      Begin TabDlg.SSTab tabParameter 
         Height          =   5565
         Left            =   30
         TabIndex        =   1
         Top             =   120
         Width           =   10275
         _ExtentX        =   18119
         _ExtentY        =   9821
         _Version        =   393216
         Tabs            =   6
         Tab             =   5
         TabsPerRow      =   6
         TabHeight       =   520
         TabCaption(0)   =   "Path Parameter"
         TabPicture(0)   =   "frmSystem_Parameter.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame2(0)"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "RS-232 Parameter"
         TabPicture(1)   =   "frmSystem_Parameter.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame2(1)"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Others"
         TabPicture(2)   =   "frmSystem_Parameter.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame3"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "D/F Priority"
         TabPicture(3)   =   "frmSystem_Parameter.frx":0054
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame5"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Alarm/Grade"
         TabPicture(4)   =   "frmSystem_Parameter.frx":0070
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Frame4"
         Tab(4).Control(1)=   "Frame6"
         Tab(4).ControlCount=   2
         TabCaption(5)   =   "Rank Level"
         TabPicture(5)   =   "frmSystem_Parameter.frx":008C
         Tab(5).ControlEnabled=   -1  'True
         Tab(5).Control(0)=   "frameRank(0)"
         Tab(5).Control(0).Enabled=   0   'False
         Tab(5).Control(1)=   "cmdSave"
         Tab(5).Control(1).Enabled=   0   'False
         Tab(5).ControlCount=   2
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save"
            Height          =   585
            Left            =   4800
            TabIndex        =   65
            Top             =   4680
            Width           =   1455
         End
         Begin VB.Frame frameRank 
            Height          =   5112
            Index           =   0
            Left            =   90
            TabIndex        =   61
            Top             =   628
            Width           =   10095
            Begin VB.CommandButton cmdAdd 
               Caption         =   "Add"
               Height          =   585
               Left            =   3200
               TabIndex        =   64
               Top             =   4080
               Width           =   1455
            End
            Begin MSFlexGridLib.MSFlexGrid flxRankLevel 
               Height          =   3725
               Left            =   120
               TabIndex        =   62
               Top             =   240
               Width           =   9885
               _ExtentX        =   17441
               _ExtentY        =   6562
               _Version        =   393216
               Rows            =   1
            End
         End
         Begin VB.Frame Frame6 
            Height          =   3705
            Left            =   -74910
            TabIndex        =   58
            Top             =   2058
            Width           =   10095
            Begin MSFlexGridLib.MSFlexGrid flxAuto_Alarm 
               Height          =   2775
               Left            =   150
               TabIndex        =   59
               Top             =   840
               Width           =   9795
               _ExtentX        =   17272
               _ExtentY        =   4890
               _Version        =   393216
               Rows            =   1
               Cols            =   10
               FixedCols       =   0
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Common defect occurred "
               Height          =   180
               Left            =   570
               TabIndex        =   60
               Top             =   330
               Width           =   2220
            End
         End
         Begin VB.Frame Frame4 
            Height          =   1395
            Left            =   -74910
            TabIndex        =   50
            Top             =   648
            Width           =   10095
            Begin VB.ComboBox cmbChange_Rank 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   14.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Left            =   7080
               TabIndex        =   56
               Top             =   270
               Width           =   885
            End
            Begin VB.TextBox txtChange_Panel_Count 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   14.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   2850
               TabIndex        =   55
               Top             =   810
               Width           =   885
            End
            Begin VB.ComboBox cmbFinal_Rank 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   14.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Left            =   2610
               TabIndex        =   54
               Top             =   270
               Width           =   885
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "(If input 0, not change the panel grade)"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   14.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   3960
               TabIndex        =   57
               Top             =   840
               Width           =   5340
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "change panel grade to"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   14.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   3750
               TabIndex        =   53
               Top             =   300
               Width           =   3060
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "until panel count"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   14.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   300
               TabIndex        =   52
               Top             =   840
               Width           =   2235
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "If final Grade is"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   14.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   270
               TabIndex        =   51
               Top             =   300
               Width           =   2025
            End
         End
         Begin VB.Frame Frame5 
            Height          =   5115
            Left            =   -74910
            TabIndex        =   28
            Top             =   648
            Width           =   10095
            Begin VB.ComboBox cmbDefect_Type 
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
               Index           =   0
               Left            =   1470
               TabIndex        =   38
               Top             =   840
               Width           =   1635
            End
            Begin VB.ComboBox cmbDefect_Type 
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
               Index           =   1
               Left            =   4530
               TabIndex        =   37
               Top             =   840
               Width           =   1635
            End
            Begin VB.ComboBox cmbDefect_Type 
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
               Index           =   2
               Left            =   7530
               TabIndex        =   36
               Top             =   840
               Width           =   1635
            End
            Begin VB.ComboBox cmbDefect_Type 
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
               Index           =   3
               Left            =   1470
               TabIndex        =   35
               Top             =   1800
               Width           =   1635
            End
            Begin VB.ComboBox cmbDefect_Type 
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
               Index           =   4
               Left            =   4530
               TabIndex        =   34
               Top             =   1800
               Width           =   1635
            End
            Begin VB.ComboBox cmbDefect_Type 
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
               Index           =   5
               Left            =   7530
               TabIndex        =   33
               Top             =   1800
               Width           =   1635
            End
            Begin VB.ComboBox cmbDefect_Type 
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
               Index           =   6
               Left            =   1470
               TabIndex        =   32
               Top             =   2760
               Width           =   1635
            End
            Begin VB.ComboBox cmbDefect_Type 
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
               Index           =   7
               Left            =   4530
               TabIndex        =   31
               Top             =   2760
               Width           =   1635
            End
            Begin VB.ComboBox cmbDefect_Type 
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
               Index           =   8
               Left            =   7530
               TabIndex        =   30
               Top             =   2760
               Width           =   1635
            End
            Begin VB.CommandButton cmdDefect_Priority_Add 
               Caption         =   "APPLY"
               Height          =   585
               Left            =   4140
               TabIndex        =   29
               Top             =   4230
               Width           =   1455
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "1st"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   14.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   0
               Left            =   840
               TabIndex        =   47
               Top             =   840
               Width           =   420
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "2nd"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   14.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   1
               Left            =   3900
               TabIndex        =   46
               Top             =   870
               Width           =   495
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "3rd"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   14.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   2
               Left            =   6930
               TabIndex        =   45
               Top             =   870
               Width           =   435
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "4th"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   14.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   3
               Left            =   840
               TabIndex        =   44
               Top             =   1830
               Width           =   420
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "5th"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   14.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   4
               Left            =   3900
               TabIndex        =   43
               Top             =   1830
               Width           =   420
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "6th"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   14.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   5
               Left            =   6900
               TabIndex        =   42
               Top             =   1830
               Width           =   420
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "7th"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   14.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   6
               Left            =   810
               TabIndex        =   41
               Top             =   2790
               Width           =   420
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "8th"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   14.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   7
               Left            =   3870
               TabIndex        =   40
               Top             =   2790
               Width           =   420
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "9th"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   14.4
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   8
               Left            =   6870
               TabIndex        =   39
               Top             =   2790
               Width           =   420
            End
         End
         Begin VB.Frame Frame3 
            Height          =   5115
            Left            =   -74910
            TabIndex        =   7
            Top             =   628
            Width           =   10095
            Begin VB.CheckBox chkFTP_Use 
               Caption         =   "Use FTP Function"
               Height          =   225
               Left            =   1050
               TabIndex        =   49
               Top             =   2580
               Width           =   1935
            End
            Begin VB.CommandButton cmdFTP_Save 
               Caption         =   "Save"
               Height          =   525
               Left            =   4260
               TabIndex        =   48
               Top             =   4410
               Width           =   1245
            End
            Begin VB.TextBox txtDefect_Path 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1860
               TabIndex        =   20
               Top             =   1620
               Width           =   4995
            End
            Begin VB.TextBox txtIndex_Path 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1860
               TabIndex        =   19
               Top             =   1230
               Width           =   4995
            End
            Begin VB.TextBox txtPassword 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   4800
               PasswordChar    =   "*"
               TabIndex        =   18
               Top             =   810
               Width           =   2055
            End
            Begin VB.TextBox txtLogonID 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1410
               TabIndex        =   17
               Top             =   810
               Width           =   1965
            End
            Begin VB.TextBox txtPort_Number 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   6000
               TabIndex        =   16
               Text            =   "5000"
               Top             =   390
               Width           =   825
            End
            Begin VB.TextBox txtIP_Address 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   2580
               TabIndex        =   15
               Text            =   "0.0.0.0"
               Top             =   390
               Width           =   1725
            End
            Begin MSComCtl2.UpDown hscRetry 
               Height          =   315
               Left            =   5400
               TabIndex        =   14
               Top             =   2070
               Width           =   240
               _ExtentX        =   445
               _ExtentY        =   550
               _Version        =   393216
               Enabled         =   -1  'True
            End
            Begin VB.TextBox txtRetry 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4770
               TabIndex        =   13
               Text            =   "0"
               Top             =   2070
               Width           =   645
            End
            Begin MSComCtl2.UpDown hscTimeOut 
               Height          =   315
               Left            =   2490
               TabIndex        =   10
               Top             =   2070
               Width           =   240
               _ExtentX        =   445
               _ExtentY        =   550
               _Version        =   393216
               Max             =   60
               Enabled         =   -1  'True
            End
            Begin VB.TextBox txtTimeOut 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   1860
               TabIndex        =   9
               Text            =   "0"
               Top             =   2070
               Width           =   645
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "time(s)"
               Height          =   180
               Left            =   5730
               TabIndex        =   27
               Top             =   2130
               Width           =   615
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Defect File Path"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   17
               Left            =   540
               TabIndex        =   26
               Top             =   1650
               Width           =   1290
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "HOST DATA Path"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   16
               Left            =   420
               TabIndex        =   25
               Top             =   1290
               Width           =   1260
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Password"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   7
               Left            =   3930
               TabIndex        =   24
               Top             =   870
               Width           =   720
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Log-On ID"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   6
               Left            =   420
               TabIndex        =   23
               Top             =   870
               Width           =   810
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Port Number"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   5
               Left            =   4890
               TabIndex        =   22
               Top             =   450
               Width           =   990
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "File Server IP Address"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   4
               Left            =   420
               TabIndex        =   21
               Top             =   450
               Width           =   1980
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Retry Limit"
               Height          =   180
               Left            =   3810
               TabIndex        =   12
               Top             =   2130
               Width           =   900
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "second"
               Height          =   180
               Left            =   2790
               TabIndex        =   11
               Top             =   2130
               Width           =   630
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Time Out"
               Height          =   180
               Left            =   1020
               TabIndex        =   8
               Top             =   2130
               Width           =   780
            End
         End
         Begin VB.Frame Frame2 
            Height          =   5115
            Index           =   1
            Left            =   -74910
            TabIndex        =   3
            Top             =   628
            Width           =   10095
            Begin MSFlexGridLib.MSFlexGrid flxRS232 
               Height          =   4725
               Left            =   120
               TabIndex        =   6
               Top             =   240
               Width           =   9825
               _ExtentX        =   17336
               _ExtentY        =   8340
               _Version        =   393216
               Rows            =   9
               Cols            =   4
            End
         End
         Begin VB.Frame Frame2 
            Height          =   5115
            Index           =   0
            Left            =   -74910
            TabIndex        =   2
            Top             =   628
            Width           =   10095
            Begin MSFlexGridLib.MSFlexGrid flxPath_Data 
               Height          =   4725
               Left            =   120
               TabIndex        =   4
               Top             =   240
               Width           =   9855
               _ExtentX        =   17378
               _ExtentY        =   8340
               _Version        =   393216
               Rows            =   8
            End
         End
      End
   End
   Begin VB.Menu mnuPort_Use 
      Caption         =   "Port_Use"
      Visible         =   0   'False
      Begin VB.Menu mnuUse_Port 
         Caption         =   "Use Port"
      End
      Begin VB.Menu mnuNot_Use_Port 
         Caption         =   "Not Use Port"
      End
   End
End
Attribute VB_Name = "frmSystem_Parameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SELECT_PATH_GRID_INDEX                       As Integer
 'Leo 2012.05.15 ver.0.9.26 -----Add Rank level
Private Sub cmdAdd_Click()
flxRankLevel.Rows = flxRankLevel.Rows + 1
End Sub

Private Sub cmdClose_Click()
    
    Call Save_Port_Use
    Call Save_Auto_Change_Grade
    
    Unload Me
    
End Sub

Private Sub Save_Port_Use()

    Dim strPath                             As String
    Dim strFileName                         As String
    Dim strTemp                             As String
    
    Dim intFileNum                          As Integer
    Dim intRow                              As Integer
    Dim intPortNo                           As Integer
    
On Error GoTo ErrorHandler

    strPath = App.PATH & "\Env\"
    strFileName = "PORT_USE.cfg"
    intFileNum = FreeFile
    
    Open strPath & strFileName For Output As intFileNum
    
    For intRow = 1 To Me.flxRS232.Rows - 1#
        strTemp = intRow & "=" & UCase(Me.flxRS232.TextMatrix(intRow, 3))
        Print #intFileNum, strTemp
    Next intRow
    
    Close intFileNum
    
    Call ENV.Set_Port_Use
    
    For intPortNo = 0 To 7
        If ENV.Get_Port_Use(intPortNo + 1) = True Then
            If frmMain.MSComm(intPortNo).PortOpen = False Then
                frmMain.MSComm(intPortNo).PortOpen = True
                If frmMain.MSComm(intPortNo).PortOpen = True Then
                    Call SaveLog("cmdClose_click", intPortNo + 1 & " port open")
                End If
            End If
        Else
            If frmMain.MSComm(intPortNo).PortOpen = True Then
                frmMain.MSComm(intPortNo).PortOpen = False
                Call SaveLog("cmdClose_Click", intPortNo & " port close")
            End If
        End If
    Next intPortNo
    
    Exit Sub
    
ErrorHandler:

    If intPortNo >= 0 Then
        Me.flxRS232.TextMatrix(intPortNo + 1, 3) = "FALSE"
        
        strPath = App.PATH & "\Env\"
        strFileName = "PORT_USE.cfg"
        intFileNum = FreeFile
        
        Open strPath & strFileName For Output As intFileNum
        For intRow = 1 To Me.flxRS232.Rows - 1
            strTemp = intRow & "=" & UCase(Me.flxRS232.TextMatrix(intRow, 3))
            Print #intFileNum, strTemp
        Next intRow
        
        Close intFileNum
        
        Call ENV.Set_Port_Use
        
        Call Show_Message("Port state error", "Port" & intPortNo + 1 & " is not available.")
    End If

End Sub

Private Sub Save_Auto_Change_Grade()

    Dim strPath                             As String
    Dim strFileName                         As String
    Dim strTemp                             As String
    
    Dim intFileNum                          As Integer
    Dim intIndex                            As Integer
    Dim intLoopCount                        As Integer
    
    strPath = App.PATH & "\Env\"
    strFileName = "Rank_Interface_Base.cfg"
    intFileNum = FreeFile
    
    Open strPath & strFileName For Output As intFileNum
    
    intLoopCount = Me.cmbChange_Rank.ListCount
    For intIndex = 0 To intLoopCount - 1
        strTemp = Me.cmbChange_Rank.List(intIndex)
        If Trim(strTemp) <> "" Then
            Print #intFileNum, strTemp
        End If
    Next intIndex
    
    Close intFileNum
    
    strFileName = "Auto_Grade.cfg"
    intFileNum = FreeFile
    
    Open strPath & strFileName For Output As intFileNum
    
    strTemp = "FINAL RANK=" & Me.cmbFinal_Rank.Text
    Print #intFileNum, strTemp
    
    strTemp = "CHANGE GRADE=" & Me.cmbChange_Rank.Text
    Print #intFileNum, strTemp
    
    strTemp = "COUNT=" & Me.txtChange_Panel_Count.Text
    Print #intFileNum, strTemp
    
    strTemp = "CURRENT COUNT=" & Me.txtChange_Panel_Count.Text
    Print #intFileNum, strTemp
    frmMain.StatusBar.Panels(2).Text = "Remained Grade Count : " & Me.txtChange_Panel_Count.Text
    
    Close intFileNum
    
End Sub

Private Sub cmdDefect_Priority_Add_Click()

    Dim dbMyDB                              As Database
    
    Dim lstRecord                           As Recordset
    
    Dim strDB_Path                          As String
    Dim strDB_FileName                      As String
    Dim strQuery                            As String
    Dim strDEFECT_TYPE                      As String
    
    Dim intIndex                            As Integer
    
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Parameter.mdb"
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
    
        strQuery = "SELECT * FROM DEFECT_TYPE_PRIORITY"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        
        If lstRecord.EOF = False Then
            lstRecord.Close
            strQuery = "DELETE * FROM DEFECT_TYPE_PRIORITY"
            dbMyDB.Execute (strQuery)
        Else
            lstRecord.Close
        End If
        
        For intIndex = 0 To Me.cmbDefect_Type.Count - 1
            Select Case Me.cmbDefect_Type(intIndex).Text
            Case "POINT D/F":
                strDEFECT_TYPE = "D"
            Case "LINE D/F":
                strDEFECT_TYPE = "L"
            Case "GAP D/F":
                strDEFECT_TYPE = "G"
            Case "MURA D/F":
                strDEFECT_TYPE = "M"
            Case "CF D/F":
                strDEFECT_TYPE = "F"
            Case "POLARIZE D/F":
                strDEFECT_TYPE = "P"
            Case "APPEARANCE D/F":
                strDEFECT_TYPE = "A"
            Case "CELL D/F":
                strDEFECT_TYPE = "C"
            Case "OTHER D/F":
                strDEFECT_TYPE = "O"
            End Select
            
            strQuery = "INSERT INTO DEFECT_TYPE_PRIORITY VALUES ("
            strQuery = strQuery & "'" & strDEFECT_TYPE & "', "
            strQuery = strQuery & intIndex + 1 & ")"
            
            dbMyDB.Execute (strQuery)
        Next intIndex
        
        dbMyDB.Close
    End If
    
'    DBEngine.CompactDatabase strDB_Path & strDB_FileName, strDB_Path & "Rank_Delete_Temp.mdb", dbLangChineseSimplified
'    Kill strDB_Path & strDB_FileName
'    Name strDB_Path & "Rank_Delete_Temp.mdb" As strDB_Path & strDB_FileName
    
End Sub

Private Sub cmdFTP_Save_Click()

    Dim strPath                             As String
    Dim strFileName                         As String
    Dim strTemp                             As String
    Dim strData                             As String
    
    Dim intFileNum                          As Integer
    
    strPath = App.PATH & "\Env\"
    strFileName = "FTP_Parameter.cfg"
    intFileNum = FreeFile
    
    Open strPath & strFileName For Output As intFileNum
    
    strTemp = "IP ADDRESS=" & Me.txtIP_Address.Text
    Print #intFileNum, strTemp
    
    strTemp = "PORT NUMBER=" & Me.txtPort_Number.Text
    Print #intFileNum, strTemp
    
    strTemp = "USERID=" & Me.txtLogonID.Text
    Print #intFileNum, strTemp
    
    strTemp = "PASSWORD=" & Me.txtPassword.Text
    Print #intFileNum, strTemp
    
    strTemp = "HOST DATA PATH=" & Me.txtIndex_Path.Text
    Print #intFileNum, strTemp
    
    strTemp = "DEFECT PATH=" & Me.txtDefect_Path.Text
    Print #intFileNum, strTemp
    
    If Me.chkFTP_Use.Value = vbChecked Then
        strData = "1"
    Else
        strData = "0"
    End If
    strTemp = "USE FTP=" & strData
    Print #intFileNum, strTemp
    
    Close intFileNum
'
'    strFileName = "Common_Parameter.cfg"
'    intFileNum = FreeFile
'
'    Open strPath & strFileName For Output As intFileNum
'
'    strTemp = "JPS NAME=" & Me.txtJPS_Name.Text
'    Call ENV.Set_JPS_Name(Me.txtJPS_Name.Text)
'    Print #intFileNum, strTemp
'
'    Close intFileNum
    
End Sub
 'Leo 2012.05.15 ver.0.9.26 -----Add Rank level
 '==========================================Start
Private Sub cmdSave_Click()
    Dim dbMyDB                              As Database
    
    Dim lstRecord                           As Recordset
    
    Dim strDB_Path                          As String
    Dim strDB_FileName                      As String
    Dim strQuery                            As String
    Dim strDEFECT_TYPE                      As String
    
    Dim RANK()                              As String
    Dim intNewCount                         As Integer
    Dim intIndex                            As Integer
    
    intNewCount = Me.flxRankLevel.Rows - 1
    If intNewCount > 0 Then
        ReDim RANK(intNewCount - 1)
        For intNewCount = 0 To (Me.flxRankLevel.Rows - 1) - 1
            RANK(intNewCount) = Me.flxRankLevel.TextMatrix(intNewCount + 1, 1)
        Next intNewCount
    End If
    
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Parameter.mdb"
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        strQuery = "DELETE * FROM Rank_Level"
        dbMyDB.Execute (strQuery)
        
        If intNewCount > 0 Then
            For intIndex = 0 To intNewCount - 1
                If RANK(intIndex) <> "" Then
                    strQuery = "INSERT INTO Rank_Level (RankCode,OrginalRankCode,RankLevel) VALUES ("
                    strQuery = strQuery & "'" & RANK(intIndex) & "',"
                    strQuery = strQuery & "'" & RANK(intIndex) & "',"
                    strQuery = strQuery & "'" & intIndex + 1 & "')"
                    dbMyDB.Execute (strQuery)
                End If
            Next intIndex
        End If
        
        'Change rank temp data
        strDB_FileName = "RANK_temp.mdb"
        If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
            Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
            For intIndex = 0 To UBound(RankLevel)
                If RankLevel(intIndex) <> "" Then
                    strQuery = "ALTER TABLE  RANK_DATA Drop COLUMN "
                    strQuery = strQuery & "RANK_" & RankLevel(intIndex)
                    dbMyDB.Execute (strQuery)
                End If
            Next intIndex
            
            If intNewCount > 0 Then
                For intIndex = 0 To intNewCount - 1
                    If RANK(intIndex) <> "" Then
                        strQuery = "Alter Table RANK_DATA Add  COLUMN "
                        strQuery = strQuery & "RANK_" & RANK(intIndex) & " Text"
                        dbMyDB.Execute (strQuery)
                    End If
                Next intIndex
            End If
            Call RANK_OBJ.Get_Rank_Levels
        End If
        dbMyDB.Close
    End If
End Sub
 '==========================================End
 'Leo 2012.05.15 ver.0.9.26 -----Add Rank level
Private Sub flxRankLevel_Click()
If flxRankLevel.Rows - 1 = 1 Then
    With txtEdit
          .Top = tabParameter.Top + frameRank(0).Top + flxRankLevel.Top + flxRankLevel.CellTop
          .Left = tabParameter.Left + frameRank(0).Left + flxRankLevel.Left + flxRankLevel.CellLeft
          .Width = flxRankLevel.CellWidth
          .Height = flxRankLevel.CellHeight
          .Text = flxRankLevel.Text
          .SelStart = 0
          .SelLength = Len(txtEdit.Text)
    End With
End If

 With txtEdit
      .Visible = True
      .SetFocus
    End With
End Sub
 'Leo 2012.05.15 ver.0.9.26 -----Add Rank level
Private Sub flxRankLevel_EnterCell()

    With txtEdit
      .Top = tabParameter.Top + frameRank(0).Top + flxRankLevel.Top + flxRankLevel.CellTop
      .Left = tabParameter.Left + frameRank(0).Left + flxRankLevel.Left + flxRankLevel.CellLeft
      .Width = flxRankLevel.CellWidth
      .Height = flxRankLevel.CellHeight
      .Text = flxRankLevel.Text
      .SelStart = 0
      .SelLength = Len(txtEdit.Text)
    End With
End Sub
 'Leo 2012.05.15 ver.0.9.26 -----Add Rank level
Private Sub txtEdit_Change()
    flxRankLevel.Text = txtEdit.Text
End Sub
 'Leo 2012.05.15 ver.0.9.26 -----Add Rank level
Private Sub txtEdit_LostFocus()
    txtEdit.Visible = False
End Sub


Private Sub flxRS232_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim intRow                              As Integer
    
    intRow = Me.flxRS232.Row
    
    If intRow > 0 Then
        If Button = vbRightButton Then
            Me.PopupMenu mnuPort_Use
        End If
    End If
    
End Sub

Private Sub Form_Load()

    Dim intData_Count                       As Integer
    Dim intIndex                            As Integer
    Dim intRow                              As Integer
    
    Dim strPath_Name                        As String
    Dim strPath_Data                        As String
    Dim strDevice_Name                      As String
    Dim strDevice_State                     As String
    
    Dim strMsg                              As String
    
On Error GoTo ErrorHandler
    
    Call Init_Form
    Call Init_Grid
    Call Fill_Grid
    
    intData_Count = ENV.Get_Path_Count
    If intData_Count > 0 Then
        For intIndex = 1 To intData_Count
            Call ENV.Get_Path_Data_by_Index(intIndex, strPath_Name, strPath_Data)
            If strPath_Name <> "" Then
                intRow = Add_Path_Grid_Row(strPath_Name)
                Me.flxPath_Data.TextMatrix(intRow, 1) = strPath_Data
            End If
        Next intIndex
    End If
    
    intData_Count = ENV.Get_Device_Count
    If intData_Count > 0 Then
        For intIndex = 1 To intData_Count
            Call ENV.Get_Device_Data_by_Index(intRow, intRow, strDevice_Name, strDevice_State)
            If strDevice_Name <> "" Then
                With Me.flxRS232
                    .TextMatrix(intRow, 1) = strDevice_Name
                    If strDevice_State = cDEVICE_ENABLE Then
                        .TextMatrix(intRow, 2) = "ENABLE"
                    Else
                        .TextMatrix(intRow, 2) = "DISABLE"
                    End If
                End With
            End If
        Next intIndex
    Else
        With Me.flxRS232
            For intIndex = 1 To .Rows - 1
                .TextMatrix(intIndex, 1) = ""
                .TextMatrix(intIndex, 2) = ""
            Next intIndex
        End With
    End If
    SELECT_PATH_GRID_INDEX = 0
    
    For intRow = 1 To Me.flxRS232.Rows - 1
        If ENV.Get_Port_Use(intRow) = True Then
            Me.flxRS232.TextMatrix(intRow, 3) = "TRUE"
        Else
            Me.flxRS232.TextMatrix(intRow, 3) = "FALSE"
        End If
        Call ENV.Get_Device_Data_by_PortID(intRow, strDevice_Name, strDevice_State)
        If strDevice_Name <> "" Then
            Me.flxRS232.TextMatrix(intRow, 1) = strDevice_Name
            Me.flxRS232.TextMatrix(intRow, 2) = strDevice_State
        End If
    Next intRow
    
    Exit Sub
    
ErrorHandler:

    strMsg = Err.Number & " - " & Err.Description
    Call SaveLog("frmSystem_Parameter_Load", strMsg)
    
End Sub

Private Sub Init_Form()

    Dim dbMyDB                              As Database
    
    Dim lstRecord                           As Recordset
    
    Dim strDB_Path                          As String
    Dim strDB_FileName                      As String
    Dim strFileName                         As String
    Dim strQuery                            As String
    Dim strDEFECT_TYPE                      As String
    Dim strTemp                             As String
    
    Dim intFileNum                          As Integer
    Dim intIndex                            As Integer
    Dim intPos                              As Integer
    
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Parameter.mdb"
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
    
        strQuery = "SELECT * FROM DEFECT_TYPE_PRIORITY"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        
        If lstRecord.EOF = False Then
            lstRecord.MoveFirst
            While lstRecord.EOF = False
                Select Case lstRecord.Fields("DEFECT_TYPE")
                Case "D":
                    strDEFECT_TYPE = "POINT D/F"
                Case "L":
                    strDEFECT_TYPE = "LINE D/F"
                Case "G":
                    strDEFECT_TYPE = "GAP D/F"
                Case "M":
                    strDEFECT_TYPE = "MURA D/F"
                Case "F":
                    strDEFECT_TYPE = "CF D/F"
                Case "P":
                    strDEFECT_TYPE = "POLARIZE D/F"
                Case "A":
                    strDEFECT_TYPE = "APPEARANCE D/F"
                Case "C":
                    strDEFECT_TYPE = "CELL D/F"
                Case "O":
                    strDEFECT_TYPE = "OTHER D/F"
                End Select
                Me.cmbDefect_Type(lstRecord.Fields("DEFECT_PRIORITY") - 1).Text = strDEFECT_TYPE
                lstRecord.MoveNext
            Wend
        End If
        lstRecord.Close
        
        dbMyDB.Close
    End If
    
    strDB_Path = App.PATH & "\Env\"
    strDB_FileName = "FTP_Parameter.cfg"
    
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        intFileNum = FreeFile
        Open strDB_Path & strDB_FileName For Input As intFileNum
        
        While Not EOF(intFileNum)
            Line Input #intFileNum, strTemp
            intPos = InStr(strTemp, "=")
            If intPos > 0 Then
                Select Case Left(strTemp, intPos - 1)
                Case "IP ADDRESS":
                    Me.txtIP_Address.Text = Mid(strTemp, intPos + 1)
                Case "PORT NUMBER":
                    Me.txtPort_Number.Text = CLng(Mid(strTemp, intPos + 1))
                Case "USERID":
                    Me.txtLogonID.Text = Mid(strTemp, intPos + 1)
                Case "PASSWORD":
                    Me.txtPassword.Text = Mid(strTemp, intPos + 1)
                Case "HOST DATA PATH":
                    Me.txtIndex_Path.Text = Mid(strTemp, intPos + 1)
                Case "DEFECT PATH":
                    Me.txtDefect_Path.Text = Mid(strTemp, intPos + 1)
                Case "USE FTP":
                    If Mid(strTemp, intPos + 1) = "1" Then
                        Me.chkFTP_Use.Value = vbChecked
                    Else
                        Me.chkFTP_Use.Value = vbUnchecked
                    End If
                End Select
            End If
        Wend
        
        Close intFileNum
    Else
        Me.txtIP_Address.Text = ""
        Me.txtPort_Number.Text = ""
        Me.txtLogonID.Text = ""
        Me.txtPassword.Text = ""
        Me.txtIndex_Path.Text = ""
        Me.txtDefect_Path.Text = ""
    End If
    
    For intIndex = 0 To Me.cmbDefect_Type.Count - 1
        With Me.cmbDefect_Type(intIndex)
            .AddItem "POINT D/F"
            .AddItem "LINE D/F"
            .AddItem "GAP D/F"
            .AddItem "MURA D/F"
            .AddItem "CF D/F"
            .AddItem "POLARIZE D/F"
            .AddItem "APPEARANCE D/F"
            .AddItem "CELL D/F"
            .AddItem "OTHER D/F"
        End With
    Next intIndex

    With Me
        .cmbFinal_Rank.Clear
       '============Leo 2012.05.22 Add Rank Level Start
       For intIndex = 0 To UBound(RankLevel)
       .cmbFinal_Rank.AddItem RankLevel(intIndex)
       Next intIndex
'        .cmbFinal_Rank.AddItem "Y"
'        .cmbFinal_Rank.AddItem "L"
'        .cmbFinal_Rank.AddItem "K"
'        .cmbFinal_Rank.AddItem "C"
'        .cmbFinal_Rank.AddItem "S"
        '============Leo 2012.05.22 Add Rank Level End
        
        .cmbChange_Rank.Clear
        strFileName = App.PATH & "\Env\Rank_Interface_Base.cfg"
        If Dir(strFileName, vbNormal) <> "" Then
            intFileNum = FreeFile
            Open strFileName For Input As intFileNum
            
            While Not EOF(intFileNum)
                Line Input #intFileNum, strTemp
                If Trim(strTemp) <> "" Then
                    .cmbChange_Rank.AddItem strTemp
                End If
            Wend
            
            Close intFileNum
        End If
        
        strFileName = App.PATH & "\Env\Auto_Grade.cfg"
        If Dir(strFileName, vbNormal) <> "" Then
            intFileNum = FreeFile
            Open strFileName For Input As intFileNum
            
            While Not EOF(intFileNum)
                Line Input #intFileNum, strTemp
                intPos = InStr(strTemp, "=")
                If intPos > 0 Then
                    Select Case Left(strTemp, intPos - 1)
                    Case "FINAL RANK":
                        .cmbFinal_Rank.Text = Mid(strTemp, intPos + 1)
                    Case "CHANGE GRADE":
                        .cmbChange_Rank.Text = Mid(strTemp, intPos + 1)
                    Case "COUNT":
                        .txtChange_Panel_Count.Text = Mid(strTemp, intPos + 1)
                    End Select
                End If
            Wend
            
            Close intFileNum
        Else
            .cmbFinal_Rank.Text = "Y"
            .cmbChange_Rank.Text = ""
            .txtChange_Panel_Count.Text = "0"
        End If
    End With
    
End Sub

Private Sub Init_Grid()

    Dim intRow              As Integer
    Dim intCol              As Integer
    
    With Me.flxPath_Data
        For intRow = 0 To .Rows - 1
            For intCol = 0 To .Cols - 1
                .Row = intRow
                .Col = intCol
                .CellAlignment = flexAlignCenterCenter
                .RowHeight(intRow) = 350
            Next intCol
        Next intRow
        
        .ColWidth(0) = 3400
        .TextMatrix(0, 0) = "PATH NAME"
        .ColWidth(1) = 6000
        .TextMatrix(0, 1) = "PATH"
        .TextMatrix(1, 0) = "EQTYPE"
        .TextMatrix(2, 0) = "PFCD.PID"
        .TextMatrix(3, 0) = "RANK"
        .TextMatrix(4, 0) = "USER"
        .TextMatrix(5, 0) = "PATTERN LIST"
        .TextMatrix(6, 0) = "VERSION"
        .TextMatrix(7, 0) = "TA HISTORY"
    End With
    
    With Me.flxRS232
        For intRow = 0 To .Rows - 1
            For intCol = 0 To .Cols - 1
                .Row = intRow
                .Col = intCol
                .CellAlignment = flexAlignCenterCenter
                .RowHeight(intRow) = 350
            Next intCol
            .TextMatrix(intRow, 0) = intRow
        Next intRow
        
        .ColWidth(0) = 1200
        .TextMatrix(0, 0) = "PORT No."
        .ColWidth(1) = 2000
        .TextMatrix(0, 1) = "DEVICE"
        .ColWidth(2) = 1600
        .TextMatrix(0, 2) = "STATUS"
        .ColWidth(3) = 1500
        .TextMatrix(0, 3) = "USE STATE"
    End With
    
    With Me.flxAuto_Alarm
        For intRow = 0 To .Rows - 1
            For intCol = 0 To .Cols - 1
                .Row = intRow
                .Col = intCol
                .CellAlignment = flexAlignCenterCenter
                .RowHeight(intRow) = 350
            Next intCol
        Next intRow
        
        .ColWidth(0) = 1200
        .TextMatrix(0, 0) = "PROC NO."
        .ColWidth(1) = 1200
        .TextMatrix(0, 1) = "PFCD"
        .ColWidth(2) = 1500
        .TextMatrix(0, 2) = "DEFECT CODE"
        .ColWidth(3) = 1000
        .TextMatrix(0, 3) = "RANK"
        .ColWidth(4) = 1600
        .TextMatrix(0, 4) = "HOW LONG (sec)"
        .ColWidth(5) = 1500
        .TextMatrix(0, 5) = "COUNT"
        .ColWidth(6) = 2000
        .TextMatrix(0, 6) = "ALARM TEXT"
        .ColWidth(7) = 1500
        .TextMatrix(0, 7) = "CURRENT"
        .ColWidth(8) = 1500
        .TextMatrix(0, 8) = "EXPIRY DATE"
        .ColWidth(9) = 1500
        .TextMatrix(0, 9) = "EXPIRY TIME"
    End With
     'Leo 2012.05.15 ver.0.9.26 -----Add Rank level
 '==========================================Start
    With Me.flxRankLevel
        .ColWidth(1) = 1200
        .TextMatrix(0, 1) = "Rank Level."
       
    End With
 '==========================================End
End Sub

Private Sub Fill_Grid()

    Dim dbMyDB                      As Database
    
    Dim lstRecord                   As Recordset
    
    Dim strDB_Path                  As String
    Dim strDB_FileName              As String
    Dim strQuery                    As String
    
    Dim intRow                      As Integer
    
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Parameter.mdb"
    
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
        strQuery = "SELECT * FROM FS_PATH_DATA"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        
        If lstRecord.EOF = False Then
            lstRecord.MoveFirst
            With Me.flxPath_Data
                .TextMatrix(1, 1) = lstRecord.Fields("EQTYPE")
                .TextMatrix(2, 1) = lstRecord.Fields("PFCD_PID")
                .TextMatrix(3, 1) = lstRecord.Fields("RANK")
                .TextMatrix(4, 1) = lstRecord.Fields("USER")
                .TextMatrix(5, 1) = lstRecord.Fields("PATTERN LIST")
                .TextMatrix(6, 1) = lstRecord.Fields("VERSION")
                .TextMatrix(7, 1) = lstRecord.Fields("TA_HISTORY")
            End With
        End If
        lstRecord.Close
        
        dbMyDB.Close
    End If
    
    strDB_FileName = "Auto_Alarm.mdb"
    
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
        strQuery = "SELECT * FROM AUTO_ALARM_DATA"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        
        If lstRecord.EOF = False Then
            lstRecord.MoveFirst
            With Me.flxAuto_Alarm
                While lstRecord.EOF = False
                    intRow = Add_Auto_Alarm_Grid_Row(lstRecord.Fields("PROCESS_NUM"))
                    .TextMatrix(intRow, 1) = lstRecord.Fields("PFCD")
                    .TextMatrix(intRow, 2) = lstRecord.Fields("DEFECT_CODE")
                    .TextMatrix(intRow, 3) = lstRecord.Fields("RANK")
                    .TextMatrix(intRow, 4) = lstRecord.Fields("COUNT_TIME")
                    .TextMatrix(intRow, 5) = lstRecord.Fields("COUNT")
                    .TextMatrix(intRow, 6) = lstRecord.Fields("ALARM_TEXT")
                    .TextMatrix(intRow, 7) = lstRecord.Fields("CURRENT_COUNT")
                    .TextMatrix(intRow, 8) = lstRecord.Fields("EXPIRY_DATE")
                    .TextMatrix(intRow, 9) = lstRecord.Fields("EXPIRY_TIME")
                    
                    lstRecord.MoveNext
                Wend
            End With
        End If
        lstRecord.Close
        
        dbMyDB.Close
    End If
     'Leo 2012.05.15 ver.0.9.26 -----Add Rank level
 '==========================================Start
     strDB_FileName = "Parameter.mdb"
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        Set dbMyDB = Workspaces(0).OpenDatabase(strDB_Path & strDB_FileName)
        
        strQuery = "SELECT * FROM Rank_Level order by RankLevel"
        
        Set lstRecord = dbMyDB.OpenRecordset(strQuery)
        
        If lstRecord.EOF = False Then
            lstRecord.MoveFirst
            With Me.flxRankLevel
                While lstRecord.EOF = False
                    intRow = Add_Rank_Level_Grid_Row(lstRecord.Fields("RankCode"))
                .TextMatrix(intRow, 0) = lstRecord.Fields("RankLevel")
                .TextMatrix(intRow, 1) = lstRecord.Fields("RankCode")
                 lstRecord.MoveNext
                Wend

            End With
        End If
        lstRecord.Close
        
        dbMyDB.Close
    End If
 '==========================================End
End Sub

Private Function Add_Path_Grid_Row(ByVal pPath_Name As String) As Integer

    Dim intRow      As Integer
    Dim intCol      As Integer
    
    With Me.flxPath_Data
        intRow = .Rows
        .AddItem pPath_Name
        .RowHeight(intRow) = 350
        For intCol = 0 To .Cols - 1
            .Row = intRow
            .Col = intCol
            .CellAlignment = flexAlignCenterCenter
        Next intCol
        
        Add_Path_Grid_Row = .Rows - 1
    End With

End Function

Private Function Add_Auto_Alarm_Grid_Row(ByVal pPFCD As String) As Integer

    Dim intRow      As Integer
    Dim intCol      As Integer
    
    With Me.flxAuto_Alarm
        intRow = .Rows
        .AddItem pPFCD
        .RowHeight(intRow) = 350
        For intCol = 0 To .Cols - 2
            .Row = intRow
            .Col = intCol
            .CellAlignment = flexAlignCenterCenter
        Next intCol
        .Col = .Cols - 1
        .CellAlignment = flexAlignLeftCenter
        
        Add_Auto_Alarm_Grid_Row = .Rows - 1
    End With

End Function
 'Leo 2012.05.15 ver.0.9.26 -----Add Rank level
 '==========================================Start
Private Function Add_Rank_Level_Grid_Row(ByVal pPFCD As String) As Integer

    Dim intRow      As Integer
    Dim intCol      As Integer
    
    With Me.flxRankLevel
        intRow = .Rows
        .AddItem pPFCD
        .RowHeight(intRow) = 350
        For intCol = 0 To .Cols - 1
            .Row = intRow
            .Col = intCol
            .CellAlignment = flexAlignCenterCenter
        Next intCol
        .Col = .Cols - 1
        .CellAlignment = flexAlignLeftCenter
        
        Add_Rank_Level_Grid_Row = .Rows - 1
    End With
 'Leo 2012.05.15 ver.0.9.26 -----Add Rank level
 '==========================================End
End Function

Private Sub Form_Unload(Cancel As Integer)

    Dim strDB_Path                          As String
    Dim strDB_FileName                      As String
    
    strDB_Path = App.PATH & "\DB\"
    strDB_FileName = "Parameter.mdb"
    If Dir(strDB_Path & strDB_FileName, vbNormal) <> "" Then
        DBEngine.CompactDatabase strDB_Path & strDB_FileName, strDB_Path & "Parameter_Temp.mdb", dbLangChineseSimplified
        Kill strDB_Path & strDB_FileName
        Name strDB_Path & "Parameter_Temp.mdb" As strDB_Path & strDB_FileName
    End If
    
    Call ENV.Init_Class

End Sub

Private Sub hscRetry_Change()

    Me.txtRetry.Text = Me.hscRetry.Value
    
End Sub

Private Sub hscTimeOut_Change()

    Me.txtTimeOut.Text = Me.hscTimeOut.Value
    
End Sub

Private Sub mnuNot_Use_Port_Click()

    Dim intRow                              As Integer
    
    intRow = Me.flxRS232.Row
    If intRow > 0 Then
        Me.flxRS232.TextMatrix(intRow, 3) = "FALSE"
    End If
    
End Sub

Private Sub mnuPort_Use_Click()

    Dim intRow                              As Integer
    
    intRow = Me.flxRS232.Row
    If intRow > 0 Then
        Me.flxRS232.TextMatrix(intRow, 3) = "TRUE"
    End If
    
End Sub

Private Sub txtRetry_Change()

    Dim intRetry                            As Integer
    
    If IsNumeric(Me.txtRetry.Text) = True Then
        intRetry = CInt(Me.txtRetry.Text)
        If (0 <= intRetry) And (intRetry < 11) Then
            Me.hscRetry.Value = intRetry
        Else
            Call Show_Message("Invalid value", "Retry limit must setted between 0 in 10")
            Me.txtRetry.Text = Me.hscRetry.Value
        End If
    Else
        Me.txtRetry.Text = ""
    End If
    
End Sub

Private Sub txtTimeOut_Change()

    Dim intTimeOut                          As Integer
    
    If IsNumeric(Me.txtTimeOut.Text) = True Then
        intTimeOut = CInt(Me.txtTimeOut.Text)
        If (0 < intTimeOut) And (intTimeOut < 61) Then
            Me.hscTimeOut.Value = intTimeOut
        Else
            Call Show_Message("Invalid value", "Timeout must setted between 1 in 60")
            Me.txtTimeOut.Text = Me.hscTimeOut.Value
        End If
    Else
        Me.txtTimeOut.Text = ""
    End If
    
End Sub
