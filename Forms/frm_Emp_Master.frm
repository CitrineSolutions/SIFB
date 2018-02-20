VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Object = "{C3A136DA-B937-492B-968D-A437638F7AAB}#1.0#0"; "CS_DateControl.ocx"
Begin VB.Form frm_Emp_Master 
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11055
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   900
      Left            =   120
      TabIndex        =   95
      Top             =   -60
      Width           =   14310
      Begin VB.Frame Frm_EmpStatus 
         Caption         =   "Employee Status"
         ForeColor       =   &H00C00000&
         Height          =   720
         Left            =   6045
         TabIndex        =   102
         Top             =   105
         Width           =   2925
         Begin VB.OptionButton Opt_Left 
            Caption         =   "Left"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1785
            TabIndex        =   104
            Top             =   270
            Width           =   885
         End
         Begin VB.OptionButton Opt_Active 
            Caption         =   "Active"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   390
            TabIndex        =   103
            Top             =   270
            Width           =   885
         End
      End
      Begin VB.CommandButton Btn_View 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   1005
         Picture         =   "frm_Emp_Master.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   101
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Save 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   30
         Picture         =   "frm_Emp_Master.frx":36B0
         Style           =   1  'Graphical
         TabIndex        =   100
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Exit 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   4470
         Picture         =   "frm_Emp_Master.frx":6D39
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Clear 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   1725
         Picture         =   "frm_Emp_Master.frx":A399
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Delete 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   2715
         Picture         =   "frm_Emp_Master.frx":DA09
         Style           =   1  'Graphical
         TabIndex        =   97
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Print 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   3435
         Picture         =   "frm_Emp_Master.frx":110B3
         Style           =   1  'Graphical
         TabIndex        =   96
         Top             =   120
         Width           =   700
      End
   End
   Begin VB.ComboBox Cmb_Title 
      Height          =   315
      Left            =   3135
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1305
      Width           =   1410
   End
   Begin VB.TextBox Txtc_EmpNo 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   1185
      MaxLength       =   7
      TabIndex        =   0
      Top             =   1305
      Width           =   1050
   End
   Begin VB.TextBox Txtc_Name 
      Height          =   300
      Left            =   5850
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1305
      Width           =   3165
   End
   Begin VB.TextBox Txtc_OtherName 
      Height          =   300
      Left            =   10920
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1305
      Width           =   3240
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6390
      Left            =   120
      TabIndex        =   56
      Top             =   1740
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   11271
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Personal Info"
      TabPicture(0)   =   "frm_Emp_Master.frx":146A1
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra_crddeb"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Department "
      TabPicture(1)   =   "frm_Emp_Master.frx":146BD
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Leave"
      TabPicture(2)   =   "frm_Emp_Master.frx":146D9
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Va_Leave"
      Tab(2).Control(1)=   "Frm_SalaryType"
      Tab(2).Control(2)=   "Va_Salary"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Additional Info"
      TabPicture(3)   =   "frm_Emp_Master.frx":146F5
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame6"
      Tab(3).Control(1)=   "Frame5"
      Tab(3).Control(2)=   "Frame3"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "Remarks"
      TabPicture(4)   =   "frm_Emp_Master.frx":14711
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Va_Remarks"
      Tab(4).ControlCount=   1
      Begin VB.Frame Frame6 
         Caption         =   "Passport / Work Permit Details"
         ForeColor       =   &H00C00000&
         Height          =   1365
         Left            =   -74850
         TabIndex        =   108
         Top             =   435
         Width           =   13890
         Begin VB.TextBox Txtc_WrkPermit 
            Height          =   300
            Left            =   1605
            MaxLength       =   50
            TabIndex        =   44
            Top             =   615
            Width           =   2040
         End
         Begin VB.TextBox Txtc_PassNo 
            Height          =   300
            Left            =   1605
            MaxLength       =   50
            TabIndex        =   43
            Top             =   315
            Width           =   2040
         End
         Begin VB.TextBox Txtc_MinRefNo 
            Height          =   300
            Left            =   10800
            MaxLength       =   50
            TabIndex        =   49
            Top             =   315
            Width           =   2730
         End
         Begin VB.TextBox Txtc_Flightno 
            Height          =   300
            Left            =   1605
            MaxLength       =   50
            TabIndex        =   45
            Top             =   915
            Width           =   2040
         End
         Begin VB.TextBox Txtc_GuaranteeNo 
            Height          =   300
            Left            =   10800
            MaxLength       =   50
            TabIndex        =   50
            Top             =   615
            Width           =   2730
         End
         Begin VB.TextBox Txtc_GuaranteePassport 
            Height          =   300
            Left            =   10800
            MaxLength       =   50
            TabIndex        =   51
            Top             =   915
            Width           =   2730
         End
         Begin CS_DateControl.DateControl Txtd_PassExpDt 
            Height          =   345
            Left            =   5955
            TabIndex        =   46
            Top             =   315
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   609
         End
         Begin CS_DateControl.DateControl Txtd_WrkPrmtExpiry 
            Height          =   345
            Left            =   5955
            TabIndex        =   47
            Top             =   615
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   609
         End
         Begin CS_DateControl.DateControl Txtd_FlightDate 
            Height          =   345
            Left            =   5955
            TabIndex        =   48
            Top             =   915
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   609
         End
         Begin VB.Label Label49 
            Alignment       =   1  'Right Justify
            Caption         =   "Expiry Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4590
            TabIndex        =   117
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            Caption         =   "Passport No."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   300
            TabIndex        =   116
            Top             =   330
            Width           =   1215
         End
         Begin VB.Label Label56 
            Alignment       =   1  'Right Justify
            Caption         =   "Ministry Ref No."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   9270
            TabIndex        =   115
            Top             =   345
            Width           =   1470
         End
         Begin VB.Label Label64 
            Alignment       =   1  'Right Justify
            Caption         =   "Flight No."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   420
            TabIndex        =   114
            Top             =   930
            Width           =   1095
         End
         Begin VB.Label Label63 
            Alignment       =   1  'Right Justify
            Caption         =   "Arrival Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4590
            TabIndex        =   113
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label46 
            Alignment       =   1  'Right Justify
            Caption         =   "Expiry Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4590
            TabIndex        =   112
            Top             =   660
            Width           =   1215
         End
         Begin VB.Label Label45 
            Alignment       =   1  'Right Justify
            Caption         =   "Permit No."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   420
            TabIndex        =   111
            Top             =   630
            Width           =   1095
         End
         Begin VB.Label Label65 
            Alignment       =   1  'Right Justify
            Caption         =   "Labour Guarantee"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8910
            TabIndex        =   110
            Top             =   645
            Width           =   1830
         End
         Begin VB.Label Label66 
            Alignment       =   1  'Right Justify
            Caption         =   "Passport Guarantee"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8910
            TabIndex        =   109
            Top             =   945
            Width           =   1830
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Experiance Details"
         ForeColor       =   &H00C00000&
         Height          =   4275
         Left            =   -67860
         TabIndex        =   107
         Top             =   1875
         Width           =   6885
         Begin FPUSpreadADO.fpSpread Va_Exp 
            Height          =   3930
            Left            =   0
            TabIndex        =   53
            Top             =   330
            Width           =   6870
            _Version        =   458752
            _ExtentX        =   12118
            _ExtentY        =   6932
            _StockProps     =   64
            AutoClipboard   =   0   'False
            DisplayRowHeaders=   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   4
            MaxRows         =   25
            ProcessTab      =   -1  'True
            SpreadDesigner  =   "frm_Emp_Master.frx":1472D
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Education Details"
         ForeColor       =   &H00C00000&
         Height          =   4260
         Left            =   -74850
         TabIndex        =   106
         Top             =   1875
         Width           =   6885
         Begin FPUSpreadADO.fpSpread Va_Edu 
            Height          =   3915
            Left            =   0
            TabIndex        =   52
            Top             =   330
            Width           =   6870
            _Version        =   458752
            _ExtentX        =   12118
            _ExtentY        =   6906
            _StockProps     =   64
            AutoClipboard   =   0   'False
            DisplayRowHeaders=   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   3
            MaxRows         =   25
            ProcessTab      =   -1  'True
            SpreadDesigner  =   "frm_Emp_Master.frx":14AD9
         End
      End
      Begin VB.Frame Frame1 
         ForeColor       =   &H8000000B&
         Height          =   5970
         Left            =   -74910
         TabIndex        =   67
         Top             =   345
         Width           =   14025
         Begin VB.CheckBox Chk_Disabled 
            Caption         =   "Is Disabled"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   11985
            TabIndex        =   37
            Top             =   1680
            Width           =   1785
         End
         Begin VB.TextBox Txtc_EmpType 
            Height          =   300
            Left            =   1620
            MaxLength       =   25
            TabIndex        =   30
            Top             =   1275
            Width           =   1620
         End
         Begin VB.Frame Frm_Salary 
            Caption         =   "Frame7"
            Height          =   3840
            Left            =   105
            TabIndex        =   118
            Top             =   2070
            Width           =   13830
            Begin VB.CheckBox Chk_MealAllow 
               Caption         =   "Meal Allowance"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   10035
               TabIndex        =   160
               Top             =   285
               Width           =   1725
            End
            Begin VB.ComboBox Cmb_PayType 
               Height          =   315
               Left            =   1530
               Style           =   2  'Dropdown List
               TabIndex        =   139
               Top             =   2865
               Width           =   2535
            End
            Begin VB.TextBox Txtc_AcctNo 
               Height          =   300
               Left            =   1530
               MaxLength       =   25
               TabIndex        =   138
               Top             =   3180
               Width           =   2535
            End
            Begin VB.TextBox Txtc_BankCode 
               Height          =   300
               Left            =   10845
               MaxLength       =   25
               TabIndex        =   137
               Top             =   2865
               Width           =   2535
            End
            Begin VB.TextBox Txtc_PAD 
               Height          =   300
               Left            =   10845
               MaxLength       =   50
               TabIndex        =   136
               Top             =   1875
               Width           =   2535
            End
            Begin VB.ComboBox Cmb_TpFlag 
               Height          =   315
               Left            =   1530
               Style           =   2  'Dropdown List
               TabIndex        =   135
               Top             =   1905
               Width           =   2535
            End
            Begin VB.TextBox Txtc_Town 
               Height          =   300
               Left            =   1530
               MaxLength       =   50
               TabIndex        =   134
               Top             =   2205
               Width           =   2535
            End
            Begin VB.TextBox Txtc_Road 
               Height          =   300
               Left            =   6210
               MaxLength       =   50
               TabIndex        =   133
               Top             =   2160
               Width           =   2535
            End
            Begin VB.TextBox Txtc_Bank 
               Height          =   300
               Left            =   6210
               TabIndex        =   132
               Top             =   2865
               Width           =   2535
            End
            Begin VB.TextBox Txtc_ITNo 
               Height          =   300
               Left            =   1530
               MaxLength       =   10
               TabIndex        =   131
               Top             =   915
               Width           =   1620
            End
            Begin VB.TextBox Txtc_TPMode 
               Height          =   300
               Left            =   6210
               MaxLength       =   50
               TabIndex        =   130
               Top             =   1860
               Width           =   2535
            End
            Begin VB.TextBox Txtn_EduAmt 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   5040
               TabIndex        =   129
               Top             =   1230
               Width           =   1635
            End
            Begin VB.TextBox Txtn_IntAmt 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   1530
               TabIndex        =   128
               Top             =   1215
               Width           =   1620
            End
            Begin VB.TextBox Txtn_PreAmt 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   11880
               TabIndex        =   127
               Top             =   915
               Width           =   1590
            End
            Begin VB.TextBox Txtn_OthAmt 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   8415
               TabIndex        =   126
               Top             =   1215
               Width           =   1590
            End
            Begin VB.TextBox Txtn_EdfAmt 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   0
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
               Height          =   300
               Left            =   8415
               TabIndex        =   125
               Top             =   915
               Width           =   1590
            End
            Begin VB.ComboBox Cmb_EDFCategory 
               Height          =   315
               Left            =   5040
               Style           =   2  'Dropdown List
               TabIndex        =   124
               Top             =   915
               Width           =   1635
            End
            Begin VB.CheckBox Chk_NPFDeduct 
               Caption         =   "NPF Deduct"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   8400
               TabIndex        =   123
               Top             =   285
               Width           =   1470
            End
            Begin VB.CheckBox Chk_PayeRelief 
               Caption         =   "PAYE Relief (50%)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   11880
               TabIndex        =   122
               Top             =   285
               Width           =   1845
            End
            Begin VB.TextBox Txtn_MLDays 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               Height          =   300
               Left            =   6705
               MaxLength       =   3
               TabIndex        =   121
               Top             =   255
               Width           =   645
            End
            Begin VB.CheckBox Chk_NoBonus 
               Caption         =   "No EOY Bonus"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   6210
               TabIndex        =   120
               Top             =   3195
               Width           =   1470
            End
            Begin VB.CheckBox Chk_NoPay 
               Caption         =   "On Leave Without Pay"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   10845
               TabIndex        =   119
               Top             =   3195
               Width           =   2325
            End
            Begin CS_DateControl.DateControl Dtp_MLFrom 
               Height          =   345
               Left            =   1530
               TabIndex        =   140
               Top             =   270
               Width           =   1620
               _ExtentX        =   2858
               _ExtentY        =   609
            End
            Begin CS_DateControl.DateControl Dtp_MLTo 
               Height          =   345
               Left            =   5040
               TabIndex        =   141
               Top             =   255
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   609
            End
            Begin VB.Line Line4 
               BorderColor     =   &H80000005&
               X1              =   30
               X2              =   13715
               Y1              =   2685
               Y2              =   2685
            End
            Begin VB.Line Line5 
               BorderColor     =   &H80000005&
               X1              =   15
               X2              =   13700
               Y1              =   1695
               Y2              =   1695
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "PAYE No."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   570
               TabIndex        =   159
               Top             =   960
               Width           =   855
            End
            Begin VB.Label Label26 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Payment Type"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   255
               TabIndex        =   158
               Top             =   2910
               Width           =   1170
            End
            Begin VB.Label Label27 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Bank Name"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   210
               Left            =   5085
               TabIndex        =   157
               Top             =   2910
               Width           =   1050
            End
            Begin VB.Label Label28 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Account No."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   450
               TabIndex        =   156
               Top             =   3225
               Width           =   975
            End
            Begin VB.Label Label36 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Bank Code"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   9870
               TabIndex        =   155
               Top             =   2910
               Width           =   885
            End
            Begin VB.Label Label43 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Road"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   210
               Left            =   5595
               TabIndex        =   154
               Top             =   2205
               Width           =   540
            End
            Begin VB.Label Label52 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "PAD Point"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   210
               Left            =   9960
               TabIndex        =   153
               Top             =   1920
               Width           =   795
            End
            Begin VB.Label Label59 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Using Transport"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   90
               TabIndex        =   152
               Top             =   1950
               Width           =   1335
            End
            Begin VB.Label Label60 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Town / Village"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   210
               Left            =   270
               TabIndex        =   151
               Top             =   2250
               Width           =   1155
            End
            Begin VB.Label Label61 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Mod of Transport"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   210
               Left            =   4560
               TabIndex        =   150
               Top             =   1905
               Width           =   1575
            End
            Begin VB.Label Lbl_EdfCat 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "EDF Catefory"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Left            =   3840
               TabIndex        =   149
               Top             =   960
               Width           =   1035
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "EDF Amount"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   7320
               TabIndex        =   148
               Top             =   960
               Width           =   1035
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Premium"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   11025
               TabIndex        =   147
               Top             =   960
               Width           =   765
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Others"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   7755
               TabIndex        =   146
               Top             =   1260
               Width           =   570
            End
            Begin VB.Label Label39 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Interest"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   765
               TabIndex        =   145
               Top             =   1260
               Width           =   660
            End
            Begin VB.Label Label40 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Education"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   4080
               TabIndex        =   144
               Top             =   1275
               Width           =   795
            End
            Begin VB.Label Label32 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "ML Date From"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   285
               TabIndex        =   143
               Top             =   300
               Width           =   1140
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "ML Date To"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   3960
               TabIndex        =   142
               Top             =   300
               Width           =   915
            End
         End
         Begin VB.ComboBox Cmb_Company 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   330
            Width           =   2535
         End
         Begin VB.CheckBox Chk_ClockCard 
            Caption         =   "Using Clock Card"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   8490
            TabIndex        =   36
            Top             =   1665
            Width           =   2325
         End
         Begin VB.TextBox Txtc_ShiftCode 
            Height          =   300
            Left            =   1620
            TabIndex        =   34
            Top             =   1590
            Width           =   1620
         End
         Begin VB.TextBox Txtc_Line 
            Height          =   300
            Left            =   5130
            MaxLength       =   50
            TabIndex        =   35
            Top             =   1590
            Width           =   1620
         End
         Begin VB.ComboBox Cmb_DayWork 
            Height          =   315
            Left            =   11970
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   1275
            Width           =   1620
         End
         Begin VB.ComboBox Cmb_StaffType 
            Height          =   315
            Left            =   5130
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   1275
            Width           =   1620
         End
         Begin VB.ComboBox Cmb_SkillSet 
            Height          =   315
            Left            =   8505
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   1275
            Width           =   1620
         End
         Begin VB.CommandButton Btn_AddCombo 
            Height          =   300
            Index           =   4
            Left            =   4260
            Picture         =   "frm_Emp_Master.frx":14E0D
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Combo Add"
            Top             =   652
            Width           =   325
         End
         Begin VB.CommandButton Btn_AddCombo 
            Height          =   300
            Index           =   3
            Left            =   13500
            Picture         =   "frm_Emp_Master.frx":1619A
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Combo Add"
            Top             =   330
            Width           =   325
         End
         Begin VB.CommandButton Btn_AddCombo 
            Height          =   300
            Index           =   2
            Left            =   8805
            Picture         =   "frm_Emp_Master.frx":17527
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Combo Add"
            Top             =   360
            Width           =   325
         End
         Begin VB.ComboBox Cmb_Branch 
            Height          =   315
            Left            =   6300
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   330
            Width           =   2535
         End
         Begin VB.ComboBox Cmb_Desig 
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   645
            Width           =   2535
         End
         Begin VB.ComboBox Cmb_Dept 
            Height          =   315
            Left            =   10935
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   330
            Width           =   2535
         End
         Begin CS_DateControl.DateControl Dtp_DOL 
            Height          =   345
            Left            =   10935
            TabIndex        =   29
            Top             =   645
            Width           =   2520
            _ExtentX        =   4445
            _ExtentY        =   609
         End
         Begin CS_DateControl.DateControl Dtp_DOJ 
            Height          =   345
            Left            =   6315
            TabIndex        =   28
            Top             =   645
            Width           =   2520
            _ExtentX        =   4445
            _ExtentY        =   609
         End
         Begin VB.Label Label48 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Staff Type"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4140
            TabIndex        =   90
            Top             =   1320
            Width           =   825
         End
         Begin VB.Label Label47 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Emp Type"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   705
            TabIndex        =   89
            Top             =   1320
            Width           =   810
         End
         Begin VB.Line Line6 
            BorderColor     =   &H80000005&
            X1              =   45
            X2              =   14000
            Y1              =   2025
            Y2              =   2025
         End
         Begin VB.Label Label44 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Shift"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   1140
            TabIndex        =   79
            Top             =   1635
            Width           =   375
         End
         Begin VB.Label Label42 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Line"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   210
            Left            =   4470
            TabIndex        =   78
            Top             =   1635
            Width           =   495
         End
         Begin VB.Label Label38 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Skill Set"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   7740
            TabIndex        =   77
            Top             =   1320
            Width           =   660
         End
         Begin VB.Label Label37 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Branch"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   5520
            TabIndex        =   76
            Top             =   360
            Width           =   705
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Days Work"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   10845
            TabIndex        =   73
            Top             =   1320
            Width           =   1005
         End
         Begin VB.Line Line3 
            BorderColor     =   &H80000005&
            X1              =   0
            X2              =   14000
            Y1              =   1110
            Y2              =   1110
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Employee At"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   480
            TabIndex        =   72
            Top             =   375
            Width           =   1035
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Department"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   9870
            TabIndex        =   71
            Top             =   375
            Width           =   975
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Designation"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   540
            TabIndex        =   70
            Top             =   697
            Width           =   975
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Date of Join"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   5265
            TabIndex        =   69
            Top             =   697
            Width           =   960
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Date Left"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   10110
            TabIndex        =   68
            Top             =   697
            Width           =   735
         End
      End
      Begin VB.Frame fra_crddeb 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5970
         Left            =   90
         TabIndex        =   59
         Top             =   330
         Width           =   14010
         Begin VB.CommandButton Btn_AddCombo 
            Height          =   300
            Index           =   1
            Left            =   9930
            Picture         =   "frm_Emp_Master.frx":188B4
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Combo Add"
            Top             =   615
            Width           =   325
         End
         Begin VB.TextBox Txtc_NICNo 
            Height          =   315
            Left            =   4890
            MaxLength       =   12
            TabIndex        =   10
            Top             =   615
            Width           =   1665
         End
         Begin VB.ComboBox Cmb_MatStatus 
            Height          =   315
            Left            =   8235
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   615
            Width           =   1680
         End
         Begin VB.CommandButton Btn_AddCombo 
            Height          =   300
            Index           =   0
            Left            =   9930
            Picture         =   "frm_Emp_Master.frx":19C41
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Combo Add"
            Top             =   285
            Width           =   325
         End
         Begin VB.TextBox Txtc_Qualify 
            Height          =   300
            Left            =   8235
            MaxLength       =   50
            TabIndex        =   18
            Top             =   2175
            Width           =   5520
         End
         Begin VB.TextBox Txtc_Specialist 
            Height          =   300
            Left            =   8235
            MaxLength       =   50
            TabIndex        =   19
            Top             =   2475
            Width           =   5520
         End
         Begin VB.TextBox Txtc_FamilyDetail 
            Height          =   810
            Left            =   8235
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   17
            Top             =   1365
            Width           =   5520
         End
         Begin VB.TextBox Txtc_Add 
            Height          =   825
            Left            =   1710
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            Top             =   1350
            Width           =   4860
         End
         Begin VB.ComboBox Cmb_Nationality 
            Height          =   315
            Left            =   8225
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   300
            Width           =   1680
         End
         Begin VB.ComboBox Cmb_Expatriate 
            Height          =   315
            ItemData        =   "frm_Emp_Master.frx":1AFCE
            Left            =   12060
            List            =   "frm_Emp_Master.frx":1AFD0
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   300
            Width           =   1680
         End
         Begin VB.TextBox Txtc_AddInfo 
            Height          =   1605
            Left            =   1695
            MaxLength       =   200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   20
            Top             =   3150
            Width           =   12045
         End
         Begin VB.TextBox Txtc_BloodGroup 
            Height          =   300
            Left            =   12060
            MaxLength       =   10
            TabIndex        =   13
            Top             =   615
            Width           =   1680
         End
         Begin VB.TextBox Txtc_EMail 
            Height          =   300
            Left            =   1695
            MaxLength       =   50
            TabIndex        =   16
            Top             =   2490
            Width           =   4860
         End
         Begin VB.TextBox Txtc_PerContNo 
            Height          =   300
            Left            =   1695
            MaxLength       =   50
            TabIndex        =   15
            Top             =   2175
            Width           =   4860
         End
         Begin VB.TextBox Txtc_SocSecNo 
            Height          =   300
            Left            =   1710
            MaxLength       =   20
            TabIndex        =   9
            Top             =   585
            Width           =   1680
         End
         Begin VB.ComboBox Cmb_Sex 
            Height          =   315
            Left            =   4891
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   300
            Width           =   1680
         End
         Begin CS_DateControl.DateControl Dtp_DOB 
            Height          =   345
            Left            =   1710
            TabIndex        =   4
            Top             =   270
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   609
         End
         Begin VB.Label lbl_add 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NIC No."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4245
            TabIndex        =   92
            Top             =   660
            Width           =   570
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Marital Status"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   7005
            TabIndex        =   93
            Top             =   660
            Width           =   1125
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Social Sec No."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   375
            TabIndex        =   94
            Top             =   645
            Width           =   1215
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000005&
            X1              =   60
            X2              =   13780
            Y1              =   2955
            Y2              =   2955
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000005&
            X1              =   45
            X2              =   13765
            Y1              =   1125
            Y2              =   1125
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Address"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   840
            TabIndex        =   88
            Top             =   1350
            Width           =   735
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Date of Birth"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   570
            TabIndex        =   87
            Top             =   345
            Width           =   1005
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Contact No."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   86
            Top             =   2198
            Width           =   975
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "E-mail"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1080
            TabIndex        =   85
            Top             =   2535
            Width           =   495
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Additional Info"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   390
            TabIndex        =   84
            Top             =   3180
            Width           =   1185
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Expatriate"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   11130
            TabIndex        =   75
            Top             =   345
            Width           =   810
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Family Details"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   6990
            TabIndex        =   66
            Top             =   1395
            Width           =   1125
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Qualification"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   7110
            TabIndex        =   65
            Top             =   2220
            Width           =   1005
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Blood Group"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   10950
            TabIndex        =   64
            Top             =   660
            Width           =   1020
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Nationality"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   210
            Left            =   7275
            TabIndex        =   62
            Top             =   345
            Width           =   840
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Gender"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4185
            TabIndex        =   61
            Top             =   345
            Width           =   615
         End
         Begin VB.Label lbl_po 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Specialist In"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   7125
            TabIndex        =   60
            Top             =   2520
            Width           =   990
         End
      End
      Begin FPUSpreadADO.fpSpread Va_Leave 
         Height          =   5310
         Left            =   -73620
         TabIndex        =   42
         Top             =   645
         Width           =   10680
         _Version        =   458752
         _ExtentX        =   18838
         _ExtentY        =   9366
         _StockProps     =   64
         AutoClipboard   =   0   'False
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   11
         MaxRows         =   50
         ProcessTab      =   -1  'True
         SpreadDesigner  =   "frm_Emp_Master.frx":1AFD2
      End
      Begin FPUSpreadADO.fpSpread Va_Remarks 
         Height          =   5820
         Left            =   -74925
         TabIndex        =   105
         Top             =   435
         Width           =   14010
         _Version        =   458752
         _ExtentX        =   24712
         _ExtentY        =   10266
         _StockProps     =   64
         AutoClipboard   =   0   'False
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   3
         MaxRows         =   50
         ProcessTab      =   -1  'True
         SpreadDesigner  =   "frm_Emp_Master.frx":1B6EA
      End
      Begin VB.Frame Frm_SalaryType 
         Height          =   855
         Left            =   -74850
         TabIndex        =   80
         Top             =   330
         Width           =   13875
         Begin VB.TextBox Txtn_CarBenefit 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   4380
            TabIndex        =   39
            Top             =   337
            Width           =   1440
         End
         Begin VB.ComboBox Cmb_SalaryType 
            Height          =   315
            Left            =   1230
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   330
            Width           =   1590
         End
         Begin VB.ComboBox Cmb_TAType 
            Height          =   315
            Left            =   7515
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   330
            Width           =   1440
         End
         Begin VB.Label Label62 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Car Benefit"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   3390
            TabIndex        =   83
            Top             =   375
            Width           =   915
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Travel Allow"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   6405
            TabIndex        =   82
            Top             =   375
            Width           =   1020
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Salary Paid"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   270
            TabIndex        =   81
            Top             =   382
            Width           =   885
         End
      End
      Begin FPUSpreadADO.fpSpread Va_Salary 
         Height          =   4740
         Left            =   -74835
         TabIndex        =   41
         Top             =   1215
         Width           =   6360
         _Version        =   458752
         _ExtentX        =   11218
         _ExtentY        =   8361
         _StockProps     =   64
         AutoClipboard   =   0   'False
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   5
         MaxRows         =   50
         ProcessTab      =   -1  'True
         SpreadDesigner  =   "frm_Emp_Master.frx":1BA3F
      End
   End
   Begin VB.Label lbl_scr_name 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Employee Master Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   255
      TabIndex        =   91
      Top             =   930
      Width           =   2445
   End
   Begin VB.Label Label31 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2700
      TabIndex        =   74
      Top             =   1350
      Width           =   360
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Other Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   9885
      TabIndex        =   63
      Top             =   1350
      Width           =   975
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Emp No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   420
      TabIndex        =   58
      Top             =   1350
      Width           =   660
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5295
      TabIndex        =   57
      Top             =   1350
      Width           =   465
   End
   Begin VB.Label lbl_date 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   12900
      TabIndex        =   55
      Top             =   930
      Width           =   375
   End
   Begin VB.Label lblheader 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   54
      Top             =   885
      Width           =   14295
   End
End
Attribute VB_Name = "frm_Emp_Master"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As New ADODB.Recordset

Private Sub Form_Load()
    lbl_date.Caption = Format(Date, "dd-mmm-yyyy")
    Call ScreenUserRight(Me)
    
    Call Create_Default_Employee
    Call Spread_Lock
    Call TGControlProperty(Me)
    Call Spread_Row_Height(Va_Salary)
    Call Spread_Row_Height(Va_Leave)
    Call Spread_Row_Height(Va_Remarks)
    Call Spread_Row_Height(Va_Edu)
    Call Spread_Row_Height(Va_Exp)
    Call Combo_Load

    Txtn_MLDays.Visible = False
    Opt_Active.Value = True
    
    Txtn_EdfAmt.Enabled = False
    Txtn_EdfAmt.BackColor = &HE0E0E0
    Call Cmb_TpFlag_Click
    
    ' --- SIFB
    Frm_Salary.Visible = False
    Frm_SalaryType.Visible = False
    Va_Salary.Visible = False
    SSTab1.TabVisible(3) = False
    SSTab1.Tab = 0
    '  --- SIFB
    
    Cmb_Nationality.ListIndex = 0
    Cmb_Company.ListIndex = 1
    Cmb_StaffType.ListIndex = 1
    Chk_ClockCard.Value = 1
    Chk_Disabled.Value = 0
    Cmb_TpFlag.ListIndex = 1
    Cmb_PayType.ListIndex = 0
    Cmb_EDFCategory.ListIndex = 1
End Sub

Private Sub Form_Activate()
    If Txtc_EmpNo.Enabled = True Then
       Txtc_EmpNo.SetFocus
    End If
End Sub

Private Sub Btn_Exit_Click()
    Unload Me
End Sub


Private Sub Btn_Print_Click()
On Error GoTo Err_Print
  Dim SelFor As String, RepTitle As String
  
    If Trim(Txtc_EmpNo) = "" Then
      Exit Sub
    End If
    
    RepTitle = "Employee Information"
    SelFor = "{V_PR_EMP_MST.c_empno} = '" & Trim(Txtc_EmpNo) & "'"
    Call Print_Rpt(SelFor, "Pr_Employee_Info.rpt")
    
    If Trim(RepTitle) <> "" Then
       Mdi_Ta_HrPay.CRY1.Formulas(1) = "ReportHead='" & UCase(RepTitle) & "'"
    End If
    Mdi_Ta_HrPay.CRY1.Action = 1
  
  Exit Sub

Err_Print:
    MsgBox "Error while Generating - " + Err.Description, vbInformation, "Information"
End Sub

Private Sub Btn_Clear_Click()
    Clear_Controls Me
    Clear_Spread Va_Salary
    Clear_Spread Va_Leave
    Clear_Spread Va_Remarks
    Clear_Spread Va_Exp
    Clear_Spread Va_Edu

    Txtc_EmpNo.Enabled = True
    Txtn_EdfAmt.Enabled = False
    Txtn_EdfAmt.BackColor = &HE0E0E0
End Sub

Private Sub Btn_Delete_Click()
On Error GoTo ErrDel
    
    If Trim(Txtc_EmpNo) = "" Then
       Exit Sub
    End If

    If g_Admin Or g_FrmSupUser Then
       If (MsgBox("Are you sure you want to delete ?", vbExclamation + vbYesNo, lbl_scr_name.Caption) = vbYes) Then
          CON.BeginTrans
          g_Sql = "update pr_emp_mst set " & GetDelFlag & " where c_empno = '" & Trim(Txtc_EmpNo) & "'"
          CON.Execute g_Sql
          CON.CommitTrans
          Call Btn_Clear_Click
       End If
    Else
       MsgBox "No access to delete. Please contact Admin", vbInformation, "Information"
    End If
    
  Exit Sub

ErrDel:
    CON.RollbackTrans
    MsgBox "Error while Deleting - " + Err.Description, vbCritical, "Critical"
End Sub

Private Sub Btn_View_Click()
  Dim Search As New Search.MyClass, SerVar
  Dim tmpFilter As String

    If Opt_Active = True Then
       tmpFilter = " and d_dol is null "
    Else
       tmpFilter = " and d_dol is not null "
    End If


    Search.Query = "select c_empno EmpNo, c_name Name, c_othername OtherName, c_branch Branch, c_dept Department, c_desig Designation, c_emptype Type " & _
                   "from pr_emp_mst where c_rec_sta = 'A' " & tmpFilter
    Search.CheckFields = "EmpNo"
    Search.ReturnField = "EmpNo"
    SerVar = Search.Search(, , CON)
    If Len(Search.col1) <> 0 Then
       Call CancelButtonClick
       Txtc_EmpNo = Search.col1
       Call Display_Records
       SSTab1.TabIndex = 0
    End If
End Sub

Private Sub Btn_Save_Click()
On Error GoTo ErrSave
  
     If ChkSave = False Then
        Exit Sub
     End If
     
     Screen.MousePointer = vbHourglass
     g_SaveFlagNull = True
     
     CON.BeginTrans
        Save_Pr_Emp_Mst
        Save_Pr_Emp_Salary_Dtl
        
        ' sifb - nj
        'Save_Pr_Emp_Leave_Dtl
        
        Save_Pr_Emp_Remarks_Dtl
        Save_Pr_Emp_Education_Dtl
        Save_Pr_Emp_Experience_Dtl
        Save_Pr_Emp_WorkPermit_Dtl
     CON.CommitTrans
        
     CON.BeginTrans
        Save_Master_Update
     CON.CommitTrans
     
     g_SaveFlagNull = False
     Screen.MousePointer = vbDefault
    
     MsgBox "Record Saved Successfully", vbInformation, "Information"
     
   
   Exit Sub
     
ErrSave:
     CON.RollbackTrans
     g_SaveFlagNull = False
     Screen.MousePointer = vbDefault
     MsgBox "Error while Saving - " + Err.Description, vbCritical, "Critical"
End Sub

Private Function ChkSave() As Boolean
    Dim rsChk As New ADODB.Recordset
    Dim tmpStr As String
    
    If Trim(Txtc_EmpNo) = "" Then
       MsgBox "Employee No. should not be empty", vbInformation, "Information"
       Txtc_EmpNo.SetFocus
       Exit Function
    ElseIf Trim(Txtc_Name) = "" Then
       MsgBox "Employee Name Should not be empty", vbInformation, "Information"
       Txtc_Name.SetFocus
       Exit Function
    
    ElseIf Trim(Cmb_Company.Text) = "" Then
       MsgBox "Company should not be empty", vbInformation, "Information"
       Cmb_Company.SetFocus
       Exit Function
    
    ElseIf Not IsDate(Dtp_DOJ.Text) Then
       MsgBox "Date of Join should not be empty", vbInformation, "Information"
       Dtp_DOJ.SetFocus
       Exit Function
    
  
    ElseIf Not AmountCheck Then
       Exit Function
    ElseIf IsColTwoDupValues(Va_Salary, 2, 5) > 0 Then
       MsgBox "Duplicate salary type found. Please check salary details", vbInformation, "Information"
       Va_Salary.SetFocus
       Exit Function
    End If
    
   
    If IsDate(Dtp_DOJ.Text) And IsDate(Dtp_DOL.Text) Then
       If CDate(Dtp_DOJ.Text) > CDate(Dtp_DOL.Text) Then
          MsgBox "Date of Left should not be less than Date of Join"
          Dtp_DOL.SetFocus
          Exit Function
       End If
    
'       If Not g_Admin And (CDate(Dtp_DOJ.Text) > (g_CurrentDate + 90) Or CDate(Dtp_DOJ.Text) < (g_CurrentDate - 90)) Then
'          MsgBox "Date of Join should be input 90 days either side of current date. Please contact Admin", vbInformation, "Information"
'          Dtp_DOJ.SetFocus
'          Exit Function
'       End If
'
'       If Not g_Admin And (CDate(Dtp_DOL.Text) > (g_CurrentDate + 90) Or CDate(Dtp_DOL.Text) < (g_CurrentDate - 90)) Then
'          MsgBox "Date left should be input 90 days either side of current date. Please contact Admin", vbInformation, "Information"
'          Dtp_DOL.SetFocus
'          Exit Function
'       End If
    
       Set rsChk = Nothing
       g_Sql = "select n_loanamount-n_loanpaid n_amount " & _
               "from pr_loan_mst where c_rec_sta = 'A' and n_loanamount - n_loanpaid > 0 and " & _
               "c_empno = '" & Trim(Txtc_EmpNo) & "'"
       rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
       If rsChk.RecordCount > 0 Then
          MsgBox "This employee having Advance balance Rs. " & Is_Null(rsChk("n_amount").Value, True) & "  " & _
                 "Please check Advance details. ", vbInformation, "Information"
       End If
    End If
    
    If Trim(Txtc_AcctNo) <> "" Then
       If Right(Trim(Txtc_Bank), 3) = "B01" And (Trim(Txtc_BankCode) = "" Or Trim(Txtc_BankCode) = "11") And Len(Trim(Txtc_AcctNo)) <> 14 Then
          MsgBox "SBM Account No. should be 14 digit. Please check your entry", vbInformation, "Information"
          Txtc_Bank.SetFocus
          Exit Function
       End If
     
        Set rsChk = Nothing
        g_Sql = "select c_empno, c_name, c_othername, d_dol " & _
                "from pr_emp_mst where c_rec_sta = 'A' and c_acctno = '" & Trim(Txtc_AcctNo) & "' and c_empno <> '" & Trim(Txtc_EmpNo) & "'"
        rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
        If rsChk.RecordCount > 0 Then
           tmpStr = "The Account No. is already exists for " & _
                    Is_Null(rsChk("c_empno").Value, False) & Space(2) & Is_Null(rsChk("c_name"), False) & Space(2) & Is_Null(rsChk("c_othername"), False)
           If IsDate(rsChk("d_dol").Value) Then
              tmpStr = tmpStr & ". But who is left on " & Format(rsChk("d_dol").Value, "dd/mm/yyyy")
           End If
           tmpStr = tmpStr & ". Please Check your entry."
           MsgBox tmpStr, vbInformation, "Information"
        End If
    End If
          
    ChkSave = True
End Function

Private Sub Save_Pr_Emp_Mst()
    Set rs = Nothing
    g_Sql = "Select * from pr_emp_mst where c_empno = '" & Trim(Txtc_EmpNo) & "'"
    rs.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
    
    If rs.RecordCount = 0 Then
       rs.AddNew
       rs("c_usr_id").Value = g_UserName
       rs("d_created").Value = GetDateTime
    Else
       rs("c_musr_id").Value = g_UserName
       rs("d_modified").Value = GetDateTime
    End If
       
       rs("c_empno").Value = Is_Null(Txtc_EmpNo, False)
       rs("c_title").Value = Is_Null(Cmb_Title, False)
       rs("c_name").Value = Is_Null(Txtc_Name, False)
       rs("c_othername").Value = Is_Null(Txtc_OtherName, False)
       
       rs("d_dob").Value = Is_Date(Dtp_DOB.Text, "S")
       rs("c_sex").Value = Is_Null(Cmb_Sex, False)
       rs("c_nationality").Value = Is_Null(Cmb_Nationality, False)
       rs("c_expatriate").Value = IIf(Trim(Cmb_Expatriate) <> "", Is_Null(Cmb_Expatriate, False), "N")
       
       rs("c_socsecno").Value = Is_Null(Txtc_SocSecNo, False)
       rs("c_nicno").Value = Is_Null(Txtc_NICNo, False)
       rs("c_matstatus").Value = Is_Null(Cmb_MatStatus, False)
       rs("c_bloodgroup").Value = Is_Null(Txtc_BloodGroup, False)
       rs("c_address").Value = Is_Null(Proper(Txtc_Add), False)
       rs("c_phone").Value = Is_Null(Txtc_PerContNo, False)
       rs("c_email").Value = Is_Null(Txtc_EMail, False)
       
       rs("c_qualification").Value = Is_Null(Txtc_Qualify, False)
       rs("c_specialistin").Value = Is_Null(Txtc_Specialist, False)
       rs("c_familydetails").Value = Is_Null(Txtc_FamilyDetail, False)
       rs("c_additionalinfo").Value = Is_Null(Txtc_AddInfo, False)
       
       rs("c_company").Value = Is_Null(Right(Cmb_Company, 7), False)
       rs("c_branch").Value = Is_Null(Cmb_Branch, False)
       rs("c_dept").Value = Is_Null(Cmb_Dept, False)
       rs("c_desig").Value = Is_Null(Cmb_Desig, False)
       
       rs("d_doj").Value = Is_Date(Dtp_DOJ.Text, "S")
       rs("d_dol").Value = Is_Date(Dtp_DOL.Text, "S")
       
       rs("c_skillset").Value = Is_Null(Cmb_SkillSet, False)
       rs("c_line").Value = Is_Null(Txtc_Line, False)
       rs("c_shiftcode").Value = Is_Null(Right(Trim(Txtc_ShiftCode), 3), False)
       rs("c_daywork").Value = Is_Null(Right(Cmb_DayWork, 3), False)
       rs("c_emptype").Value = Is_Null(Proper(Txtc_EmpType), False)
       rs("c_stafftype").Value = Is_Null(Right(Cmb_StaffType, 1), False)
       
       rs("n_mldays").Value = Is_Null(Txtn_MLDays, True)
       rs("d_mlfrom").Value = Is_Date(Dtp_MLFrom.Text, "S")
       rs("d_mlto").Value = Is_Date(Dtp_MLTo.Text, "S")
       
       rs("c_tpflag").Value = Is_Null(Cmb_TpFlag, False)
       rs("c_tpmode").Value = Is_Null(Txtc_TPMode, False)
       rs("c_town").Value = Is_Null(Txtc_Town, False)
       rs("c_road").Value = Is_Null(Txtc_Road, False)
       rs("c_pad").Value = Is_Null(Txtc_PAD, False)
       
       rs("c_paytype").Value = Is_Null(Right(Cmb_PayType, 2), False)
       rs("c_bank").Value = Is_Null(Right(Trim(Txtc_Bank), 3), False)
       rs("c_bankcode").Value = Is_Null(Txtc_BankCode, False)
       rs("c_acctno").Value = Is_Null(Txtc_AcctNo, False)
       rs("c_itno").Value = Is_Null(Txtc_ITNo, False)
       
       rs("c_edfcat").Value = Is_Null(Left(Trim(Cmb_EDFCategory), 1), False)
       rs("n_edfamount").Value = Is_Null_D(Txtn_EdfAmt, True)
       rs("n_eduamount").Value = Is_Null_D(Txtn_EduAmt, True)
       rs("n_intamount").Value = Is_Null_D(Txtn_IntAmt, True)
       rs("n_preamount").Value = Is_Null_D(Txtn_PreAmt, True)
       rs("n_othamount").Value = Is_Null_D(Txtn_OthAmt, True)
       
       rs("c_salarytype").Value = Is_Null(Right(Cmb_SalaryType, 2), False)
       rs("c_tatype").Value = Is_Null(Right(Cmb_TAType, 1), False)
       rs("n_carbenefit").Value = Is_Null_D(Txtn_CarBenefit, True)
       
       rs("c_clockcard").Value = Chk_ClockCard.Value
       rs("c_disabled").Value = Chk_Disabled.Value
       rs("c_nopay").Value = Chk_NoPay.Value
       rs("c_mealallow").Value = Chk_MealAllow
       rs("c_payerelief").Value = Chk_PayeRelief.Value
       rs("c_npfdeduct").Value = Chk_NPFDeduct.Value
       rs("c_nobonus").Value = Chk_NoBonus.Value
       
       rs("c_rec_sta").Value = "A"
       rs.Update
End Sub

Private Sub Save_Pr_Emp_Salary_Dtl()
 Dim i As Long

    g_Sql = "delete from pr_emp_salary_dtl where c_empno = '" & Trim(Txtc_EmpNo) & "' and c_yrstatus = 'C'"
    CON.Execute (g_Sql)
    
    Set rs = Nothing
    g_Sql = "Select * from pr_emp_salary_dtl where 1=2"
    rs.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
    
    For i = 1 To Va_Salary.DataRowCnt
        Va_Salary.Row = i
        Va_Salary.Col = 3
        If Val(Va_Salary.Text) > 0 Then
           rs.AddNew
           rs("c_empno").Value = Is_Null(Txtc_EmpNo, False)
           Va_Salary.Col = 1
              rs("d_prfrom").Value = Is_Date(Va_Salary.Text, "S")
           Va_Salary.Col = 2
              rs("c_salary").Value = Is_Null(Right(Trim(Va_Salary.Text), 7), False)
           Va_Salary.Col = 3
              rs("n_amount").Value = Is_Null_D(Va_Salary.Text, True, True)
           rs("c_yrstatus").Value = "C"
           rs.Update
        End If
    Next i
End Sub

Private Sub Save_Pr_Emp_Leave_Dtl()
 Dim i As Long

        g_Sql = "delete from pr_emp_leave_dtl where c_empno = '" & Trim(Txtc_EmpNo) & "'"
        CON.Execute (g_Sql)

        Set rs = Nothing
        g_Sql = "Select * from pr_emp_leave_dtl where 1=2"
        rs.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
        
        For i = 1 To Va_Leave.DataRowCnt
            Va_Leave.Row = i
            Va_Leave.Col = 2
            If Trim(Va_Leave.Text) <> "" Then
               rs.AddNew
               rs("c_empno").Value = Is_Null(Txtc_EmpNo, False)
               Va_Leave.Col = 1
                  rs("d_prfrom").Value = Is_Date(Va_Leave.Text, "S")
               Va_Leave.Col = 2
                  rs("c_leave").Value = Is_Null(Right(Trim(Va_Leave.Text), 7), False)
               Va_Leave.Col = 3
                  rs("n_opbal").Value = Is_Null_D(Va_Leave.Text, True, True)
               Va_Leave.Col = 4
                  rs("n_entitle").Value = Is_Null_D(Va_Leave.Text, True, True)
               Va_Leave.Col = 5
                  rs("n_alloted").Value = Is_Null_D(Va_Leave.Text, True, True)
               Va_Leave.Col = 6
                  rs("n_utilised").Value = Is_Null_D(Va_Leave.Text, True, True)
               Va_Leave.Col = 7
                  rs("n_adjusted").Value = Is_Null_D(Va_Leave.Text, True, True)
               Va_Leave.Col = 8
                  rs("n_clbal").Value = Is_Null_D(Va_Leave.Text, True, True)
               Va_Leave.Col = 9
                  rs("n_othleave").Value = Is_Null_D(Va_Leave.Text, True, True)
               Va_Leave.Col = 10
                  rs("n_sickexcess").Value = Is_Null_D(Va_Leave.Text, True, True)
               Va_Leave.Col = 11
                  If Trim(Va_Leave.Text) = "" Then
                     rs("c_yrstatus").Value = "C"
                  Else
                     rs("c_yrstatus").Value = Is_Null(Va_Leave.Text, False)
                  End If
                  
               rs.Update
            End If
        Next i
End Sub


Private Sub Save_Pr_Emp_Remarks_Dtl()
 Dim i As Long
 
    g_Sql = "delete from pr_emp_remarks_dtl where c_empno = '" & Trim(Txtc_EmpNo) & "'"
    CON.Execute (g_Sql)
    
    Set rs = Nothing
    g_Sql = "Select * from pr_emp_remarks_dtl where 1=2"
    rs.Open g_Sql, CON, adOpenDynamic, adLockOptimistic

    For i = 1 To Va_Remarks.DataRowCnt
        Va_Remarks.Row = i
        Va_Remarks.Col = 1
        If Trim(Va_Remarks.Text) <> "" Then
           rs.AddNew
           rs("c_empno").Value = Is_Null(Txtc_EmpNo, False)
           Va_Remarks.Col = 1
              rs("d_date").Value = Is_Date(Va_Remarks.Text, "S")
           Va_Remarks.Col = 2
              rs("c_type").Value = Is_Null(Va_Remarks.Value, True)
           Va_Remarks.Col = 3
              rs("c_remarks").Value = Is_Null(Va_Remarks.Text, False)
           rs.Update
        End If
    Next i
End Sub

Private Sub Save_Pr_Emp_Education_Dtl()
 Dim i As Long
 Dim vSeq As Integer
 Dim tmpStr As String
    
    g_Sql = "delete from pr_emp_education_dtl where c_empno = '" & Trim(Txtc_EmpNo) & "'"
    CON.Execute (g_Sql)
       
    Set rs = Nothing
    g_Sql = "Select * from pr_emp_Education_dtl where 1=2"
    rs.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
       
    vSeq = 0
    For i = 1 To Va_Edu.DataRowCnt
        Va_Edu.Row = i
        Va_Edu.Col = 1
           tmpStr = Trim(Va_Edu.Text)
        Va_Edu.Col = 2
           If Trim(tmpStr) <> "" Or Trim(Va_Edu.Text) <> "" Then
              rs.AddNew
              rs("c_empno").Value = Is_Null(Txtc_EmpNo, False)
              vSeq = vSeq + 1
              rs("n_seq").Value = vSeq
              Va_Edu.Col = 1
                 rs("c_education").Value = Is_Null(Va_Edu.Text, False)
              Va_Edu.Col = 2
                 rs("c_subject").Value = Is_Null(Va_Edu.Text, False)
              Va_Edu.Col = 3
                 rs("c_grade").Value = Is_Null(Va_Edu.Text, False)
              rs.Update
           End If
    Next i
End Sub

Private Sub Save_Pr_Emp_Experience_Dtl()
 Dim i As Long
 Dim vSeq As Integer
    
    g_Sql = "delete from pr_emp_experience_dtl where c_empno = '" & Trim(Txtc_EmpNo) & "'"
    CON.Execute (g_Sql)
       
    Set rs = Nothing
    g_Sql = "Select * from pr_emp_Experience_dtl where 1=2"
    rs.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
    
    vSeq = 0
    For i = 1 To Va_Exp.DataRowCnt
        Va_Exp.Row = i
        Va_Exp.Col = 1
        If Trim(Va_Exp.Text) <> "" Then
           rs.AddNew
           rs("c_empno").Value = Is_Null(Txtc_EmpNo, False)
           vSeq = vSeq + 1
           rs("n_seq").Value = vSeq
           Va_Exp.Col = 1
              rs("c_employer").Value = Is_Null(Va_Exp.Text, False)
           Va_Exp.Col = 2
              rs("c_desig").Value = Is_Null(Va_Exp.Text, False)
           Va_Exp.Col = 3
              rs("d_fromDate").Value = Is_Date(Va_Exp.Text, "S")
           Va_Exp.Col = 4
              rs("d_todate").Value = Is_Date(Va_Exp.Text, "S")
           rs.Update
        End If
    Next i
End Sub

Private Sub Save_Pr_Emp_WorkPermit_Dtl()
 Dim i As Long
    
    g_Sql = "delete from pr_emp_workpermit_dtl where c_empno = '" & Trim(Txtc_EmpNo) & "'"
    CON.Execute (g_Sql)
    
   ' If Trim(Txtc_PassNo) <> "" Or Trim(Txtc_WrkPermit) <> "" Or Trim(Txtc_MinRefNo) <> "" Or Trim(Txtc_GuaranteeNo) <> "" Or Trim(Txtc_GuaranteePassport) <> "" Then
       Set rs = Nothing
       g_Sql = "Select * from pr_emp_workpermit_dtl where 1=2"
       rs.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
       
       rs.AddNew
       rs("c_empno").Value = Is_Null(Txtc_EmpNo, False)
       rs("c_passno").Value = Is_Null(Txtc_PassNo, False)
       rs("d_passexpdt").Value = Is_Date(Txtd_PassExpDt.Text, "S")
       
       rs("c_wrkprmtno").Value = Is_Null(Txtc_WrkPermit, False)
       rs("d_wrkdate").Value = Is_Date(Txtd_WrkPrmtExpiry.Text, "S")
       
       rs("c_minrefno").Value = Is_Null(Txtc_MinRefNo, False)
       rs("c_guaranteeno").Value = Is_Null(Txtc_GuaranteeNo, False)
       rs("c_guaranteepassport").Value = Is_Null(Txtc_GuaranteePassport, False)
       
       rs("c_flightno").Value = Is_Null(Txtc_Flightno, False)
       rs("d_flightdate").Value = Is_Date(Txtd_FlightDate.Text, "S")
       
       rs.Update
  '  End If
End Sub

Private Sub Save_Master_Update()
  Dim rsUpd As New ADODB.Recordset
  Dim i, nBasic, nFR2, nFR3, nFR4, nFR5 As Double

    Set rsUpd = Nothing
    g_Sql = "select c_salary, n_amount from pr_emp_salary_dtl " & _
            "where c_salary in ('SAL0001','SAL0002','SAL0003','SAL0004','SAL0005') and c_yrstatus = 'C' and " & _
            "c_empno = '" & Trim(Txtc_EmpNo) & "'"
    rsUpd.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If rsUpd.RecordCount > 0 Then
       rsUpd.MoveFirst
       nBasic = 0: nFR2 = 0: nFR3 = 0: nFR4 = 0: nFR5 = 0
       For i = 1 To rsUpd.RecordCount
          If rsUpd("c_salary").Value = "SAL0001" Then
             nBasic = Is_Null(rsUpd("n_amount").Value, True)
          ElseIf rsUpd("c_salary").Value = "SAL0002" Then
             nFR2 = Is_Null(rsUpd("n_amount").Value, True)
          ElseIf rsUpd("c_salary").Value = "SAL0003" Then
             nFR3 = Is_Null(rsUpd("n_amount").Value, True)
          ElseIf rsUpd("c_salary").Value = "SAL0004" Then
             nFR4 = Is_Null(rsUpd("n_amount").Value, True)
          ElseIf rsUpd("c_salary").Value = "SAL0005" Then
             nFR5 = Is_Null(rsUpd("n_amount").Value, True)
          End If
          rsUpd.MoveNext
       Next i
       g_Sql = "update pr_emp_mst set n_basic = " & nBasic & ", " & _
               "n_fixedrate2 = " & nFR2 & ", n_fixedrate3 = " & nFR3 & ", n_fixedrate4 = " & nFR4 & ", n_fixedrate5 = " & nFR5 & " " & _
               "where c_empno = '" & Trim(Txtc_EmpNo) & "'"
       CON.Execute g_Sql
    End If
End Sub

Private Sub Display_Records()
On Error GoTo Err_Display
  Dim DyDisp As New ADODB.Recordset
  Dim dyParty As New ADODB.Recordset
  Dim i, j As Long

  Set DyDisp = Nothing
  g_Sql = "select * from pr_emp_mst where c_empno = '" & Trim(Txtc_EmpNo) & "' and c_rec_sta = 'A'"
  DyDisp.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
  If DyDisp.RecordCount <= 0 Then
     Call Get_Clocking_Employee
     Exit Sub
  End If
  
  If DyDisp.RecordCount > 0 Then
     Txtc_EmpNo.Enabled = False
     Txtc_EmpNo = Is_Null(DyDisp("c_empno").Value, False)
     Txtc_Name = Is_Null(DyDisp("c_name").Value, False)
     Txtc_OtherName = Is_Null(DyDisp("c_othername").Value, False)
     Dtp_DOB.Text = Is_Date(DyDisp("d_dob").Value, "D")
     
     For i = 0 To Cmb_Title.ListCount - 1
       If Trim(Cmb_Title.List(i)) = DyDisp("c_title").Value Then
          Cmb_Title.ListIndex = i
          Exit For
       End If
     Next i
     For i = 0 To Cmb_Sex.ListCount - 1
       If Trim(Cmb_Sex.List(i)) = DyDisp("c_sex").Value Then
          Cmb_Sex.ListIndex = i
          Exit For
       End If
     Next i
     For i = 0 To Cmb_Expatriate.ListCount - 1
       If Trim(Cmb_Expatriate.List(i)) = DyDisp("c_expatriate").Value Then
          Cmb_Expatriate.ListIndex = i
          Exit For
       End If
     Next i
     For i = 0 To Cmb_Nationality.ListCount - 1
       If Trim(Cmb_Nationality.List(i)) = DyDisp("c_nationality").Value Then
          Cmb_Nationality.ListIndex = i
          Exit For
       End If
     Next i
     For i = 0 To Cmb_MatStatus.ListCount - 1
       If Trim(Cmb_MatStatus.List(i)) = DyDisp("c_matstatus").Value Then
          Cmb_MatStatus.ListIndex = i
          Exit For
       End If
     Next i

     Txtc_SocSecNo = Is_Null(DyDisp("c_socsecno").Value, False)
     Txtc_NICNo = Is_Null(DyDisp("c_nicno").Value, False)
     Txtc_BloodGroup = Is_Null(DyDisp("c_bloodgroup").Value, False)
     Txtc_Add = Is_Null(DyDisp("c_address").Value, False)
     Txtc_PerContNo = Is_Null(DyDisp("c_phone").Value, False)
     Txtc_EMail = Is_Null(DyDisp("c_email").Value, False)
     Txtc_Qualify = Is_Null(DyDisp("c_qualification").Value, False)
     Txtc_Specialist = Is_Null(DyDisp("c_specialistin").Value, False)
     Txtc_FamilyDetail = Is_Null(DyDisp("c_familydetails").Value, False)
     Txtc_AddInfo = Is_Null(DyDisp("c_additionalinfo").Value, False)

     Call DisplayComboCompany(Me, Is_Null(DyDisp("c_company").Value, False))
     Call DisplayComboBranch(Me, Is_Null(DyDisp("c_branch").Value, False))
     Call DisplayComboDept(Me, Is_Null(DyDisp("c_dept").Value, False))
     Call DisplayComboDesig(Me, Is_Null(DyDisp("c_desig").Value, False))
    
     Dtp_DOJ.Text = Is_Date(DyDisp("d_doj").Value, "D")
     Dtp_DOL.Text = Is_Date(DyDisp("d_dol").Value, "D")
     
     Txtn_MLDays = Is_Null(DyDisp("n_mldays").Value, True)
     Dtp_MLFrom.Text = Is_Date(DyDisp("d_mlfrom").Value, "D")
     Dtp_MLTo.Text = Is_Date(DyDisp("d_mlto").Value, "D")

     For i = 0 To Cmb_SkillSet.ListCount - 1
       If Trim(Cmb_SkillSet.List(i)) = Is_Null(DyDisp("c_skillset").Value, False) Then
          Cmb_SkillSet.ListIndex = i
          Exit For
       End If
     Next i
     For i = 0 To Cmb_PayType.ListCount - 1
       If Trim(Right(Cmb_PayType.List(i), 2)) = Is_Null(DyDisp("c_paytype").Value, False) Then
          Cmb_PayType.ListIndex = i
          Exit For
       End If
     Next i
     Txtc_EmpType = Is_Null(DyDisp("c_emptype").Value, False)
     Txtc_Line = Is_Null(DyDisp("c_line").Value, False)
     Txtc_ShiftCode = Is_Null(DyDisp("c_shiftcode").Value, False)
     Call Txtc_ShiftCode_Validate(True)

     For i = 0 To Cmb_DayWork.ListCount - 1
       If Trim(Right(Cmb_DayWork.List(i), 3)) = Is_Null(DyDisp("c_daywork").Value, False) Then
          Cmb_DayWork.ListIndex = i
          Exit For
       End If
     Next i
     For i = 0 To Cmb_SalaryType.ListCount - 1
       If Trim(Right(Cmb_SalaryType.List(i), 2)) = Is_Null(DyDisp("c_salarytype").Value, False) Then
          Cmb_SalaryType.ListIndex = i
          Exit For
       End If
     Next i

     For i = 0 To Cmb_StaffType.ListCount - 1
       If Trim(Right(Cmb_StaffType.List(i), 1)) = Is_Null(DyDisp("c_stafftype").Value, False) Then
          Cmb_StaffType.ListIndex = i
          Exit For
       End If
     Next i
     For i = 0 To Cmb_TAType.ListCount - 1
       If Trim(Right(Cmb_TAType.List(i), 1)) = Is_Null(DyDisp("c_tatype").Value, False) Then
          Cmb_TAType.ListIndex = i
          Exit For
       End If
     Next i
     For i = 0 To Cmb_EDFCategory.ListCount - 1
       If Trim(Left(Trim(Cmb_EDFCategory.List(i)), 1)) = Is_Null(DyDisp("c_edfcat").Value, False) Then
          Cmb_EDFCategory.ListIndex = i
          Exit For
       End If
     Next i
     
     Txtn_EdfAmt = Format_Num(Is_Null(DyDisp("n_edfamount").Value, True))
     Txtn_EduAmt = Format_Num(Is_Null(DyDisp("n_eduamount").Value, True))
     Txtn_IntAmt = Format_Num(Is_Null(DyDisp("n_intamount").Value, True))
     Txtn_PreAmt = Format_Num(Is_Null(DyDisp("n_preamount").Value, True))
     Txtn_OthAmt = Format_Num(Is_Null(DyDisp("n_othamount").Value, True))

     Chk_ClockCard.Value = Is_Null(DyDisp("c_clockcard").Value, False)
     Chk_Disabled.Value = Is_Null(DyDisp("c_disabled").Value, False)
     Chk_NoPay.Value = Is_Null(DyDisp("c_nopay").Value, True)
     Chk_PayeRelief.Value = Is_Null(DyDisp("c_payerelief").Value, True)
     Chk_NPFDeduct.Value = Is_Null(DyDisp("c_npfdeduct").Value, True)
     Chk_NoBonus.Value = Is_Null(DyDisp("c_nobonus").Value, True)
     Chk_MealAllow.Value = Is_Null(DyDisp("c_mealallow").Value, True)

    
     For i = 0 To Cmb_TpFlag.ListCount - 1
         If Trim(Cmb_TpFlag.List(i)) = Is_Null(DyDisp("c_tpflag").Value, False) Then
            Cmb_TpFlag.ListIndex = i
            Exit For
         End If
     Next i
     Txtc_TPMode = Is_Null(DyDisp("c_tpmode").Value, False)
     Txtc_Town = Is_Null(DyDisp("c_town").Value, False)
     Txtc_Road = Is_Null(DyDisp("c_road").Value, False)
     Txtc_PAD = Is_Null(DyDisp("c_pad").Value, False)
    
     Txtc_Bank = Is_Null(DyDisp("c_bank").Value, False)
     Txtc_BankCode = Is_Null(DyDisp("c_bankcode").Value, False)
     Call Txtc_Bank_Validate(True)
     
     Txtc_ITNo = Is_Null(DyDisp("c_itno").Value, False)
     Txtc_AcctNo = Is_Null(DyDisp("c_acctno").Value, False)
     Txtn_CarBenefit = Format_Num(Is_Null_D(DyDisp("n_carbenefit").Value, True))
     
     Call Cmb_EDFCategory_Click
  End If

   Set DyDisp = Nothing
   g_Sql = "select a.c_empno, b.d_prfrom, b.c_salary, b.n_amount, c.c_payname " & _
           "from pr_emp_mst a, pr_emp_salary_dtl b, pr_paystructure_dtl c " & _
           "where a.c_empno = b.c_empno and b.c_salary = c.c_salary and a.c_company = c.c_company and " & _
           "b.c_yrstatus = 'C' and a.c_rec_sta = 'A' and a.c_empno = '" & Trim(Txtc_EmpNo) & "' " & _
           "order by a.c_empno, b.d_prfrom, c.n_seq, b.c_salary "

    DyDisp.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If DyDisp.RecordCount > 0 Then
       Va_Salary.MaxRows = DyDisp.RecordCount + 25
       DyDisp.MoveFirst
       For i = 1 To DyDisp.RecordCount
          Va_Salary.Row = i
          Va_Salary.Col = 1
             Va_Salary.Text = Is_DateSpread(DyDisp("d_prfrom").Value, False)
          Va_Salary.Col = 2
             Va_Salary.Text = Is_Null(DyDisp("c_payname").Value, False) & Space(100) & Is_Null(DyDisp("c_salary").Value, False)
          Va_Salary.Col = 3
             Va_Salary.Text = Is_Null(DyDisp("n_amount").Value, True)

         DyDisp.MoveNext
        Next i
    End If

   Set DyDisp = Nothing
   g_Sql = "select a.c_empno, b.d_prfrom, b.c_leave, b.n_opbal, b.n_entitle, b.n_alloted, b.n_utilised, b.n_adjusted, b.n_clbal, " & _
           "b.n_othleave, b.n_sickexcess, b.c_yrstatus, c.c_leavename " & _
           "from pr_emp_mst a, pr_emp_leave_dtl b, pr_leave_mst c " & _
           "where a.c_empno = b.c_empno and  b.c_leave = c.c_leave and a.c_rec_sta = 'A' and a.c_empno = '" & Trim(Txtc_EmpNo) & "' " & _
           "order by b.d_prfrom desc  "

    DyDisp.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If DyDisp.RecordCount > 0 Then
       Va_Leave.MaxRows = DyDisp.RecordCount + 25
       DyDisp.MoveFirst
       For i = 1 To DyDisp.RecordCount
          Va_Leave.Row = i
          Va_Leave.Col = 1
             Va_Leave.Text = Is_DateSpread(DyDisp("d_prfrom").Value, False)
          Va_Leave.Col = 2
             Va_Leave.Text = Is_Null(DyDisp("c_leavename").Value, False) & Space(100) & Is_Null(DyDisp("c_leave").Value, False)
          Va_Leave.Col = 3
             Va_Leave.Text = Spread_NumFormat(DyDisp("n_opbal").Value, True)
          Va_Leave.Col = 4
             Va_Leave.Text = Spread_NumFormat(DyDisp("n_entitle").Value, True)
          Va_Leave.Col = 5
             Va_Leave.Text = Spread_NumFormat(DyDisp("n_alloted").Value, True)
          Va_Leave.Col = 6
             Va_Leave.Text = Spread_NumFormat(DyDisp("n_utilised").Value, True)
          Va_Leave.Col = 7
             Va_Leave.Text = Spread_NumFormat(DyDisp("n_adjusted").Value, True)
          Va_Leave.Col = 8
             Va_Leave.Text = Spread_NumFormat(DyDisp("n_clbal").Value, True)
          Va_Leave.Col = 9
             Va_Leave.Text = Spread_NumFormat(DyDisp("n_othleave").Value, True)
          Va_Leave.Col = 10
             Va_Leave.Text = Spread_NumFormat(DyDisp("n_sickexcess").Value, True)
          Va_Leave.Col = 11
             Va_Leave.Text = Is_Null(DyDisp("c_yrstatus").Value, False)
         DyDisp.MoveNext
        Next i
    End If

   ' Remarks
   Set DyDisp = Nothing
   g_Sql = "select a.c_empno, b.d_date, b.c_type, b.c_remarks " & _
           "from pr_emp_mst a, pr_emp_remarks_dtl b " & _
           "where a.c_empno = b.c_empno and a.c_rec_sta = 'A' and a.c_empno = '" & Trim(Txtc_EmpNo) & "' " & _
           "order by a.c_empno, b.d_date "

    DyDisp.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If DyDisp.RecordCount > 0 Then
       Va_Remarks.MaxRows = DyDisp.RecordCount + 25
       DyDisp.MoveFirst
       For i = 1 To DyDisp.RecordCount
          Va_Remarks.Row = i
          Va_Remarks.Col = 1
             Va_Remarks.Text = Is_DateSpread(DyDisp("d_date").Value, False)
          Va_Remarks.Col = 2
             Va_Remarks.Value = Is_Null(DyDisp("c_type").Value, True)
          Va_Remarks.Col = 3
             Va_Remarks.Text = Is_Null(DyDisp("c_remarks").Value, False)
         DyDisp.MoveNext
        Next i
    End If
    
   ' Education
   Set DyDisp = Nothing
   g_Sql = "select a.c_empno, b.c_education, b.c_subject, b.c_Grade " & _
           "from pr_emp_mst a, pr_emp_education_dtl b " & _
           "where a.c_empno = b.c_empno and a.c_rec_sta ='A' and a.c_empno = '" & Trim(Txtc_EmpNo) & "' " & _
           "order by a.c_empno, b.n_seq "
    
    DyDisp.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If DyDisp.RecordCount > 0 Then
       Va_Edu.MaxRows = DyDisp.RecordCount + 25
       DyDisp.MoveFirst
       For i = 1 To DyDisp.RecordCount
          Va_Edu.Row = i
          Va_Edu.Col = 1
             Va_Edu.Text = Is_Null(DyDisp("c_education").Value, False)
          Va_Edu.Col = 2
             Va_Edu.Text = Is_Null(DyDisp("c_subject").Value, False)
          Va_Edu.Col = 3
             Va_Edu.Text = Is_Null(DyDisp("c_grade").Value, False)
         DyDisp.MoveNext
        Next i
    End If
  
   ' Experiance
   Set DyDisp = Nothing
   g_Sql = "select a.c_empno, b.c_employer, b.c_desig, b.d_fromdate, b.d_todate " & _
           "from pr_emp_mst a, pr_emp_experience_dtl b " & _
           "where a.c_empno = b.c_empno and a.c_rec_sta ='A' and a.c_empno = '" & Trim(Txtc_EmpNo) & "' " & _
           "order by a.c_empno, b.n_seq "
    
    DyDisp.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If DyDisp.RecordCount > 0 Then
       Va_Exp.MaxRows = DyDisp.RecordCount + 25
       DyDisp.MoveFirst
       For i = 1 To DyDisp.RecordCount
          Va_Exp.Row = i
          Va_Exp.Col = 1
             Va_Exp.Text = Is_Null(DyDisp("c_employer").Value, False)
          Va_Exp.Col = 2
             Va_Exp.Text = Is_Null(DyDisp("c_desig").Value, False)
          Va_Exp.Col = 3
             Va_Exp.Text = Is_DateSpread(DyDisp("d_fromdate").Value, False)
          Va_Exp.Col = 4
             Va_Exp.Text = Is_DateSpread(DyDisp("d_todate").Value, False)
         DyDisp.MoveNext
        Next i
    End If
  
    ' work permit
    Set DyDisp = Nothing
    g_Sql = "select * from pr_emp_workpermit_dtl where c_empno = '" & Trim(Txtc_EmpNo) & "'"
    DyDisp.Open g_Sql, CON, adOpenForwardOnly, adLockReadOnly
    If DyDisp.RecordCount > 0 Then
       Txtc_PassNo = Is_Null(DyDisp("c_passno").Value, False)
       Txtd_PassExpDt.Text = Is_Date(DyDisp("d_passexpdt").Value, "D")
       
       Txtc_WrkPermit = Is_Null(DyDisp("c_wrkprmtNo").Value, False)
       Txtd_WrkPrmtExpiry.Text = Is_Date(DyDisp("d_wrkdate").Value, "D")
       
       Txtc_MinRefNo = Is_Null(DyDisp("c_minrefno").Value, False)
       Txtc_GuaranteeNo = Is_Null(DyDisp("c_guaranteeno").Value, False)
       Txtc_GuaranteePassport = Is_Null(DyDisp("c_guaranteepassport").Value, False)
       
       Txtc_Flightno = Is_Null(DyDisp("c_flightno").Value, False)
       Txtd_FlightDate.Text = Is_Date(DyDisp("d_flightdate").Value, "D")
    End If
    
    Call Spread_Row_Height(Va_Salary)
    Call Spread_Row_Height(Va_Leave)
    Call Spread_Row_Height(Va_Remarks)
    Call Spread_Row_Height(Va_Edu)
    Call Spread_Row_Height(Va_Exp)

  Exit Sub

Err_Display:
    If Err > 0 Then
        If Err = 94 Then
          Resume Next
        End If
        MsgBox Err.Description
         Set DyDisp = Nothing
    End If
End Sub

Private Sub Get_Clocking_Employee()
  Dim rsChk As New ADODB.Recordset
  Dim i As Integer
  Dim IsFound As Boolean
  
    Set rsChk = Nothing
    g_Sql = "select c_empno from pr_emp_mst where c_rec_sta = 'A' and c_empno = '" & Trim(Txtc_EmpNo) & "'"
    rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If rsChk.RecordCount > 0 Then
       Exit Sub
    End If
  
  
    Set rsChk = Nothing
    g_Sql = "select employeeid, departmentcode, employeebranchcode, employeefirstname, employeemiddlename, employeelastname, " & _
            "employeeaddress, employeecountry, employeecity, employeetitle, employeegender, employeebirthdate, employeehiredate, employeeoutdate, " & _
            "employeehomephone, employeehandphone, employeeemail " & _
            "From hitfpta.dbo.employee Where employeeid = '" & Trim(Txtc_EmpNo) & "'"
    rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If rsChk.RecordCount > 0 Then
       Txtc_Name = Is_Null(rsChk("employeefirstname").Value, False)
       Txtc_OtherName = Is_Null(rsChk("employeemiddlename").Value, False) & " " & Is_Null(rsChk("employeelastname").Value, False)
       Txtc_Add = Is_Null(rsChk("employeeaddress").Value, False) & vbCrLf & Is_Null(rsChk("employeecity").Value, False)
       Dtp_DOB.Text = Is_Date(Is_Null(rsChk("employeebirthdate").Value, False), "D")
       Dtp_DOJ.Text = Is_Date(Is_Null(rsChk("employeehiredate").Value, False), "D")
       Dtp_DOJ.Text = Is_Date(Is_Null(rsChk("employeeoutdate").Value, False), "D")
       Txtc_PerContNo = Trim(Is_Null(rsChk("employeehomephone").Value, False) & "    " & Is_Null(rsChk("employeehandphone").Value, False))
       Txtc_EMail = Is_Null(rsChk("employeeemail").Value, False)
       Txtc_EmpType = "Worker"
       
       Cmb_Company.ListIndex = 1
       Cmb_Expatriate.ListIndex = 1
       Cmb_StaffType.ListIndex = 1
       Chk_ClockCard.Value = 1
       Cmb_EDFCategory.ListIndex = 1
       Cmb_PayType.ListIndex = 0
       Cmb_DayWork.ListIndex = 0
       Cmb_TpFlag.ListIndex = 1
       Txtc_ShiftCode = "S01"
       Call Txtc_ShiftCode_Validate(True)
       
       If Is_Null(rsChk("employeetitle").Value, False) <> "" Then
          IsFound = False
          For i = 0 To Cmb_Title.ListCount - 1
            If Trim(Cmb_Title.List(i)) = Proper(Is_Null(rsChk("employeetitle").Value, False)) Then
               Cmb_Title.ListIndex = i
               IsFound = True
               Exit For
            End If
          Next i
          If Not IsFound Then
             Cmb_Title.AddItem Proper(Is_Null(rsChk("employeetitle").Value, False))
             Cmb_Title.ListIndex = Cmb_Title.ListCount - 1
          End If
       End If
    
       If Is_Null(rsChk("employeegender").Value, False) <> "" Then
          IsFound = False
          For i = 0 To Cmb_Sex.ListCount - 1
            If Trim(Cmb_Sex.List(i)) = Proper(Is_Null(rsChk("employeegender").Value, False)) Then
               Cmb_Sex.ListIndex = i
               IsFound = True
               Exit For
            End If
          Next i
          If Not IsFound Then
             Cmb_Sex.AddItem Proper(Is_Null(rsChk("employeegender").Value, False))
             Cmb_Sex.ListIndex = Cmb_Sex.ListCount - 1
          End If
       End If
    
       If Is_Null(rsChk("employeebranchcode").Value, False) <> "" Then
          IsFound = False
          For i = 0 To Cmb_Branch.ListCount - 1
            If Trim(Cmb_Branch.List(i)) = Proper(Is_Null(rsChk("employeebranchcode").Value, False)) Then
               Cmb_Branch.ListIndex = i
               IsFound = True
               Exit For
            End If
          Next i
          If Not IsFound Then
             Cmb_Branch.AddItem Proper(Is_Null(rsChk("employeebranchcode").Value, False))
             Cmb_Branch.ListIndex = Cmb_Branch.ListCount - 1
          End If
       End If
    
       If Is_Null(rsChk("departmentcode").Value, False) <> "" Then
          IsFound = False
          For i = 0 To Cmb_Dept.ListCount - 1
            If Trim(Cmb_Dept.List(i)) = Proper(Is_Null(rsChk("departmentcode").Value, False)) Then
               Cmb_Dept.ListIndex = i
               IsFound = True
               Exit For
            End If
          Next i
          If Not IsFound Then
             Cmb_Dept.AddItem Proper(Is_Null(rsChk("departmentcode").Value, False))
             Cmb_Dept.ListIndex = Cmb_Dept.ListCount - 1
          End If
       End If
    End If
End Sub

Private Sub Combo_Load()
Dim rsCombo As New ADODB.Recordset
Dim i As Long
Dim Str As String
    
    ' Title
    Cmb_Title.Clear
    Cmb_Title.AddItem "Mr."
    Cmb_Title.AddItem "Ms."
    Cmb_Title.AddItem "Mrs."
    Set rsCombo = Nothing
    g_Sql = "Select distinct c_title from pr_emp_mst where c_rec_sta='A' and c_title is not null and c_title not in ('Mr.','Ms.','Mrs.') " & _
            "order by c_title"
    rsCombo.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    For i = 0 To rsCombo.RecordCount - 1
        Cmb_Title.AddItem Is_Null(rsCombo("c_title").Value, False)
        rsCombo.MoveNext
    Next i
    
    'Gender
    Cmb_Sex.Clear
    Cmb_Sex.AddItem "Male"
    Cmb_Sex.AddItem "Female"
    Cmb_Sex.AddItem "Transgender"
    Set rsCombo = Nothing
    g_Sql = "Select distinct c_sex from pr_emp_mst where c_rec_sta='A' and c_sex is not null and c_sex not in ('Male','Female','Transgender') " & _
            "order by c_sex"
    rsCombo.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    For i = 0 To rsCombo.RecordCount - 1
        Cmb_Sex.AddItem Is_Null(rsCombo("c_sex").Value, False)
        rsCombo.MoveNext
    Next i
    
    'Expatriate
    Cmb_Expatriate.Clear
    Cmb_Expatriate.AddItem "Yes"
    Cmb_Expatriate.AddItem "No"
    
    
    ' Nationality
    Cmb_Nationality.Clear
    Cmb_Nationality.AddItem "Mauritian"
    Set rsCombo = Nothing
    g_Sql = "Select distinct c_nationality from pr_emp_mst where c_rec_sta='A' and c_nationality is not null and c_nationality not in ('Mauritian') " & _
            "order by c_nationality"
    rsCombo.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    For i = 0 To rsCombo.RecordCount - 1
        Cmb_Nationality.AddItem Is_Null(rsCombo("c_nationality").Value, False)
        rsCombo.MoveNext
    Next i
    
    
    ' Marital status
    Cmb_MatStatus.Clear
    Cmb_MatStatus.AddItem "Single"
    Cmb_MatStatus.AddItem "Married"
    Set rsCombo = Nothing
    g_Sql = "Select distinct c_matstatus from pr_emp_mst where c_rec_sta='A' and c_matstatus is not null and c_matstatus not in ('Single','Married') " & _
            "order by c_matstatus"
    rsCombo.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    For i = 0 To rsCombo.RecordCount - 1
        Cmb_MatStatus.AddItem Is_Null(rsCombo("c_matstatus").Value, False)
        rsCombo.MoveNext
    Next i
    
    Call LoadComboCompany(Me)
    Call LoadComboBranch(Me)
    Call LoadComboDept(Me)
    Call LoadComboDesig(Me)
    
    ' Skill
    Cmb_SkillSet.Clear
    Cmb_SkillSet.AddItem "Skilled"
    Cmb_SkillSet.AddItem "Un-Skilled"
    Set rsCombo = Nothing
    g_Sql = "Select distinct c_skillset from pr_emp_mst where c_rec_sta='A' and c_skillset is not null and c_skillset not in ('Skilled','Un-Skilled') " & _
            "order by c_skillset"
    rsCombo.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    For i = 0 To rsCombo.RecordCount - 1
        Cmb_SkillSet.AddItem Is_Null(rsCombo("c_skillset").Value, False)
        rsCombo.MoveNext
    Next i
      
    Cmb_PayType.Clear
    Cmb_PayType.AddItem "Cash" & Space(50) & "CA"
    Cmb_PayType.AddItem "Cheque" & Space(50) & "CH"
    Cmb_PayType.AddItem "Bank A/c" & Space(50) & "BA"
    
    Cmb_DayWork.Clear
    Cmb_DayWork.AddItem "5 Days" & Space(50) & "5D"
    Cmb_DayWork.AddItem "6 Days" & Space(50) & "6D"
    Cmb_DayWork.AddItem "5 Days - Shift" & Space(50) & "5DS"
    Cmb_DayWork.AddItem "6 Days - Shift" & Space(50) & "6DS"
    
    Cmb_TpFlag.Clear
    Cmb_TpFlag.AddItem "Yes"
    Cmb_TpFlag.AddItem "No"
      
    Cmb_TAType.Clear
    Cmb_TAType.AddItem ""
    Cmb_TAType.AddItem "Actual" & Space(50) & "A"
    Cmb_TAType.AddItem "Fixed" & Space(50) & "F"
    
    Cmb_StaffType.Clear
    Cmb_StaffType.AddItem "Flat" & Space(50) & "F"
    Cmb_StaffType.AddItem "OverTime" & Space(50) & "O"
    
    Cmb_SalaryType.Clear
    Cmb_SalaryType.AddItem "Monthly" & Space(50) & "ML"
    Cmb_SalaryType.AddItem "Hourly" & Space(50) & "HR"
    
  
    ' EDF category
    Set rsCombo = Nothing
    g_Sql = "Select * from pr_edfmast order by c_category"
    rsCombo.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    Cmb_EDFCategory.Clear
    Cmb_EDFCategory.AddItem ""
    For i = 0 To rsCombo.RecordCount - 1
        Cmb_EDFCategory.AddItem Is_Null(rsCombo("c_category").Value, False) & " - " & Is_Null(rsCombo("c_desp").Value, False) & Space(100) & "~" & Is_Null(rsCombo("n_edfamt").Value, True)
        rsCombo.MoveNext
    Next i
    
End Sub

Private Sub Txtc_AcctNo_KeyPress(KeyAscii As Integer)
  '  Call OnlyNumeric(Txtc_AcctNo, KeyAscii, 20)
End Sub

Private Sub Txtc_AcctNo_Validate(Cancel As Boolean)
 Dim rsChk As New ADODB.Recordset
 Dim tmpStr As String
 
 
   If Trim(Txtc_AcctNo) <> "" Then
      If Right(Trim(Txtc_Bank), 3) = "009" And (Trim(Txtc_BankCode) = "" Or Trim(Txtc_BankCode) = "11") And Len(Trim(Txtc_AcctNo)) <> 14 Then
         MsgBox "SBM Account No. should be 14 digit. Please check your entry", vbInformation, "Information"
         Txtc_Bank.SetFocus
         Cancel = True
         Exit Sub
      End If
      
      Set rsChk = Nothing
      g_Sql = "select c_empno, c_name, c_othername, d_dol " & _
              "from pr_emp_mst where c_rec_sta = 'A' and c_acctno='" & Trim(Txtc_AcctNo) & "' and " & _
              "c_empno <> '" & Trim(Txtc_EmpNo) & "'"
      rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
      If rsChk.RecordCount > 0 Then
         tmpStr = "The Account No. is already exists for " & _
                  Is_Null(rsChk("c_empno").Value, False) & Space(2) & Is_Null(rsChk("c_name"), False) & Space(2) & Is_Null(rsChk("c_othername"), False)
         If IsDate(rsChk("d_dol").Value) Then
            tmpStr = tmpStr & ". But who is left on " & Format(rsChk("d_dol").Value, "dd/mm/yyyy")
         End If
         tmpStr = tmpStr & ". Please Check your entry."
         
         MsgBox tmpStr, vbInformation, "Information"
         Txtc_AcctNo.SetFocus
      End If
   End If

End Sub

Private Sub Txtc_Bank_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim Search As New Search.MyClass, SerVar
    If KeyCode = vbKeyDelete Then
       Txtc_Bank = ""
    End If
   
    If KeyCode = vbKeyF2 Then
       Search.Query = "Select c_code Code, c_bankname BankName, c_bankcode BankCode from pr_bankmast where c_rec_sta='A'"
       Search.CheckFields = "Code, BankName, BankCode"
       Search.ReturnField = "Code, BankName, BankCode"
       SerVar = Search.Search(, , CON)
       If Len(Search.col1) <> 0 Then
          Txtc_Bank = Search.col2 & Space(100) & Search.col1
       End If
    End If
End Sub

Private Sub Txtc_Bank_Validate(Cancel As Boolean)
 Dim rsChk As New ADODB.Recordset
   If Trim(Txtc_Bank) <> "" Then
      Set rsChk = Nothing
      g_Sql = "select c_code, c_bankname, c_bankcode from pr_bankmast " & _
              "where c_rec_sta='A' and c_code='" & Trim(Right(Trim(Txtc_Bank), 3)) & "'"
      rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
      If rsChk.RecordCount > 0 Then
         Txtc_Bank = Is_Null(rsChk("c_bankname").Value, False) & Space(100) & Is_Null(rsChk("c_code").Value, False)
         Txtc_BankCode = Is_Null(rsChk("c_bankcode").Value, False)
      Else
         MsgBox "Bank not found in Master. Press <F2> to Select", vbInformation, "Information"
         Txtc_Bank.SetFocus
         Cancel = True
      End If
   End If
End Sub

Private Sub Txtc_ITNo_KeyPress(KeyAscii As Integer)
    Call OnlyNumeric(Txtc_ITNo, KeyAscii, 8)
End Sub

Private Sub Txtc_Line_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim Search As New Search.MyClass, SerVar
    
    If KeyCode = vbKeyF2 Then
       Search.Query = "Select distinct c_line Line from pr_emp_mst where c_rec_sta='A'"
       Search.CheckFields = "Line"
       Search.ReturnField = "Line"
       SerVar = Search.Search(, , CON)
       If Len(Search.col1) <> 0 Then
          Txtc_Line = Search.col1
       End If
    End If
End Sub

Private Sub Txtc_ShiftCode_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim Search As New Search.MyClass, SerVar
 
    If KeyCode = vbKeyDelete Then
       Txtc_ShiftCode = ""
    End If
   
    If KeyCode = vbKeyF2 Then
       Search.Query = "select c_shiftname ShiftName, c_code Code " & _
                      "from pr_shiftstructure_mst "
       Search.CheckFields = "ShiftName, Code"
       Search.ReturnField = "ShiftName, Code"
       SerVar = Search.Search(, , CON)
       If Len(Search.col1) <> 0 Then
          Txtc_ShiftCode = Search.col1 & Space(100) & Search.col2
       End If
    End If
End Sub

Private Sub Txtc_ShiftCode_Validate(Cancel As Boolean)
 Dim rsChk As New ADODB.Recordset
   
   If Trim(Txtc_ShiftCode) <> "" Then
      Set rsChk = Nothing
      g_Sql = "select * from pr_shiftstructure_mst where c_code = '" & Trim(Right(Trim(Txtc_ShiftCode), 3)) & "'"
      rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
      If rsChk.RecordCount > 0 Then
         Txtc_ShiftCode = Is_Null(rsChk("c_shiftname").Value, False) & Space(100) & Is_Null(rsChk("c_code").Value, False)
      Else
         MsgBox "Shift details not found in Master. Press <F2> to Select", vbInformation, "Information"
         Txtc_ShiftCode.SetFocus
         Cancel = True
      End If
   End If
End Sub

Private Sub Txtc_EmpNo_Validate(Cancel As Boolean)
    Call Display_Records
End Sub

Private Sub Txtc_ITNo_Validate(Cancel As Boolean)
 Dim rsChk As New ADODB.Recordset
 Dim tmpStr As String
 
   If Trim(Txtc_ITNo) <> "" Then
      If Len(Trim(Txtc_ITNo)) <> 8 Then
         MsgBox "PAYE No. should be 8 digit. Please check your entry", vbInformation, "Information"
         Txtc_ITNo.SetFocus
         Cancel = True
         Exit Sub
      End If
      
      Set rsChk = Nothing
      g_Sql = "select c_empno, c_name, c_othername, c_itno from pr_emp_mst " & _
              "where c_rec_sta = 'A' and  c_itno = '" & Trim(Txtc_ITNo) & "' and c_empno <> '" & Trim(Txtc_EmpNo) & "'"
      rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
      If rsChk.RecordCount > 0 Then
         tmpStr = "The PAYE No. is already exists for " & _
                  Is_Null(rsChk("c_empno").Value, False) & Space(2) & Is_Null(rsChk("c_name"), False) & Space(2) & Is_Null(rsChk("c_othername"), False)
         MsgBox tmpStr, vbInformation, "Information"
         Txtc_ITNo.SetFocus
         Cancel = True
      End If
   End If
End Sub

Private Sub Display_PayStruc()
  Dim rsPay As New ADODB.Recordset
  Dim i As Integer
    
    If Va_Salary.DataRowCnt = 0 Then
       Set rsPay = Nothing
       g_Sql = "select b.c_salary, b.c_payname from pr_paystructure_mst a, pr_paystructure_dtl b " & _
               "where a.c_company = b.c_company and a.c_rec_sta='A' and b.c_master = 'Y' and " & _
               "a.c_company = '" & Trim(Right(Trim(Cmb_Company), 7)) & "' " & _
               "order by b.n_seq "
       rsPay.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
       For i = 1 To rsPay.RecordCount
           Va_Salary.Row = i
           Va_Salary.Col = 2
              Va_Salary.Text = Is_Null(rsPay("c_payname").Value, False) & Space(100) & Is_Null(rsPay("c_salary").Value, False)
           rsPay.MoveNext
       Next i
    End If
End Sub

Private Sub Dtp_DOJ_Validate(Cancel As Boolean)
  Dim i As Integer
    If IsDate(Dtp_DOJ.Text) Then
       For i = 1 To Va_Salary.DataRowCnt
           Va_Salary.Row = i
           Va_Salary.Col = 5
              If Trim(Va_Salary.Text) <> "O" Then
                 Va_Salary.Col = 1
                    Va_Salary.Text = Is_Date(Dtp_DOJ.Text, "D")
              End If
       Next i
       
'       If Not g_Admin And (CDate(Dtp_DOJ.Text) > (g_CurrentDate + 60) Or CDate(Dtp_DOJ.Text) < (g_CurrentDate - 60)) Then
'          MsgBox "Date Join should be input 60 days either side of current date. Please contact Admin", vbInformation, "Information"
'          Dtp_DOJ.SetFocus
'          Cancel = True
'          Exit Sub
'       End If
    End If
End Sub

Private Sub Dtp_DOL_Validate(Cancel As Boolean)
  Dim rsChk As New ADODB.Recordset
  
  If IsDate(Dtp_DOJ.Text) And IsDate(Dtp_DOL.Text) Then
     If CDate(Dtp_DOJ.Text) > CDate(Dtp_DOL.Text) Then
        MsgBox "Date of Left should not be less than Date of Join"
        Dtp_DOL.SetFocus
        Cancel = True
     End If
     
'     If Not g_Admin And (CDate(Dtp_DOL.Text) > (g_CurrentDate + 60) Or CDate(Dtp_DOL.Text) < (g_CurrentDate - 60)) Then
'        MsgBox "Date left should be input 60 days either side of current date. Please contact Admin", vbInformation, "Information"
'        Dtp_DOL.SetFocus
'        Cancel = True
'        Exit Sub
'     End If
     
     Set rsChk = Nothing
     g_Sql = "select n_loanamount-n_loanpaid n_amount " & _
             "from pr_loan_mst where c_rec_sta = 'A' and n_loanamount - n_loanpaid > 0 and " & _
             "c_empno = '" & Trim(Txtc_EmpNo) & "'"
     rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
     If rsChk.RecordCount > 0 Then
        MsgBox "This employee having Advance balance Rs. " & Is_Null(rsChk("n_amount").Value, True) & "  " & _
               "Please check Advance details. ", vbInformation, "Information"
     End If
  End If
End Sub


Private Sub Txtn_EdfAmt_KeyPress(KeyAscii As Integer)
    Call OnlyNumeric(Txtn_EdfAmt, KeyAscii, 10, 2)
End Sub

Private Sub Txtn_EdfAmt_Validate(Cancel As Boolean)
    Txtn_EdfAmt = Format_Num(Txtn_EdfAmt)
End Sub

Private Sub Txtn_EduAmt_KeyPress(KeyAscii As Integer)
    Call OnlyNumeric(Txtn_EduAmt, KeyAscii, 10, 2)
End Sub

Private Sub Txtn_EduAmt_Validate(Cancel As Boolean)
    Txtn_EduAmt = Format_Num(Txtn_EduAmt)
End Sub

Private Sub Txtn_IntAmt_KeyPress(KeyAscii As Integer)
    Call OnlyNumeric(Txtn_IntAmt, KeyAscii, 10, 2)
End Sub

Private Sub Txtn_IntAmt_Validate(Cancel As Boolean)
    Txtn_IntAmt = Format_Num(Txtn_IntAmt)
End Sub

Private Sub Txtn_PreAmt_KeyPress(KeyAscii As Integer)
    Call OnlyNumeric(Txtn_PreAmt, KeyAscii, 10, 2)
End Sub

Private Sub Txtn_PreAmt_Validate(Cancel As Boolean)
    Txtn_PreAmt = Format_Num(Txtn_PreAmt)
End Sub

Private Sub Txtn_OthAmt_KeyPress(KeyAscii As Integer)
    Call OnlyNumeric(Txtn_OthAmt, KeyAscii, 10, 2)
End Sub

Private Sub Txtn_OthAmt_Validate(Cancel As Boolean)
    Txtn_OthAmt = Format_Num(Txtn_OthAmt)
End Sub

Private Sub Txtn_CarBenefit_KeyPress(KeyAscii As Integer)
    Call OnlyNumeric(Txtn_CarBenefit, KeyAscii, 10, 2)
End Sub

Private Sub Txtn_CarBenefit_Validate(Cancel As Boolean)
    Txtn_CarBenefit = Format_Num(Txtn_CarBenefit)
End Sub

Private Sub Va_Leave_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim Search As New Search.MyClass, SerVar
 
    If ((Shift And 1) = 1) And (KeyCode = vbKeyDelete) Then
       Va_Leave.Row = Va_Leave.ActiveRow
       Va_Leave.Col = 1
       Va_Leave.Action = 5
    End If
End Sub

Private Sub Va_Remarks_KeyDown(KeyCode As Integer, Shift As Integer)
    If ((Shift And 1) = 1) And KeyCode = vbKeyInsert Then
       Call SpreadInsertRow(Va_Remarks, Va_Remarks.ActiveRow)
       Call Spread_Row_Height(Va_Remarks)
    ElseIf ((Shift And 1) = 1) And KeyCode = vbKeyDelete Then
       Call SpreadDeleteRow(Va_Remarks, Va_Remarks.ActiveRow)
    ElseIf KeyCode = vbKeyDelete Then
       Call SpreadCellDataClear(Va_Remarks, Va_Remarks.ActiveRow, Va_Remarks.ActiveCol)
    End If
End Sub

Private Sub Va_Edu_KeyDown(KeyCode As Integer, Shift As Integer)
    If ((Shift And 1) = 1) And KeyCode = vbKeyInsert Then
       Call SpreadInsertRow(Va_Edu, Va_Edu.ActiveRow)
       Call Spread_Row_Height(Va_Edu)
    ElseIf ((Shift And 1) = 1) And KeyCode = vbKeyDelete Then
       Call SpreadDeleteRow(Va_Edu, Va_Edu.ActiveRow)
    ElseIf KeyCode = vbKeyDelete Then
       Call SpreadCellDataClear(Va_Edu, Va_Edu.ActiveRow, Va_Edu.ActiveCol)
    End If
End Sub

Private Sub Va_Exp_KeyDown(KeyCode As Integer, Shift As Integer)
    If ((Shift And 1) = 1) And KeyCode = vbKeyInsert Then
       Call SpreadInsertRow(Va_Exp, Va_Exp.ActiveRow)
       Call Spread_Row_Height(Va_Exp)
    ElseIf ((Shift And 1) = 1) And KeyCode = vbKeyDelete Then
       Call SpreadDeleteRow(Va_Exp, Va_Exp.ActiveRow)
    ElseIf KeyCode = vbKeyDelete Then
       Call SpreadCellDataClear(Va_Exp, Va_Exp.ActiveRow, Va_Exp.ActiveCol)
    End If
End Sub

Private Sub Va_Salary_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim Search As New Search.MyClass, SerVar
 
    If (KeyCode = vbKeyDelete) Then
     Call SpreadCellDataClear(Va_Salary, Va_Salary.ActiveRow, Va_Salary.ActiveCol)
    
    ElseIf Va_Salary.ActiveCol = 2 And KeyCode = vbKeyF2 Then
       Search.Query = "select b.c_payname PayName, b.c_salary Code " & _
                      "from pr_paystructure_mst a, pr_paystructure_dtl b " & _
                      "where a.c_company = b.c_company and a.c_company = '" & Trim(Right(Trim(Cmb_Company), 7)) & "' and " & _
                      "a.c_rec_sta = 'A' and b.c_master = 'Y' "
       Search.CheckFields = "PayName, Code"
       Search.ReturnField = "PayName, Code"
       SerVar = Search.Search(, , CON)
       If Len(Search.col1) <> 0 Then
          Va_Salary.Row = Va_Salary.ActiveRow
          Va_Salary.Col = 2
             Va_Salary.Text = Trim(Search.col1) & Space(100) & Trim(Search.col2)
       End If
    End If
End Sub

Private Sub Va_Salary_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
   Dim i As Long
   Dim tmpAmount As Double
   Dim tmpSalary As String
   
   If Col = 3 Then
      For i = 1 To Va_Salary.DataRowCnt
          Va_Salary.Row = i
          Va_Salary.Col = 1
             If Trim(Va_Salary.Text) = "" Then
                If IsDate(Dtp_DOJ.Text) Then
                   Va_Salary.Text = Is_Date(Dtp_DOJ.Text, "D")
                End If
             End If
          Va_Salary.Col = 2
             tmpSalary = Right(Trim(Va_Salary.Text), 7)
          Va_Salary.Col = 3
             If Trim(Right(Trim(Cmb_SalaryType), 2)) = "HR" And (tmpSalary = "SAL0001" Or tmpSalary = "SAL0002" Or tmpSalary = "SAL0003" Or tmpSalary = "SAL0004" Or tmpSalary = "SAL0005") Then
                tmpAmount = Val(Va_Salary.Text)
                Va_Salary.Col = 4
                   If tmpAmount > 0 Then
                      Va_Salary.Text = Round(tmpAmount * 45, 2)
                   End If
             Else
                Va_Salary.Col = 4
                   Va_Salary.Text = ""
             End If
      Next i
   End If
End Sub

Private Sub Spread_Lock()
  Dim i As Long
    
    For i = 1 To Va_Salary.MaxCols
        Va_Salary.Row = -1
        Va_Salary.Col = i
        If i = 1 Or i = 2 Or i = 4 Or i = 5 Then
           Va_Salary.Lock = True
        Else
           Va_Salary.Lock = False
        End If
    Next i
    
    For i = 1 To Va_Leave.MaxCols
        Va_Leave.Row = -1
        Va_Leave.Col = i
           Va_Leave.Lock = True
    Next i


    For i = 1 To Va_Remarks.MaxCols
        Va_Remarks.Row = -1
        Va_Remarks.Col = i
           Va_Remarks.Lock = False
    Next i
End Sub

Private Sub CancelButtonClick()
    Clear_Controls Me
    Clear_Spread Va_Salary
    Clear_Spread Va_Leave
    Clear_Spread Va_Remarks
    

    Chk_ClockCard.Value = False
    Chk_Disabled.Value = False
    Chk_NoPay.Value = False
    Chk_PayeRelief.Value = False
    Chk_NPFDeduct.Value = False
    Chk_NoBonus.Value = False
       
    Frm_EmpStatus.Enabled = True
    Opt_Active.Enabled = True
    Opt_Left.Enabled = True
End Sub

Private Sub Cmb_PayType_Click()
    If Trim(Right(Trim(Cmb_PayType), 2)) <> "BA" Then
       Txtc_Bank.Text = ""
    End If
End Sub

Private Sub Cmb_StaffType_Click()
    If Trim(Right(Trim(Cmb_StaffType), 2)) = "O" Then
       Chk_ClockCard.Value = 1
    End If
End Sub

Private Sub Cmb_Nationality_Click()
     If Trim(Left(Trim(Cmb_Nationality), 5)) = "Mauri" Then
        Cmb_Expatriate.ListIndex = 1
     Else
        Cmb_Expatriate.ListIndex = 0
     End If
End Sub

Private Function AmountCheck() As Boolean
   Dim i As Integer, tmpAmount As Double
   Dim tmpType As String, tmpFlag As Boolean
   
    tmpFlag = False
    For i = 1 To Va_Salary.DataRowCnt
        Va_Salary.Row = i
        Va_Salary.Col = 2
           tmpType = Trim(Right(Trim(Va_Salary.Text), 7))
        Va_Salary.Col = 3
           tmpAmount = Is_Null_D(Va_Salary.Text, True)
           
           If Trim(Right(Trim(Cmb_SalaryType), 2)) = "ML" Then
              If tmpAmount > 100000 Then tmpFlag = True
           Else
              If tmpType = "SAL0001" Then
                 If tmpAmount > 100 Then tmpFlag = True
              Else
                 If tmpAmount > 10000 Then tmpFlag = True
              End If
           End If
    Next i
    
    If tmpFlag Then
       If MsgBox("Please check Salary / Wages Amount. It seems high. Do you want to Continue?", _
                 vbYesNo + vbDefaultButton2, "Confirmation") = vbYes Then
          AmountCheck = True
          Exit Function
       End If
    Else
       AmountCheck = True
       Exit Function
    End If
End Function

Private Sub Cmb_TpFlag_Click()
  
    If Trim(Cmb_TpFlag) = "Yes" Then
       Txtc_TPMode.Enabled = True
       Txtc_Town.Enabled = True
       Txtc_Road.Enabled = True
       Txtc_PAD.Enabled = True
    
       Txtc_TPMode.BackColor = vbWhite
       Txtc_Town.BackColor = vbWhite
       Txtc_Road.BackColor = vbWhite
       Txtc_PAD.BackColor = vbWhite
    Else
       Txtc_TPMode = ""
       Txtc_Town = ""
       Txtc_Road = ""
       Txtc_PAD = ""
       
       Txtc_TPMode.Enabled = False
       Txtc_Town.Enabled = False
       Txtc_Road.Enabled = False
       Txtc_PAD.Enabled = False
       
       Txtc_TPMode.BackColor = &HE0E0E0
       Txtc_Town.BackColor = &HE0E0E0
       Txtc_Road.BackColor = &HE0E0E0
       Txtc_PAD.BackColor = &HE0E0E0
    End If
End Sub

Private Sub Txtc_TPMode_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim Search As New Search.MyClass, SerVar
    
    If KeyCode = vbKeyF2 Then
       Search.Query = "Select distinct c_tpmode TPMode from pr_emp_mst where c_rec_sta='A'"
       Search.CheckFields = "TPMode"
       Search.ReturnField = "TPMode"
       SerVar = Search.Search(, , CON)
       If Len(Search.col1) <> 0 Then
          Txtc_TPMode = Search.col1
       End If
    End If
End Sub

Private Sub Txtc_Town_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim Search As New Search.MyClass, SerVar
    
    If KeyCode = vbKeyF2 Then
       Search.Query = "Select distinct c_town Town from pr_emp_mst where c_rec_sta='A'"
       Search.CheckFields = "Town"
       Search.ReturnField = "Town"
       SerVar = Search.Search(, , CON)
       If Len(Search.col1) <> 0 Then
          Txtc_Town = Search.col1
       End If
    End If
End Sub

Private Sub Txtc_Road_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim Search As New Search.MyClass, SerVar
    
    If KeyCode = vbKeyF2 Then
       Search.Query = "Select distinct c_road Road from pr_emp_mst where c_rec_sta='A'"
       Search.CheckFields = "Road"
       Search.ReturnField = "Road"
       SerVar = Search.Search(, , CON)
       If Len(Search.col1) <> 0 Then
          Txtc_Road = Search.col1
       End If
    End If
End Sub

Private Sub Txtc_Pad_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim Search As New Search.MyClass, SerVar
    
    If KeyCode = vbKeyF2 Then
       Search.Query = "Select distinct c_pad PAD from pr_emp_mst where c_rec_sta='A'"
       Search.CheckFields = "PAD"
       Search.ReturnField = "PAD"
       SerVar = Search.Search(, , CON)
       If Len(Search.col1) <> 0 Then
          Txtc_PAD = Search.col1
       End If
    End If
End Sub

Private Sub Cmb_DayWork_Validate(Cancel As Boolean)
   If Trim(Cmb_DayWork) <> "" Then
      Call Txtc_ShiftCode_Validate(True)
   End If
End Sub

Private Sub Btn_AddCombo_Click(Index As Integer)
  Dim vDesp As String

    If Index = 0 Then  'Nationality
       vDesp = InputBox("Enter Nationality to be added", "Add Nationality")
       If Trim(vDesp) <> "" Then
          Cmb_Nationality.AddItem Proper(Trim(vDesp))
          Cmb_Nationality.ListIndex = Cmb_Nationality.ListCount - 1
       End If
    ElseIf Index = 1 Then  'Marital status
       vDesp = InputBox("Enter Marital Status to be added", "Add Marital Status")
       If Trim(vDesp) <> "" Then
          Cmb_MatStatus.AddItem Proper(Trim(vDesp))
          Cmb_MatStatus.ListIndex = Cmb_MatStatus.ListCount - 1
       End If
    ElseIf Index = 2 Then  'Branch
       vDesp = InputBox("Enter Branch to be added", "Add Branch")
       If Trim(vDesp) <> "" Then
       Cmb_Branch.AddItem Proper(Trim(vDesp))
       Cmb_Branch.ListIndex = Cmb_Branch.ListCount - 1
       End If
    ElseIf Index = 3 Then  'Department
       vDesp = InputBox("Enter Department to be added", "Add Department")
       If Trim(vDesp) <> "" Then
       Cmb_Dept.AddItem Proper(Trim(vDesp))
       Cmb_Dept.ListIndex = Cmb_Dept.ListCount - 1
       End If
    ElseIf Index = 4 Then  'Designation
       vDesp = InputBox("Enter Designation to be added", "Add Designation")
       If Trim(vDesp) <> "" Then
       Cmb_Desig.AddItem Proper(Trim(vDesp))
       Cmb_Desig.ListIndex = Cmb_Desig.ListCount - 1
       End If

    End If
End Sub

Private Sub Cmb_EDFCategory_Click()
   Dim vVar As Variant
    If Trim(Cmb_EDFCategory) <> "" Then
       vVar = Split(Cmb_EDFCategory, "~")
       Txtn_EdfAmt = Format_Num(Val(vVar(1)))
    End If
End Sub

