VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Object = "{C3A136DA-B937-492B-968D-A437638F7AAB}#1.0#0"; "CS_DateControl.ocx"
Begin VB.Form frm_Report_Filter 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12600
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   12600
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Fme_Period 
      Caption         =   "Filter Option"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   4005
      Left            =   375
      TabIndex        =   31
      Top             =   645
      Width           =   11760
      Begin VB.Frame Frm_Employee 
         Caption         =   "Employee"
         ForeColor       =   &H00C00000&
         Height          =   1500
         Left            =   375
         TabIndex        =   48
         Top             =   1695
         Width           =   4170
         Begin VB.ComboBox Cmb_Sex 
            Height          =   315
            Left            =   1845
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   330
            Width           =   1740
         End
         Begin VB.ComboBox Cmb_Nationality 
            Height          =   315
            Left            =   1845
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   645
            Width           =   1740
         End
         Begin VB.ComboBox Cmb_DayWork 
            Height          =   315
            Left            =   1845
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   975
            Width           =   1740
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   675
            TabIndex        =   51
            Top             =   367
            Width           =   1035
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   675
            TabIndex        =   50
            Top             =   682
            Width           =   1035
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   675
            TabIndex        =   49
            Top             =   1012
            Width           =   1035
         End
      End
      Begin VB.CommandButton Btn_Select 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Select Leave"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   390
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   3510
         Width           =   2610
      End
      Begin VB.TextBox Txtc_Select 
         Height          =   300
         Left            =   3045
         TabIndex        =   26
         Top             =   3525
         Width           =   8250
      End
      Begin VB.Frame Frm_PayType 
         Caption         =   "Payment Type"
         ForeColor       =   &H00C00000&
         Height          =   1500
         Left            =   9675
         TabIndex        =   44
         Top             =   1700
         Width           =   1600
         Begin VB.OptionButton Opt_PayTypeBoth 
            Caption         =   "Both"
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
            Left            =   375
            TabIndex        =   24
            Top             =   975
            Width           =   885
         End
         Begin VB.OptionButton Opt_PayTypeBank 
            Caption         =   "Bank"
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
            Left            =   375
            TabIndex        =   23
            Top             =   660
            Width           =   885
         End
         Begin VB.OptionButton Opt_PayTypeCash 
            Caption         =   "Cash"
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
            TabIndex        =   22
            Top             =   345
            Width           =   885
         End
      End
      Begin VB.Frame Frm_SalaryType 
         Caption         =   "Salary Type"
         ForeColor       =   &H00C00000&
         Height          =   1500
         Left            =   7995
         TabIndex        =   43
         Top             =   1700
         Width           =   1600
         Begin VB.OptionButton Opt_SalaryTypeMon 
            Caption         =   "Monthly"
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
            TabIndex        =   19
            Top             =   345
            Width           =   975
         End
         Begin VB.OptionButton Opt_SalaryTypeHr 
            Caption         =   "Hourly"
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
            Left            =   375
            TabIndex        =   20
            Top             =   660
            Width           =   960
         End
         Begin VB.OptionButton Opt_SalaryTypeBoth 
            Caption         =   "Both"
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
            Left            =   375
            TabIndex        =   21
            Top             =   975
            Width           =   975
         End
      End
      Begin VB.Frame Frm_StaffType 
         Caption         =   "Staff Type"
         ForeColor       =   &H00C00000&
         Height          =   1500
         Left            =   6315
         TabIndex        =   42
         Top             =   1700
         Width           =   1600
         Begin VB.OptionButton Opt_StaffTypeBoth 
            Caption         =   "Both"
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
            Left            =   375
            TabIndex        =   18
            Top             =   975
            Width           =   885
         End
         Begin VB.OptionButton Opt_StaffTypeOt 
            Caption         =   "OT"
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
            Left            =   375
            TabIndex        =   17
            Top             =   660
            Width           =   885
         End
         Begin VB.OptionButton Opt_StaffTypeFlat 
            Caption         =   "Flat"
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
            TabIndex        =   16
            Top             =   345
            Width           =   885
         End
      End
      Begin VB.Frame Frm_EmpStatus 
         Caption         =   "Employee Status"
         ForeColor       =   &H00C00000&
         Height          =   1500
         Left            =   4635
         TabIndex        =   41
         Top             =   1700
         Width           =   1600
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
            Left            =   375
            TabIndex        =   13
            Top             =   345
            Width           =   885
         End
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
            Left            =   375
            TabIndex        =   14
            Top             =   660
            Width           =   885
         End
         Begin VB.OptionButton Opt_BothStatus 
            Caption         =   "Both"
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
            Left            =   375
            TabIndex        =   15
            Top             =   975
            Width           =   885
         End
      End
      Begin VB.ComboBox Cmb_Desig 
         Height          =   315
         Left            =   8850
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   390
         Width           =   2475
      End
      Begin VB.ComboBox Cmb_EmpType 
         Height          =   315
         Left            =   8850
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   720
         Width           =   2460
      End
      Begin VB.ComboBox Cmb_Company 
         Height          =   315
         Left            =   4620
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   390
         Width           =   2505
      End
      Begin VB.ComboBox Cmb_Branch 
         Height          =   315
         Left            =   4620
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   705
         Width           =   2505
      End
      Begin VB.TextBox Txtc_EmployeeName 
         Height          =   300
         Left            =   8850
         TabIndex        =   9
         Top             =   1050
         Width           =   2445
      End
      Begin VB.ComboBox Cmb_Dept 
         Height          =   315
         Left            =   4620
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1020
         Width           =   2520
      End
      Begin VB.TextBox Txtc_Month 
         Height          =   300
         Left            =   1440
         TabIndex        =   0
         Top             =   405
         Width           =   555
      End
      Begin VB.TextBox Txtc_Year 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   2055
         TabIndex        =   1
         Top             =   405
         Width           =   765
      End
      Begin CS_DateControl.DateControl Txtd_FromDate 
         Height          =   345
         Left            =   1440
         TabIndex        =   2
         Top             =   705
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
      End
      Begin CS_DateControl.DateControl Txtd_ToDate 
         Height          =   345
         Left            =   1440
         TabIndex        =   3
         Top             =   1035
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   609
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   30
         X2              =   11700
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   30
         X2              =   11700
         Y1              =   1500
         Y2              =   1500
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Desig"
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
         Height          =   240
         Left            =   7725
         TabIndex        =   40
         Top             =   435
         Width           =   1035
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   7725
         TabIndex        =   39
         Top             =   750
         Width           =   1035
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Company"
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
         Height          =   240
         Left            =   3495
         TabIndex        =   38
         Top             =   420
         Width           =   1035
      End
      Begin VB.Label Lbl_Unit 
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
         Left            =   3960
         TabIndex        =   37
         Top             =   750
         Width           =   570
      End
      Begin VB.Label Lbl_PeriodTo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Period To"
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
         TabIndex        =   36
         Top             =   1095
         Width           =   795
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Employee"
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
         Left            =   7950
         TabIndex        =   35
         Top             =   1080
         Width           =   810
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3495
         TabIndex        =   34
         Top             =   1050
         Width           =   1035
      End
      Begin VB.Label Lbl_PeriodFrom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Period From"
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
         Left            =   315
         TabIndex        =   33
         Top             =   780
         Width           =   1020
      End
      Begin VB.Label Lbl_Period 
         Alignment       =   1  'Right Justify
         Caption         =   "Period"
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
         Height          =   240
         Left            =   765
         TabIndex        =   32
         Top             =   435
         Width           =   570
      End
   End
   Begin VB.Frame Fme_Generate 
      Height          =   1350
      Left            =   375
      TabIndex        =   30
      Top             =   4605
      Width           =   11760
      Begin VB.CommandButton Btn_Exit 
         BackColor       =   &H00C0E0FF&
         Cancel          =   -1  'True
         Caption         =   "&Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   5910
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   540
         Width           =   1155
      End
      Begin VB.CommandButton Btn_View 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Preview"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   4290
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   540
         Width           =   1155
      End
      Begin MSComDlg.CommonDialog comDialog 
         Left            =   285
         Top             =   705
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Frame Frm_Select 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5265
      Left            =   480
      TabIndex        =   45
      Top             =   840
      Width           =   7080
      Begin VB.CommandButton Btn_Ok 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Ok"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2370
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   4800
         Width           =   1710
      End
      Begin FPUSpreadADO.fpSpread Va_Details 
         Height          =   4560
         Left            =   45
         TabIndex        =   46
         Top             =   135
         Width           =   6945
         _Version        =   458752
         _ExtentX        =   12250
         _ExtentY        =   8043
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
         MaxRows         =   20
         ProcessTab      =   -1  'True
         ScrollBars      =   2
         SpreadDesigner  =   "frm_Report_Filter.frx":0000
         VisibleCols     =   1
      End
   End
   Begin VB.Label lbl_scr_name 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Report Filter Options"
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
      Left            =   5250
      TabIndex        =   29
      Top             =   270
      Width           =   1710
   End
   Begin VB.Shape shp_scr_name 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   360
      Top             =   240
      Width           =   11775
   End
End
Attribute VB_Name = "frm_Report_Filter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public RepName As String
Private vPayPeriod As Long
Private vPayPeriodFrom As Date, vPayPeriodTo As Date
Private vF1 As String, vF2 As String, vF3 As String, vF4 As String, vF5 As String
Private SelFor As String, RepTitle As String, RepDate As String, sDate As String
Private vMonthName As String, vSelOption As String

Private Sub Form_Load()
   vSelOption = ""
   Frm_Select.Visible = False
   Btn_Select.Visible = False
   Txtc_Select.Visible = False
   Line2.Visible = False
   
   ' --- SIFB
   Frm_SalaryType.Visible = False
   Frm_PayType.Visible = False
   ' --- SIFB
   
   ' Employee details
   If RepName = "empmstlist" Then
      lbl_scr_name.Caption = "Employee Details"
      Call PeriodDate_DisplayControl(False, False)
      Opt_Active.Value = True
      
   ElseIf RepName = "empmstdoj" Then
      lbl_scr_name.Caption = "New Entry List"
      Call PeriodDate_DisplayControl(False, True)
      Opt_Active.Value = True
   
   ElseIf RepName = "empmstdol" Then
      lbl_scr_name.Caption = "Date Left List"
      Call PeriodDate_DisplayControl(False, True)
      Opt_Active.Value = True
   
   
   ' Leave Details
   ElseIf RepName = "leavedtl" Then
      lbl_scr_name.Caption = "Leave Details"
      Call PeriodDate_DisplayControl(False, True)
      
      vSelOption = "LEAVE"
      Line2.Visible = True
      Opt_Active.Value = True
      Btn_Select.Visible = True
      Btn_Select.Caption = "Select Leave"
      Txtc_Select.Visible = True
      Txtc_Select.Enabled = False
   
   
   ' Loan details
   ElseIf RepName = "loandtl" Then
      lbl_scr_name.Caption = "Loan Details"
      Call PeriodDate_DisplayControl(False, True)
      Opt_Active.Value = True
   
   ElseIf RepName = "loanosdtl" Then
      lbl_scr_name.Caption = "Loan Outstanding Details"
      Call PeriodDate_DisplayControl(False, True, "No")
      Opt_Active.Value = True
   
   
   ' Additional income/deduction Details
   ElseIf RepName = "addincded" Then
      lbl_scr_name.Caption = "Additional Income/Deduction Details"
      Call PeriodDate_DisplayControl(True, False)
   
   ' Mode of Payment Details
   ElseIf RepName = "salpaymodedtl" Then
      lbl_scr_name.Caption = "Salary by Payment Mode"
      Call PeriodDate_DisplayControl(True, False)
      
      vSelOption = "PAYMODE"
      Line2.Visible = True
      Btn_Select.Visible = True
      Btn_Select.Caption = "Select Payment Type"
      Txtc_Select.Visible = True
      Txtc_Select.Enabled = False
   
   ' Pay Component Details
   ElseIf RepName = "salpaycompdtl" Then
      lbl_scr_name.Caption = "Salary Details by Pay Component"
      Call PeriodDate_DisplayControl(True, False)
      
      vSelOption = "PAYCOMP"
      Line2.Visible = True
      Btn_Select.Visible = True
      Btn_Select.Caption = "Select Pay Component"
      Txtc_Select.Visible = True
      Txtc_Select.Enabled = False
   
   ' EDF & PAYE Details
   ElseIf RepName = "saledfpayedtl" Then
      lbl_scr_name.Caption = "List of EDF && PAYE Details"
      Call PeriodDate_DisplayControl(True, False)
      
   ' Notes & Coin Analysis
   ElseIf RepName = "salnotecoin" Then
      lbl_scr_name.Caption = "Notes && Coins Analysis "
      Call PeriodDate_DisplayControl(True, False)
   
   ' Cash Payment Register
   ElseIf RepName = "salcashreg" Then
      lbl_scr_name.Caption = "Cash Payment Register "
      Call PeriodDate_DisplayControl(True, False)
      
      
   ' Social security contributions
   ElseIf RepName = "salsocsecdtl" Then
      lbl_scr_name.Caption = "Contribution Details"
      Call PeriodDate_DisplayControl(True, False)
   
   
   ' PAYE Details
   ElseIf RepName = "salpayedtl" Then
      lbl_scr_name.Caption = "PAYE Details"
      Call PeriodDate_DisplayControl(True, False)
   
   End If
   
   Call Combo_Load
   Call TGControlProperty(Me)
   Cmb_Company.ListIndex = 1
   
   vMonthName = ""
End Sub

Private Sub Btn_View_Click()
  Dim tmpStr As String, RepOpt As String
  
   tmpStr = "": RepOpt = ""
   SelFor = "": RepTitle = "": RepDate = "": sDate = ""
   vF1 = "": vF2 = "": vF3 = "": vF4 = "": vF5 = ""
 
   ' Employee master
   If RepName = "empmstlist" Then
      SelFor = EmpMaster_RepFilter(Me)
      RepTitle = MakeReportHead(Me, "Employee List", True)
      RepDate = MakeReportSubHead(Me)
      Call Print_Rpt(SelFor, "Pr_Emp_List.rpt")
   
   ' Entry (DOJ) details
   ElseIf RepName = "empmstdoj" Then
      SelFor = EmpMaster_RepFilter(Me)
      RepTitle = MakeReportHead(Me, "New Entry List", False)
      RepDate = MakeReportSubHead(Me)
      
      vF1 = "": vF2 = "": vF3 = "": vF4 = "": vF5 = ""
      vF1 = "ISNULL({V_PR_EMP_MST.D_DOL})"
      
      If IsDate(Txtd_FromDate.Text) Then
         vF2 = "{V_PR_EMP_MST.D_DOJ}>=DATETIME(" & Year(Txtd_FromDate.Text) & "," & Format(Month(Txtd_FromDate.Text), "00") & "," & Format(Day(Txtd_FromDate.Text), "00") & ")"
      End If
      If IsDate(Txtd_ToDate.Text) Then
         vF3 = "{V_PR_EMP_MST.D_DOJ}<=DATETIME(" & Year(Txtd_ToDate.Text) & "," & Format(Month(Txtd_ToDate.Text), "00") & "," & Format(Day(Txtd_ToDate.Text), "00") & ")"
      End If
      
      SelFor = ReportFilterOption(SelFor, vF1, vF2, vF3)
      
      If IsDate(Txtd_FromDate.Text) And IsDate(Txtd_ToDate.Text) Then
         RepDate = RepDate & Space(15) & "Date :  " & Format(Txtd_FromDate.Text, "dd/mm/yyyy") & "  to  " & Format(Txtd_ToDate.Text, "dd/mm/yyyy")
      ElseIf IsDate(Txtd_FromDate.Text) Then
         RepDate = RepDate & Space(15) & " As from Date :  " & Format(Txtd_FromDate.Text, "dd/mm/yyyy")
      ElseIf IsDate(Txtd_ToDate.Text) Then
         RepDate = RepDate & Space(15) & " Up to Date :  " & Format(Txtd_ToDate.Text, "dd/mm/yyyy")
      End If
      
      Call Print_Rpt(SelFor, "Pr_Emp_Left_Dtl.rpt")
      
   ' Date Left (DOL)
   ElseIf RepName = "empmstdol" Then
      SelFor = EmpMaster_RepFilter(Me)
      RepTitle = MakeReportHead(Me, "Date Left List", False)
      RepDate = MakeReportSubHead(Me)
      
      vF1 = "": vF2 = "": vF3 = "": vF4 = "": vF5 = ""
      vF1 = "NOT ISNULL({V_PR_EMP_MST.D_DOL})"
      
      If IsDate(Txtd_FromDate.Text) Then
         vF2 = "{V_PR_EMP_MST.D_DOL}>=DATETIME(" & Year(Txtd_FromDate.Text) & "," & Format(Month(Txtd_FromDate.Text), "00") & "," & Format(Day(Txtd_FromDate.Text), "00") & ")"
      End If
      If IsDate(Txtd_ToDate.Text) Then
         vF3 = "{V_PR_EMP_MST.D_DOL}<=DATETIME(" & Year(Txtd_ToDate.Text) & "," & Format(Month(Txtd_ToDate.Text), "00") & "," & Format(Day(Txtd_ToDate.Text), "00") & ")"
      End If
      
      SelFor = ReportFilterOption(SelFor, vF1, vF2, vF3)
      
      If IsDate(Txtd_FromDate.Text) And IsDate(Txtd_ToDate.Text) Then
         RepDate = RepDate & Space(15) & "Date :  " & Format(Txtd_FromDate.Text, "dd/mm/yyyy") & "  to  " & Format(Txtd_ToDate.Text, "dd/mm/yyyy")
      ElseIf IsDate(Txtd_FromDate.Text) Then
         RepDate = RepDate & Space(15) & " As from Date :  " & Format(Txtd_FromDate.Text, "dd/mm/yyyy")
      ElseIf IsDate(Txtd_ToDate.Text) Then
         RepDate = RepDate & Space(15) & " Up to Date :  " & Format(Txtd_ToDate.Text, "dd/mm/yyyy")
      End If
      
      Call Print_Rpt(SelFor, "Pr_Emp_Left_Dtl.rpt")
   
   ' Leave details
   ElseIf RepName = "leavedtl" Then
      If Not IsDate(Txtd_FromDate.Text) Or Not IsDate(Txtd_ToDate.Text) Then
         MsgBox "Please Enter Date", vbInformation, "Information"
         Txtd_FromDate.SetFocus
         Exit Sub
      End If
      
      tmpStr = "1. Leave Balance Details " & vbCrLf & _
               "2. Leave Taken Details " & vbCrLf & _
               "3. Leave Taken Summary " & vbCrLf & vbCrLf & _
               "4. Sick Leave Computation " & vbCrLf & _
               "5. Vacation Leave Computation "
      
      RepOpt = InputBox(tmpStr, "Select Your Option", "1")
      
      If Val(RepOpt) = 0 Then
         Exit Sub
      End If
      
      SelFor = EmpMaster_RepFilter(Me)
      RepDate = MakeReportSubHead(Me)
      
      If Val(RepOpt) = 1 Then
         RepTitle = MakeReportHead(Me, "Leave Balance", False)
         RepDate = RepDate & Space(15) & "Leave Period - " & Trim(Str(Year(Txtd_FromDate.Text)))
         SelFor = Leave_RepFilter(1, SelFor)
         Call Print_Rpt(SelFor, "Pr_Leave_Details.rpt")
         
      ElseIf Val(RepOpt) = 2 Then
         RepTitle = MakeReportHead(Me, "Leave Details", False)
         RepDate = RepDate & Space(15) & "Date :  " & Format(Txtd_FromDate.Text, "dd/mm/yyyy") & "  to  " & Format(Txtd_ToDate.Text, "dd/mm/yyyy")
         SelFor = Leave_RepFilter(2, SelFor)
         Call Print_Rpt(SelFor, "Pr_LeaveTaken_Dtl.rpt")
         
      ElseIf Val(RepOpt) = 3 Then
         RepTitle = MakeReportHead(Me, "Leave Summary", False)
         RepDate = RepDate & Space(15) & "Date :  " & Format(Txtd_FromDate.Text, "dd/mm/yyyy") & "  to  " & Format(Txtd_ToDate.Text, "dd/mm/yyyy")
         SelFor = Leave_RepFilter(3, SelFor)
         Call Print_Rpt(SelFor, "Pr_LeaveTaken_Sum.rpt")
         
      ElseIf Val(RepOpt) = 4 Then
         RepTitle = MakeReportHead(Me, "Sick Leave Computation", False)
         RepDate = RepDate & Space(15) & "Leave Period - " & Trim(Str(Year(Txtd_FromDate.Text)))
         SelFor = Leave_RepFilter(4, SelFor)
         Call Print_Rpt(SelFor, "Pr_Leave_Details_F2.rpt")
      
      ElseIf Val(RepOpt) = 5 Then
         RepTitle = MakeReportHead(Me, "Vacation Leave Computation", False)
         RepDate = RepDate & Space(15) & "Leave Period - " & Trim(Str(Year(Txtd_FromDate.Text)))
         SelFor = Leave_RepFilter(5, SelFor)
         Call Print_Rpt(SelFor, "Pr_Leave_Details_F3.rpt")
      End If
      
   
   ' Loan Details
   ElseIf RepName = "loandtl" Then
      If Not IsDate(Txtd_FromDate.Text) Or Not IsDate(Txtd_ToDate.Text) Then
         MsgBox "Please Enter Date", vbInformation, "Information"
         Txtd_FromDate.SetFocus
         Exit Sub
      End If
      
      tmpStr = "1. Loan Paid List " & vbCrLf & _
               "2. Loan Return List " & vbCrLf & vbCrLf & _
               "3. Loan Details " & vbCrLf & _
               "4. Loan Summary "
      RepOpt = InputBox(tmpStr, "Select Your Option", "1")
      
      If Val(RepOpt) = 0 Then
         Exit Sub
      End If
      
      SelFor = EmpMaster_RepFilter(Me)
      RepDate = MakeReportSubHead(Me)
      RepDate = RepDate & Space(15) & "Date :  " & Format(Txtd_FromDate.Text, "dd/mm/yyyy") & "  to  " & Format(Txtd_ToDate.Text, "dd/mm/yyyy")
      
      If Val(RepOpt) = 1 Then
         RepTitle = MakeReportHead(Me, "Loan Paid List", False)
         Call Print_Rpt(SelFor, "Pr_Loan_Paid_List.rpt")
      ElseIf Val(RepOpt) = 2 Then
         RepTitle = MakeReportHead(Me, "Loan Return List", False)
         Call Print_Rpt(SelFor, "Pr_Loan_Return_List.rpt")
      ElseIf Val(RepOpt) = 3 Then
         RepTitle = MakeReportHead(Me, "Loan Details", False)
         Call Print_Rpt(SelFor, "Pr_Loan_List.rpt")
      ElseIf Val(RepOpt) = 4 Then
         RepTitle = MakeReportHead(Me, "Loan Summary", False)
         Call Print_Rpt(SelFor, "Pr_Loan_Dtl.rpt")
      Else
         Exit Sub
      End If
   
   ' Loan Outstanding Details
   ElseIf RepName = "loanosdtl" Then
      If Not IsDate(Txtd_FromDate.Text) Then
         MsgBox "Please Enter Date", vbInformation, "Information"
         Txtd_FromDate.SetFocus
         Exit Sub
      End If
      
      SelFor = EmpMaster_RepFilter(Me)
      RepTitle = MakeReportHead(Me, "Loan Outstanding", True)
      RepDate = MakeReportSubHead(Me)
     
      g_Sql = "HR_LOAN_OS_UPD '" & Is_Date(Txtd_FromDate.Text, "S") & "'"
      CON.Execute g_Sql
      
      SelFor = SelFor & " AND {PR_LOAN_MST.N_LOANAMOUNT} > {PR_LOAN_MST.N_OSAMOUNT}  "
      Call Print_Rpt(SelFor, "Pr_Loan_OS_List.rpt")
   
   
   ' Additional Income / Deduction Details
   ElseIf RepName = "addincded" Then
      If Not ChkPayPeriod_Entered Then
         Exit Sub
      End If
      
      tmpStr = "1. Details" & vbCrLf & _
               "2. Department Summary " & vbCrLf & _
               "3. Type Summary " & vbCrLf & _
               "4. Summary "
      
      RepOpt = InputBox(tmpStr, "Select Your Option", "1")
      
      If Val(RepOpt) = 0 Then
         Exit Sub
      End If
      
      SelFor = EmpMaster_RepFilter(Me)
      If Trim(SelFor) = "" Then
         vF1 = "{PR_ADDPAY_MST.C_REC_STA}='A' AND {PR_ADDPAY_MST.N_PERIOD}=" & vPayPeriod
      End If
      SelFor = ReportFilterOption(SelFor, vF1)
      RepDate = MakeReportSubHead(Me)
      
      If Val(RepOpt) = 1 Then
         RepTitle = MakeReportHead(Me, "Additional Income/Deduction Details - " & vMonthName & " " & Trim(Txtc_Year), False)
         Call Print_Rpt(SelFor, "Pr_AddIncDed_Dtl.rpt")
      
      ElseIf Val(RepOpt) = 2 Then
         RepTitle = MakeReportHead(Me, "Additional Income/Deduction Dept Summary - " & vMonthName & " " & Trim(Txtc_Year), False)
         Call Print_Rpt(SelFor, "Pr_AddIncDed_Dept_Sum.rpt")
      
      ElseIf Val(RepOpt) = 3 Then
         RepTitle = MakeReportHead(Me, "Additional Income/Deduction Type Summary - " & vMonthName & " " & Trim(Txtc_Year), False)
         Call Print_Rpt(SelFor, "Pr_AddIncDed_Type_Sum.rpt")
      
      ElseIf Val(RepOpt) = 4 Then
         RepTitle = MakeReportHead(Me, "Additional Income/Deduction Summary - " & vMonthName & " " & Trim(Txtc_Year), False)
         Call Print_Rpt(SelFor, "Pr_AddIncDed_Sum.rpt")
      
      Else
         Exit Sub
      End If
      
   ' Mode of Payment Details
   ElseIf RepName = "salpaymodedtl" Then
      If Not ChkPayPeriod_Entered Then
         Exit Sub
      End If
      
      tmpStr = "1. Payment Mode Details " & vbCrLf & _
               "2. Payment Mode Dept Summary " & vbCrLf & _
               "3. Payment Mode Bank Summary"
      RepOpt = InputBox(tmpStr, "Select Your Option", "1")
      
      If Val(RepOpt) = 0 Then
         Exit Sub
      End If
      
      SelFor = SalaryMaster_RepFilter(Me)
      If Trim(Txtc_Select) <> "" Then
         vF1 = "{V_PR_SALARY_MST.C_PAYTYPE} IN ['" & Trim(Right(Trim(Txtc_Select), 200)) & "']"
      End If
      SelFor = ReportFilterOption(SelFor, vF1)
      RepDate = MakeReportSubHead(Me)
      
      If Val(RepOpt) = 1 Then
         RepTitle = MakeReportHead(Me, "Payment Type List - " & vMonthName & " " & Trim(Txtc_Year), False)
         Call Print_Rpt(SelFor, "Pr_Salary_PayMode_Dtl.rpt")
      ElseIf Val(RepOpt) = 2 Then
         RepTitle = MakeReportHead(Me, "Payment Type Summary by Dept - " & vMonthName & " " & Trim(Txtc_Year), False)
         Call Print_Rpt(SelFor, "Pr_Salary_PayMode_Sum.rpt")
      ElseIf Val(RepOpt) = 3 Then
         RepTitle = MakeReportHead(Me, "Payment Type Summary by Bank - " & vMonthName & " " & Trim(Txtc_Year), False)
         Call Print_Rpt(SelFor, "Pr_Salary_PayMode_Bank_Sum.rpt")
      Else
         Exit Sub
      End If
            
   ' Pay Component Details
   ElseIf RepName = "salpaycompdtl" Then
      If Not ChkPayPeriod_Entered Then
         Exit Sub
      End If
      
      tmpStr = "1. Pay Component Details " & vbCrLf & _
               "2. Pay Component Summary"
      RepOpt = InputBox(tmpStr, "Select Your Option", "1")
      
      If Val(RepOpt) = 0 Then
         Exit Sub
      End If
      
      SelFor = SalaryMaster_RepFilter(Me)
      If Trim(Txtc_Select) <> "" Then
         vF1 = "{PR_SALARY_DTL.C_SALARY} IN ['" & Trim(Right(Trim(Txtc_Select), 200)) & "']"
      End If
      SelFor = ReportFilterOption(SelFor, vF1)
      RepDate = MakeReportSubHead(Me)
      
      If Val(RepOpt) = 1 Then
         RepTitle = MakeReportHead(Me, "Salary Pay Component Details - " & vMonthName & " " & Trim(Txtc_Year), False)
         Call Print_Rpt(SelFor, "Pr_Salary_PayComp_Dtl.rpt")
      ElseIf Val(RepOpt) = 2 Then
         RepTitle = MakeReportHead(Me, "Salary Pay Component Summary - " & vMonthName & " " & Trim(Txtc_Year), False)
         Call Print_Rpt(SelFor, "Pr_Salary_PayComp_Sum.rpt")
      Else
         Exit Sub
      End If
            
   ' EDF & PAYE Details
   ElseIf RepName = "saledfpayedtl" Then
      If Not ChkPayPeriod_Entered Then
         Exit Sub
      End If
      
      SelFor = SalaryMaster_RepFilter(Me)
      vF1 = "{PR_SALARY_TOT_DTL.N_PAYEE} > 0"
      SelFor = ReportFilterOption(SelFor, vF1)
      RepDate = MakeReportSubHead(Me)
      
      Call Print_Rpt(SelFor, "Pr_Salary_Paye_Details.rpt")
   
   ' Notes & Coin Analysis
   ElseIf RepName = "salnotecoin" Then
      If Not ChkPayPeriod_Entered Then
         Exit Sub
      End If
      
      tmpStr = "1. Details " & vbCrLf & _
               "2. Branch & Department Summary " & vbCrLf & _
               "3. Summary"
      RepOpt = InputBox(tmpStr, "Select Your Option", "1")
      
      If Val(RepOpt) = 0 Then
         Exit Sub
      End If
      
      SelFor = SalaryMaster_RepFilter(Me)
      RepDate = MakeReportSubHead(Me)
      
      If Val(RepOpt) = 1 Then
         RepTitle = MakeReportHead(Me, "Notes & Coin List - " & vMonthName & " " & Trim(Txtc_Year), False)
         Call Print_Rpt(SelFor, "Pr_Sal_Notes_Dtl.rpt")
      ElseIf Val(RepOpt) = 2 Then
         RepTitle = MakeReportHead(Me, "Notes & Coin by Dept Summary - " & vMonthName & " " & Trim(Txtc_Year), False)
         Call Print_Rpt(SelFor, "Pr_Sal_Notes_Dept_Sum.rpt")
      ElseIf Val(RepOpt) = 3 Then
         RepTitle = MakeReportHead(Me, "Notes & Coin Summary - " & vMonthName & " " & Trim(Txtc_Year), False)
         Call Print_Rpt(SelFor, "Pr_Sal_Notes_Br_Sum.rpt")
      Else
         Exit Sub
      End If
   
   ' Cash Payment Analysis
   ElseIf RepName = "salcashreg" Then
      If Not ChkPayPeriod_Entered Then
         Exit Sub
      End If
      
      SelFor = SalaryMaster_RepFilter(Me)
      RepTitle = MakeReportHead(Me, "Cash Register - " & vMonthName & " " & Trim(Txtc_Year), False)
      RepDate = MakeReportSubHead(Me)
      Call Print_Rpt(SelFor, "Pr_Salary_SigRep.rpt")
            
            
   ' Social Security Details
   ElseIf RepName = "salsocsecdtl" Then
      If Not ChkPayPeriod_Entered Then
         Exit Sub
      End If
    
      tmpStr = "1. Contribution Details" & vbCrLf & _
               "2. Contribution Summary by Dept " & vbCrLf & _
               "3. Contribution Summary "
               
      RepOpt = InputBox(tmpStr, "Select Your Option", "1")
      
      If Val(RepOpt) = 0 Then
         Exit Sub
      End If
      
      SelFor = SalaryMaster_RepFilter(Me)
      RepDate = MakeReportSubHead(Me)
      
      If Val(RepOpt) = 1 Then
         RepTitle = MakeReportHead(Me, "Contribution List - " & vMonthName & " " & Trim(Txtc_Year), False)
         Call Print_Rpt(SelFor, "Pr_Sal_SSec_Dtl.rpt")
      
      ElseIf Val(RepOpt) = 2 Then
         RepTitle = MakeReportHead(Me, "Contribution Summary - " & vMonthName & " " & Trim(Txtc_Year), False)
         Call Print_Rpt(SelFor, "Pr_Sal_SSec_Br_Sum.rpt")
      
      ElseIf Val(RepOpt) = 3 Then
         RepTitle = MakeReportHead(Me, "Contribution Summary - " & vMonthName & " " & Trim(Txtc_Year), False)
         Call Print_Rpt(SelFor, "Pr_Sal_SSec_Sum.rpt")
      
      Else
         Exit Sub
      End If
   
   ' PAYE details
   ElseIf RepName = "salpayedtl" Then
      If Not ChkPayPeriod_Entered Then
         Exit Sub
      End If
      
      tmpStr = "1. Contribution Details" & vbCrLf & _
               "2. Contribution Summary " & vbCrLf & _
               "3. List of PAYE No. Not Alloted "
      
      RepOpt = InputBox(tmpStr, "Select Your Option", "1")
      
      If Val(RepOpt) = 0 Then
         Exit Sub
      End If
      
      SelFor = SalaryMaster_RepFilter(Me)
      SelFor = SelFor & " {V_PR_SALARY_MST.N_PAYE} <> 0 "
      RepDate = MakeReportSubHead(Me)
      
      If Val(RepOpt) = 1 Then
         RepTitle = MakeReportHead(Me, "PAYE Contribution List - " & vMonthName & " " & Trim(Txtc_Year), False)
         Call Print_Rpt(SelFor, "Pr_Sal_Payee_Dtl.rpt")
      ElseIf Val(RepOpt) = 2 Then
         RepTitle = MakeReportHead(Me, "PAYE Contribution Summary - " & vMonthName & " " & Trim(Txtc_Year), False)
         Call Print_Rpt(SelFor, "Pr_Sal_Payee_Sum.rpt")
      ElseIf Val(RepOpt) = 3 Then
         RepTitle = MakeReportHead(Me, "No PAYE Number List - " & vMonthName & " " & Trim(Txtc_Year), False)
         SelFor = SelFor & " AND ISNULL({PR_EMP_MST.C_ITNO}) "
         Call Print_Rpt(SelFor, "Pr_Sal_Payee_Dtl.rpt")
      Else
         Exit Sub
      End If
      
   End If
  
   If Trim(RepTitle) <> "" Then
      Mdi_Ta_HrPay.CRY1.Formulas(1) = "ReportHead='" & UCase(Trim(RepTitle)) & "'"
   End If

   If Trim(RepDate) <> "" Then
      Mdi_Ta_HrPay.CRY1.Formulas(2) = "RepHeadDate='" & Trim(RepDate) & "'"
   End If
   Mdi_Ta_HrPay.CRY1.Action = 1
  
End Sub

Private Sub Btn_Exit_Click()
   Unload Me
End Sub


Private Sub Combo_Load()
  Dim rsCombo As New ADODB.Recordset
  Dim i As Integer
  
    Call LoadComboCompany(Me)
    Call LoadComboBranch(Me)
    Call LoadComboDept(Me)
    Call LoadComboDesig(Me)
    Call LoadComboEmpType(Me)

    'Gender
    Cmb_Sex.Clear
    Cmb_Sex.AddItem ""
    Cmb_Sex.AddItem "Male"
    Cmb_Sex.AddItem "Female"
    Cmb_Sex.AddItem "Transgender"
    Set rsCombo = Nothing
    g_Sql = "Select distinct c_sex from pr_emp_mst where c_rec_sta='A' and c_sex is not null and c_sex not in ('Male','Female','Transgender') " & _
            "order by c_sex"
    rsCombo.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    For i = 0 To rsCombo.RecordCount - 1
        If Is_Null(rsCombo("c_sex").Value, False) <> "" Then
           Cmb_Sex.AddItem Is_Null(rsCombo("c_sex").Value, False)
        End If
        rsCombo.MoveNext
    Next i
    
    
    ' Nationality
    Cmb_Nationality.Clear
    Cmb_Nationality.AddItem ""
    Cmb_Nationality.AddItem "Mauritian"
    Set rsCombo = Nothing
    g_Sql = "Select distinct c_nationality from pr_emp_mst where c_rec_sta='A' and c_nationality is not null and c_nationality not in ('Mauritian') " & _
            "order by c_nationality"
    rsCombo.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    For i = 0 To rsCombo.RecordCount - 1
        If Is_Null(rsCombo("c_nationality").Value, False) <> "" Then
           Cmb_Nationality.AddItem Is_Null(rsCombo("c_nationality").Value, False)
        End If
        rsCombo.MoveNext
    Next i

    Cmb_DayWork.Clear
    Cmb_DayWork.AddItem "5 Days" & Space(50) & "5D"
    Cmb_DayWork.AddItem "6 Days" & Space(50) & "6D"
    Cmb_DayWork.AddItem "5 Days - Shift" & Space(50) & "5DS"
    Cmb_DayWork.AddItem "6 Days - Shift" & Space(50) & "6DS"

End Sub

Private Sub Btn_Select_Click()
    If vSelOption = "" Then
       Exit Sub
    End If
    
    If vSelOption = "LEAVE" Then
       Call Load_Leave_Mst
    ElseIf vSelOption = "PAYMODE" Then
       Call Load_Payment_Mode
    ElseIf vSelOption = "PAYMODE" Then
       Call Load_Pay_Components
    End If
    
    Call SpreadHeaderFont(Va_Details, "Arial", 7, False)
    Call Spread_Row_Height(Va_Details, 14, 15)
    Frm_Select.ZOrder (0)
    Frm_Select.Visible = True
End Sub


Private Sub Load_Leave_Mst()
  Dim rsChk As New ADODB.Recordset
  Dim i As Integer
 
    Set rsChk = Nothing
    g_Sql = "select c_leave, c_leavename from pr_leave_mst where c_rec_sta = 'A' order by c_leavename "
    rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    Va_Details.MaxRows = rsChk.RecordCount
    For i = 1 To rsChk.RecordCount
        Va_Details.Row = i
        Va_Details.Col = 2
           Va_Details.Text = Is_Null(rsChk("c_leavename").Value, False)
        Va_Details.Col = 3
           Va_Details.Text = Is_Null(rsChk("c_leave").Value, False)
        rsChk.MoveNext
    Next i
End Sub

Private Sub Load_Payment_Mode()
  Dim i As Integer
 
    For i = 1 To 3
        Va_Details.Row = i
        Va_Details.Col = 2
           If i = 1 Then
              Va_Details.Text = "Cash"
           ElseIf i = 2 Then
              Va_Details.Text = "Cheque"
           ElseIf i = 3 Then
              Va_Details.Text = "Bank"
           End If
        Va_Details.Col = 3
           If i = 1 Then
              Va_Details.Text = "Cash"
           ElseIf i = 2 Then
              Va_Details.Text = "Cheque"
           ElseIf i = 3 Then
              Va_Details.Text = "Bank"
           End If
    Next i
End Sub

Private Sub Load_Pay_Components()
  Dim rsChk As New ADODB.Recordset
  Dim i As Integer
 
    Set rsChk = Nothing
    g_Sql = "Select c_payname, c_salary From pr_paystructure_Dtl Where c_company = 'COM0001' order by n_seq "
    rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    Va_Details.MaxRows = rsChk.RecordCount
    For i = 1 To rsChk.RecordCount
        Va_Details.Row = i
        Va_Details.Col = 2
           Va_Details.Text = Is_Null(rsChk("c_payname").Value, False)
        Va_Details.Col = 3
           Va_Details.Text = Is_Null(rsChk("c_salary").Value, False)
        rsChk.MoveNext
    Next i
End Sub

Private Sub Btn_Ok_Click()
  Dim tmpStr As String, tmpStrCode As String
  Dim i As Integer
  
    For i = 1 To Va_Details.DataRowCnt
        Va_Details.Row = i
        Va_Details.Col = 1
           If Va_Details.Value = True Then
              Va_Details.Col = 2
              If Trim(tmpStr) = "" Then
                 tmpStr = Proper(Trim(Va_Details.Text))
              Else
                 tmpStr = tmpStr & ", " & Proper(Trim(Va_Details.Text))
              End If
              
              Va_Details.Col = 3
              If Trim(tmpStrCode) = "" Then
                 tmpStrCode = UCase(Trim(Va_Details.Text))
              Else
                 tmpStrCode = tmpStrCode & "', '" & UCase(Trim(Va_Details.Text))
              End If
           End If
    Next i
    Txtc_Select.Text = tmpStr & Space(250) & tmpStrCode
    Frm_Select.Visible = False
End Sub

Private Sub Txtc_EmployeeName_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim Search As New Search.MyClass, SerVar
 
   If KeyCode = vbKeyDelete Then
      Txtc_EmployeeName = ""
   End If
   
   If KeyCode = vbKeyF2 Then
      Search.Query = "select c_empno EmpNo, c_name Name, c_othername OtherName, c_dept Dept, " & _
                     "c_desig Desig, c_branch Branch, c_emptype Type " & _
                     "from pr_emp_mst where c_rec_sta = 'A' and c_name like ('" & Trim(Left(Trim(Txtc_EmployeeName), 10)) & "%')"
      Search.CheckFields = "EmpNo, Name"
      Search.ReturnField = "EmpNo, Name"
      SerVar = Search.Search(, , CON)
      If Len(Search.col1) <> 0 Then
         Txtc_EmployeeName = Search.col2 & Space(100) & Search.col1
      End If
   End If
End Sub

Private Sub Txtc_EmployeeName_Validate(Cancel As Boolean)
 Dim rsChk As New ADODB.Recordset
 Dim i As Integer
 
  If Trim(Txtc_EmployeeName) <> "" Then
     Set rsChk = Nothing
     g_Sql = "select c_empno, c_name, c_othername, c_branch, c_dept from pr_emp_mst " & _
             "where c_rec_sta = 'A' and c_empno = '" & Trim(Right(Trim(Txtc_EmployeeName), 7)) & "'"
     rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
     If rsChk.RecordCount > 0 Then
        Txtc_EmployeeName = Is_Null(rsChk("c_name").Value, False) & " " & Is_Null(rsChk("c_othername").Value, False) & Space(100) & rsChk("c_empno").Value
     Else
        MsgBox "Employee not found. Press <F2> to select.", vbInformation, "Information"
        Cancel = True
     End If
  End If
End Sub

Private Sub Txtc_Month_KeyPress(KeyAscii As Integer)
    Call OnlyNumeric(Txtc_Month, KeyAscii, 2)
End Sub

Private Sub Txtc_Month_Validate(Cancel As Boolean)
    If Trim(Txtc_Month) <> "" Then
       Call MakeMonthTwoDigits(Me)
       If (Val(Txtc_Month) <= 0 Or Val(Txtc_Month) > 13) Then
          MsgBox "Not a valid month", vbInformation, "Information"
          Txtc_Month.SetFocus
          Cancel = True
          Exit Sub
       End If
       
       If Val(Txtc_Month) = 13 Then
          vMonthName = "EOY Bonus"
       Else
          vMonthName = MonthName(Val(Txtc_Month))
       End If
    End If
    
    If Trim(Txtc_Month) <> "" And Trim(Txtc_Year) <> "" Then
       Call Assign_PayPeriodDate
       Txtd_FromDate.Text = Is_Date(vPayPeriodFrom, "D")
       Txtd_ToDate.Text = Is_Date(vPayPeriodTo, "D")
    End If
End Sub

Private Sub txtc_year_KeyPress(KeyAscii As Integer)
    Call OnlyNumeric(Txtc_Year, KeyAscii, 4)
End Sub

Private Sub txtc_year_Validate(Cancel As Boolean)
  If Trim(Txtc_Year) <> "" Then
     If Len(Txtc_Year) <> 4 Then
        MsgBox "Not a valid year", vbInformation, "Information"
        Txtc_Year.SetFocus
        Cancel = True
     End If
  End If
  
  If Trim(Txtc_Month) <> "" And Trim(Txtc_Year) <> "" Then
     Call Assign_PayPeriodDate
     Txtd_FromDate.Text = Is_Date(vPayPeriodFrom, "D")
     Txtd_ToDate.Text = Is_Date(vPayPeriodTo, "D")
  End If
End Sub

Private Sub Assign_PayPeriodDate()
  Dim rsChk As New ADODB.Recordset
  
    vPayPeriod = Is_Null(Format(Trim(Txtc_Year), "0000") & Trim(Txtc_Month), True)
     
    Set rsChk = Nothing
    g_Sql = "select d_fromdate, d_todate from pr_payperiod_dtl " & _
            "where n_period = " & vPayPeriod & " and c_type = 'W' "
    rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If rsChk.RecordCount > 0 Then
       vPayPeriodFrom = rsChk("d_fromdate").Value
       vPayPeriodTo = rsChk("d_todate").Value
    End If
End Sub


Private Function Leave_RepFilter(ByVal vRepNo As Integer, ByVal vExistingFilters As String) As String
  Dim vF1 As String, vF2 As String, vF3 As String, vF4 As String, vF5 As String
  Dim SelFor As String
        
     vF1 = "": vF2 = "": vF3 = "": vF4 = "": vF5 = ""
     SelFor = vExistingFilters
     
     If vRepNo = 1 Or vRepNo = 4 Or vRepNo = 5 Then
         If IsDate(Txtd_FromDate.Text) Then
            vF1 = "YEAR({PR_EMP_LEAVE_DTL.D_PRFROM}) = " & Year(Txtd_FromDate.Text)
         End If
     Else
        If IsDate(Txtd_FromDate.Text) And IsDate(Txtd_ToDate.Text) Then
           vF1 = "{PR_CLOCK_EMP.D_DATE}>=DATETIME(" & Year(Txtd_FromDate.Text) & "," & Format(Month(Txtd_FromDate.Text), "00") & "," & Format(Day(Txtd_FromDate.Text), "00") & ")"
           vF1 = Trim(vF1) & " AND {PR_CLOCK_EMP.D_DATE}<=DATETIME(" & Year(Txtd_ToDate.Text) & "," & Format(Month(Txtd_ToDate.Text), "00") & "," & Format(Day(Txtd_ToDate.Text), "00") & ")"
        End If
     End If
     
     vF2 = "NOT ({PR_LEAVE_MST.C_LEAVE} IN ['P','WO'])"
        
     If vRepNo = 4 Then
        vF3 = "{PR_LEAVE_MST.C_LEAVE} = 'SL'"
     ElseIf vRepNo = 5 Then
        vF3 = "{PR_LEAVE_MST.C_LEAVE} = 'VL'"
     Else
        If Trim(Txtc_Select) <> "" Then
           vF3 = "{PR_LEAVE_MST.C_LEAVE} IN ['" & Trim(Right(Trim(Txtc_Select), 200)) & "']"
        End If
     End If
     
     Leave_RepFilter = ReportFilterOption(SelFor, vF1, vF2, vF3)

End Function

Private Sub PeriodDate_DisplayControl(ByVal vPeriodVisible As Boolean, ByVal vDateVisible As Boolean, Optional ByVal vToDateFlag As String)
    If Not vPeriodVisible And Not vDateVisible Then
       Txtc_Month.Enabled = False
       Txtc_Year.Enabled = False
       Txtd_FromDate.Enabled = False
       Txtd_ToDate.Enabled = False
    End If
    
    If vPeriodVisible And Not vDateVisible Then
       Lbl_Period.Top = Lbl_Period.Top + 300
       Txtc_Month.Top = Txtc_Month.Top + 300
       Txtc_Year.Top = Txtc_Year.Top + 300
    
       Lbl_PeriodFrom.Visible = False
       Txtd_FromDate.Visible = False
       Lbl_PeriodTo.Visible = False
       Txtd_ToDate.Visible = False
    End If
       
    If Not vPeriodVisible And vDateVisible Then
       Lbl_Period.Visible = False
       Txtc_Month.Visible = False
       Txtc_Year.Visible = False
       
       Lbl_PeriodFrom.Top = Lbl_PeriodFrom.Top - 150
       Lbl_PeriodTo.Top = Lbl_PeriodTo.Top - 150
       Txtd_FromDate.Top = Txtd_FromDate.Top - 150
       Txtd_ToDate.Top = Txtd_ToDate.Top - 150
    
       If vToDateFlag = "No" Then
          Lbl_PeriodFrom.Caption = "Date"
          Lbl_PeriodTo.Visible = False
          Txtd_ToDate.Visible = False
       End If
    
    End If
        
End Sub

Private Function ChkPayPeriod_Entered() As Boolean
    If vPayPeriod = 0 Then
       MsgBox "Period should not be blank", vbInformation, "Information"
       Txtc_Month.SetFocus
       Exit Function
    End If
    ChkPayPeriod_Entered = True
End Function
