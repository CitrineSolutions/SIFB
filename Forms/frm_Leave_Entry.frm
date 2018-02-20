VERSION 5.00
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Object = "{C3A136DA-B937-492B-968D-A437638F7AAB}#1.0#0"; "CS_DateControl.ocx"
Begin VB.Form frm_Leave_Entry 
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   11055
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin VB.Frame Frm_Filter 
      Caption         =   "Additional Filter Option"
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
      Height          =   3630
      Left            =   10815
      TabIndex        =   49
      Top             =   1515
      Width           =   7785
      Begin VB.ComboBox Cmb_Type 
         Height          =   315
         Left            =   1275
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2310
         Width           =   2370
      End
      Begin VB.ComboBox Cmb_Leave 
         Height          =   315
         Left            =   1275
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1980
         Width           =   2370
      End
      Begin VB.ComboBox Cmb_EmpType 
         Height          =   315
         Left            =   5460
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   960
         Width           =   1815
      End
      Begin VB.ComboBox Cmb_Shift 
         Height          =   315
         Left            =   5460
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1290
         Width           =   1815
      End
      Begin VB.ComboBox Cmb_Branch 
         Height          =   315
         Left            =   1275
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   960
         Width           =   2370
      End
      Begin VB.ComboBox Cmb_Dept 
         Height          =   315
         Left            =   1275
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1290
         Width           =   2370
      End
      Begin VB.ComboBox Cmb_Company 
         Height          =   315
         Left            =   1275
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   630
         Width           =   2370
      End
      Begin VB.ComboBox Cmb_Desig 
         Height          =   315
         Left            =   5460
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   630
         Width           =   1815
      End
      Begin VB.CommandButton Btn_Cancel 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4410
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3060
         Width           =   1575
      End
      Begin VB.CommandButton Btn_Ok 
         BackColor       =   &H00C0E0FF&
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2340
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3060
         Width           =   1575
      End
      Begin CS_DateControl.DateControl Dtp_FromDate 
         Height          =   345
         Left            =   5460
         TabIndex        =   8
         Top             =   1980
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   609
      End
      Begin CS_DateControl.DateControl Dtp_ToDate 
         Height          =   345
         Left            =   5460
         TabIndex        =   9
         Top             =   2310
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   609
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000005&
         X1              =   0
         X2              =   13720
         Y1              =   2850
         Y2              =   2850
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         X1              =   0
         X2              =   13720
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Type"
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
         Left            =   750
         TabIndex        =   59
         Top             =   2385
         Width           =   405
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "From"
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
         Left            =   4920
         TabIndex        =   58
         Top             =   2040
         Width           =   435
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "To"
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
         Left            =   5145
         TabIndex        =   57
         Top             =   2370
         Width           =   210
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Leave"
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
         Left            =   675
         TabIndex        =   56
         Top             =   2070
         Width           =   495
      End
      Begin VB.Label Label15 
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
         ForeColor       =   &H00404040&
         Height          =   210
         Left            =   4485
         TabIndex        =   55
         Top             =   1005
         Width           =   870
      End
      Begin VB.Label Label14 
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
         ForeColor       =   &H00404040&
         Height          =   210
         Left            =   4920
         TabIndex        =   54
         Top             =   1335
         Width           =   435
      End
      Begin VB.Label Label13 
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
         ForeColor       =   &H00404040&
         Height          =   210
         Left            =   585
         TabIndex        =   53
         Top             =   1005
         Width           =   570
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Dept"
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
         Left            =   780
         TabIndex        =   52
         Top             =   1350
         Width           =   375
      End
      Begin VB.Label Company 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00404040&
         Height          =   210
         Left            =   375
         TabIndex        =   51
         Top             =   690
         Width           =   780
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00404040&
         Height          =   210
         Left            =   4890
         TabIndex        =   50
         Top             =   675
         Width           =   465
      End
   End
   Begin FPUSpreadADO.fpSpread Va_Leave 
      Height          =   6225
      Left            =   12885
      TabIndex        =   34
      Top             =   2325
      Width           =   4170
      _Version        =   458752
      _ExtentX        =   7355
      _ExtentY        =   10980
      _StockProps     =   64
      AutoClipboard   =   0   'False
      ColsFrozen      =   1
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
      MaxRows         =   49
      ProcessTab      =   -1  'True
      SpreadDesigner  =   "frm_Leave_Entry.frx":0000
   End
   Begin VB.Frame Fra_Leave 
      Height          =   1035
      Left            =   12900
      TabIndex        =   35
      Top             =   1290
      Width           =   4170
      Begin VB.TextBox Txtn_Close 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   3495
         MaxLength       =   10
         TabIndex        =   41
         Top             =   630
         Width           =   600
      End
      Begin VB.TextBox Txtn_Adj 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   2820
         MaxLength       =   10
         TabIndex        =   40
         Top             =   630
         Width           =   600
      End
      Begin VB.TextBox Txtn_Used 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   2130
         MaxLength       =   10
         TabIndex        =   39
         Top             =   630
         Width           =   600
      End
      Begin VB.TextBox Txtn_Alloted 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   38
         Top             =   630
         Width           =   600
      End
      Begin VB.TextBox Txtn_Entitle 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   750
         MaxLength       =   10
         TabIndex        =   37
         Top             =   630
         Width           =   600
      End
      Begin VB.TextBox Txtn_Open 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   75
         MaxLength       =   10
         TabIndex        =   36
         Top             =   630
         Width           =   600
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   2820
         TabIndex        =   48
         Top             =   420
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Used"
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
         Left            =   2145
         TabIndex        =   47
         Top             =   420
         Width           =   420
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Close"
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
         Left            =   3495
         TabIndex        =   46
         Top             =   420
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Earned"
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
         Left            =   1455
         TabIndex        =   45
         Top             =   420
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Entitle"
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
         Left            =   750
         TabIndex        =   44
         Top             =   420
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Open"
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
         TabIndex        =   43
         Top             =   420
         Width           =   435
      End
      Begin VB.Label Lbl_EmpDtl 
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
         ForeColor       =   &H00800080&
         Height          =   210
         Left            =   90
         TabIndex        =   42
         Top             =   135
         Width           =   465
      End
   End
   Begin VB.Frame Frame2 
      Height          =   900
      Left            =   120
      TabIndex        =   31
      Top             =   -30
      Width           =   16950
      Begin VB.TextBox Txtc_No 
         Height          =   300
         Left            =   6480
         TabIndex        =   18
         Top             =   240
         Width           =   1230
      End
      Begin VB.CommandButton Btn_Print 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   3435
         Picture         =   "frm_Leave_Entry.frx":043C
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Delete 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   2715
         Picture         =   "frm_Leave_Entry.frx":3A2A
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Clear 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   1725
         Picture         =   "frm_Leave_Entry.frx":70D4
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Exit 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   4470
         Picture         =   "frm_Leave_Entry.frx":A744
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Save 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   30
         Picture         =   "frm_Leave_Entry.frx":DDA4
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_View 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   1005
         Picture         =   "frm_Leave_Entry.frx":1142D
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   700
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "No."
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
         Left            =   6180
         TabIndex        =   33
         Top             =   270
         Width           =   255
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "(Press <F2> to view closed periods)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   210
         Left            =   6465
         TabIndex        =   32
         Top             =   555
         Width           =   2955
      End
   End
   Begin VB.Frame fra_crddeb 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   120
      TabIndex        =   25
      Top             =   1275
      Width           =   12765
      Begin VB.CommandButton Btn_GetEmp 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Get Employee"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   10410
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   255
         Width           =   1440
      End
      Begin VB.TextBox Txtc_Year 
         Height          =   300
         Left            =   4680
         MaxLength       =   4
         TabIndex        =   21
         Top             =   285
         Width           =   1125
      End
      Begin VB.TextBox Txtc_Month 
         Height          =   300
         Left            =   4050
         MaxLength       =   2
         TabIndex        =   20
         Top             =   285
         Width           =   600
      End
      Begin VB.TextBox Txtc_Remarks 
         Height          =   315
         Left            =   4050
         MaxLength       =   50
         TabIndex        =   22
         Top             =   585
         Width           =   3810
      End
      Begin VB.TextBox Txtc_Code 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   1110
         TabIndex        =   19
         Top             =   285
         Width           =   1320
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Remarks"
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
         Left            =   3255
         TabIndex        =   30
         Top             =   630
         Width           =   750
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
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
         Height          =   210
         Left            =   3435
         TabIndex        =   29
         Top             =   330
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   630
         TabIndex        =   28
         Top             =   330
         Width           =   435
      End
   End
   Begin FPUSpreadADO.fpSpread Va_Details 
      Height          =   6225
      Left            =   120
      TabIndex        =   24
      Top             =   2325
      Width           =   12750
      _Version        =   458752
      _ExtentX        =   22490
      _ExtentY        =   10980
      _StockProps     =   64
      AutoClipboard   =   0   'False
      ColsFrozen      =   1
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
      MaxCols         =   10
      MaxRows         =   50
      ProcessTab      =   -1  'True
      SpreadDesigner  =   "frm_Leave_Entry.frx":14ADD
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
      Left            =   15075
      TabIndex        =   27
      Top             =   990
      Width           =   360
   End
   Begin VB.Label lbl_scr_name 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Leave Entry Details"
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
      Left            =   270
      TabIndex        =   26
      Top             =   1005
      Width           =   1560
   End
   Begin VB.Shape shp_scr_name 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   120
      Top             =   945
      Width           =   16935
   End
End
Attribute VB_Name = "frm_Leave_Entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As New ADODB.Recordset
Dim vPeriod As Long
Dim vPayPeriodFrom As Date, vPayPeriodTo As Date

Private Sub Form_Activate()
    Txtc_Month.SetFocus
End Sub

Private Sub Form_Load()
    lbl_date.Caption = Format(Date, "dd-mmm-yyyy")
    lbl_scr_name.Caption = "Leave Entry Details"
    
    Clear_Spread Va_Details
    Clear_Spread Va_Leave
    
    Call Spread_Lock
    Call Load_Combo
    Call TGControlProperty(Me)
    Call Spread_Row_Height(Va_Details)
    Call Spread_Row_Height(Va_Leave)
    
    Frm_Filter.Visible = False
        
    Txtc_Code.Enabled = False
    Txtn_Open.Enabled = False
    Txtn_Entitle.Enabled = False
    Txtn_Alloted.Enabled = False
    Txtn_Used.Enabled = False
    Txtn_Adj.Enabled = False
    Txtn_Close.Enabled = False
End Sub

Private Sub Btn_Exit_Click()
    Unload Me
End Sub

Private Sub Btn_Clear_Click()
    Call CancelButtonClick
End Sub

Private Sub Btn_Delete_Click()
On Error GoTo ErrDel
   Dim rsChk As New ADODB.Recordset
   
    If Trim(Txtc_Code) = "" Then
       Exit Sub
    End If
    
    If ChkPeriodOpen(vPeriod, "W") Then
       If (MsgBox("Are you sure you want to delete ?", vbExclamation + vbYesNo, "Caution") = vbYes) Then
          CON.BeginTrans
          CON.Execute "update pr_leaveentry_mst set " & GetDelFlag & " where c_code = '" & Trim(Txtc_Code) & "'"
          CON.CommitTrans
       End If
       Call CancelButtonClick
    End If
   
   Exit Sub

ErrDel:
    CON.RollbackTrans
    MsgBox "Error while Deleting - " + Err.Description, vbCritical, "Critical"
End Sub

Private Sub Btn_Print_Click()
On Error GoTo Err_Print
  Dim SelFor As String
  
    If Trim(Txtc_Code) = "" Then
       Exit Sub
    End If
  
    SelFor = "{PR_LEAVEENTRY_MST.C_CODE}='" & Trim(Txtc_Code) & "'"
    Call Print_Rpt(SelFor, "Pr_LeaveEntry_Doc.rpt")
    Mdi_Ta_HrPay.CRY1.Action = 1
  
  Exit Sub

Err_Print:
    MsgBox "Error while Generating - " & Err.Description, vbInformation, "Information"
End Sub

Private Sub Btn_View_Click()
  Dim Search As New Search.MyClass, SerVar

    Search.Query = "select a.n_period Period, a.c_remarks Remarks, a.c_code Code " & _
                   "from pr_leaveentry_mst a, pr_payperiod_dtl b " & _
                   "where a.n_period = b.n_period and b.c_type = 'W' and b.c_period_closed = 'N' and a.c_rec_sta='A' "
    Search.CheckFields = "Code"
    Search.ReturnField = "Code"
    SerVar = Search.Search(, , CON)
    If Len(Search.col1) <> 0 Then
        Call CancelButtonClick
        Txtc_Code = Search.col1
        Call Display_Records
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
        
        Save_Pr_LeaveEntry_Mst
        Save_Pr_LeaveEntry_Dtl
     
     CON.CommitTrans
     
     g_Sql = "HR_LEAVEENTRY_DTL_UPD "
     CON.Execute g_Sql
     
     g_Sql = "HR_LEAVE_REUPDATE_PROC " & Val(Txtc_Year)
     CON.Execute g_Sql
     
     g_SaveFlagNull = False
     Screen.MousePointer = vbDefault
    
     MsgBox "Record Saved Successfully", vbInformation, "Information"
     
     Exit Sub
     
ErrSave:
     Screen.MousePointer = vbDefault
     g_SaveFlagNull = False
     CON.RollbackTrans
     MsgBox "Error while Saving - " + Err.Description, vbCritical, "Critical"
End Sub

Private Function ChkSave() As Boolean
  Dim i As Integer
  Dim tmpQty As Double
  Dim tmpLeave As String
  
  If Trim(Txtc_Month) = "" Then
     MsgBox "Period should not be empty", vbInformation, "Information"
     Txtc_Month.SetFocus
     Exit Function
  ElseIf Trim(Txtc_Year) = "" Then
     MsgBox "Period should not be empty", vbInformation, "Information"
     Txtc_Year.SetFocus
     Exit Function
  ElseIf Not ChkPeriodOpen(vPeriod, "W") Then
     Exit Function
  End If
    
  Call Calculate_Days
  
  For i = 1 To Va_Details.DataRowCnt
      Va_Details.Row = i
      Va_Details.Col = 7
         tmpQty = Is_Null_D(Va_Details.Text, True)
         If tmpQty > 0 Then
            Va_Details.Col = 1
               If Trim(Va_Details.Text) = "" Then
                  MsgBox "Employee No. Should not be Empty", vbInformation, "Information"
                  Va_Details.SetFocus
                  Exit Function
               End If
            Va_Details.Col = 3
               If Trim(Va_Details.Text) = "" Then
                  MsgBox "Leave Should not be Empty", vbInformation, "Information"
                  Va_Details.SetFocus
                  Exit Function
               Else
                  tmpLeave = Trim(Right(Trim(Va_Details.Text), 5))
               End If
            Va_Details.Col = 4
               If Trim(Va_Details.Text) = "" Then
                  MsgBox "Type Should not be Empty", vbInformation, "Information"
                  Va_Details.SetFocus
                  Exit Function
               End If
            Va_Details.Col = 5
               If Not IsDate(Va_Details.Text) Then
                  MsgBox "From is not a valid date", vbInformation, "Information"
                  Va_Details.SetFocus
                  Exit Function
               End If
            Va_Details.Col = 6
               If Not IsDate(Va_Details.Text) Then
                  MsgBox "To is not a valid date", vbInformation, "Information"
                  Va_Details.SetFocus
                  Exit Function
               End If
               
            Va_Details.Col = 8
               If Is_Null_D(Va_Details.Text, True) < tmpQty And (tmpLeave = "CL" Or tmpLeave = "SL" Or tmpLeave = "VL") Then
                  MsgBox "Leave days should not be more than leave balance", vbInformation, "Information"
                  Va_Details.SetFocus
                  Exit Function
               End If
         End If
  Next i
  
  ChkSave = True
End Function

Private Sub Save_Pr_LeaveEntry_Mst()
    
       Set rs = Nothing
       g_Sql = "Select * from pr_leaveentry_mst where c_code = '" & Trim(Txtc_Code) & "' and c_rec_sta='A'"
       rs.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
       If rs.RecordCount = 0 Then
          rs.AddNew
          Call Start_Generate_New
          rs("c_usr_id").Value = Is_Null(g_UserName, False)
          rs("d_created").Value = GetDateTime
       Else
          rs("c_musr_id").Value = Is_Null(g_UserName, False)
          rs("d_modified").Value = GetDateTime
       End If
       
       rs("c_code").Value = Is_Null(Txtc_Code, False)
       rs("n_period").Value = Is_Null(vPeriod, True)
       rs("c_remarks").Value = Is_Null(Txtc_Remarks, False)
       rs("c_rec_sta").Value = "A"
       rs.Update
End Sub

Private Sub Save_Pr_LeaveEntry_Dtl()
Dim i As Long

      Set rs = Nothing
      g_Sql = "delete from pr_leaveentry_dtl where c_code = '" & Trim(Txtc_Code) & "'"
      CON.Execute (g_Sql)
      
      g_Sql = "Select * from pr_leaveentry_dtl where 1=2"
      rs.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
    
      For i = 1 To Va_Details.DataRowCnt
          Va_Details.Row = i
          Va_Details.Col = 7
             If Is_Null_D(Va_Details.Text, True) > 0 Then
                rs.AddNew
                rs("c_code").Value = Is_Null(Txtc_Code, False)

                Va_Details.Col = 1
                   rs("c_empno").Value = Is_Null(Va_Details.Text, False)
                Va_Details.Col = 3
                   rs("c_leave").Value = Is_Null(Right(Trim(Va_Details.Text), 7), False)
                Va_Details.Col = 4
                   rs("c_type").Value = Is_Null(Right(Trim(Va_Details.Text), 2), False)
                Va_Details.Col = 5
                   rs("d_leavefrom").Value = Is_Date(Va_Details.Text, "S")
                Va_Details.Col = 6
                   rs("d_leaveto").Value = Is_Date(Va_Details.Text, "S")
                Va_Details.Col = 7
                   rs("n_days").Value = Is_Null_D(Va_Details.Text, True)
                Va_Details.Col = 8
                   rs("n_bal").Value = Is_Null_D(Va_Details.Text, True)
                Va_Details.Col = 9
                   rs("c_remarks").Value = Is_Null(Va_Details.Text, False)
                rs.Update
             End If
      Next i
End Sub

Private Sub Start_Generate_New()
  Dim MaxNo As ADODB.Recordset
  Dim vCode As String
  
  vCode = Trim(Str(vPeriod - 200000))
  g_Sql = "Select max(substring(c_code,6,5)) from pr_leaveentry_mst where c_code like '" & vCode & "%'"
  Set MaxNo = CON.Execute(g_Sql)
  Txtc_Code = vCode + "/" + Format(Is_Null(MaxNo(0).Value, True) + 1, "00000")
End Sub

Private Sub Display_Records()
On Error GoTo Err_Display
  Dim DyDisp As New ADODB.Recordset
  Dim i, j As Long
  Dim vType As String
  
  Set DyDisp = Nothing
  g_Sql = "select * from pr_leaveentry_mst where c_code = '" & Trim(Txtc_Code) & "' and c_rec_sta='A'"
  DyDisp.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
  If DyDisp.RecordCount = 0 Then
     Exit Sub
  End If
  
  Txtc_Code = Is_Null(DyDisp("c_code").Value, False)
  vPeriod = Is_Null(DyDisp("n_period").Value, True)
  Txtc_Month = Right(Trim(Str(vPeriod)), 2)
  Txtc_Year = Left(Trim(Str(vPeriod)), 4)
  Txtc_Remarks = Is_Null(DyDisp("c_remarks").Value, False)
  Call Assign_PayPeriodDate
  
  ' // Details
   Set DyDisp = Nothing
   g_Sql = "select a.c_empno, b.c_name, b.c_othername, a.c_leave, a.c_type, a.d_leavefrom, a.d_leaveto, a.n_days, a.n_bal, a.c_remarks, c.c_leavename " & _
           "from pr_leaveentry_dtl a, pr_emp_mst b, pr_leave_mst c " & _
           "where a.c_empno = b.c_empno and a.c_leave = c.c_leave and b.c_rec_sta = 'A' and a.c_code = '" & Trim(Txtc_Code) & "' " & _
           "order by a.c_empno, c.c_leavename "
   DyDisp.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
   Va_Details.MaxRows = DyDisp.RecordCount + 25
   
   If DyDisp.RecordCount > 0 Then
      DyDisp.MoveFirst
      For i = 1 To DyDisp.RecordCount
          Va_Details.Row = i
          Va_Details.Col = 1
             Va_Details.Text = Is_Null(DyDisp("c_empno").Value, False)
          Va_Details.Col = 2
             Va_Details.Text = Proper(Is_Null(DyDisp("c_name").Value, False)) & " " & Proper(Is_Null(DyDisp("c_othername").Value, False))
          Va_Details.Col = 3
             Va_Details.Text = Is_Null(DyDisp("c_leavename").Value & Space(150) & DyDisp("c_leave").Value, False)
          Va_Details.Col = 4
              j = SelComboString(Va_Details, Is_Null(DyDisp("c_type").Value, False), i, 4, True)
              Va_Details.TypeComboBoxIndex = j
          Va_Details.Col = 5
             Va_Details.Text = Is_Date(DyDisp("d_leavefrom").Value, "D")
          Va_Details.Col = 6
             Va_Details.Text = Is_Date(DyDisp("d_leaveto").Value, "D")
          Va_Details.Col = 7
             Va_Details.Text = Is_Null(DyDisp("n_days").Value, True)
          Va_Details.Col = 8
             Va_Details.Text = Is_Null(DyDisp("n_bal").Value, True)
          Va_Details.Col = 9
             Va_Details.Text = Is_Null(DyDisp("c_remarks").Value, False)
          Va_Details.Col = 10
             Va_Details.Text = Is_Null(DyDisp("n_bal").Value, True)
          DyDisp.MoveNext
      Next i
      Call Calculate_Days
      Call Spread_Row_Height(Va_Details)
   End If
 
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

Private Sub CancelButtonClick()
    Clear_Controls Me
    Clear_Spread Va_Details
    Clear_Spread Va_Leave
End Sub

Private Function Check_Existing_Emp() As Boolean
  Dim rsChk As New ADODB.Recordset
            
     Set rsChk = Nothing
     Va_Details.Row = Va_Details.ActiveRow
     Va_Details.Col = 1
        g_Sql = "select c_empno, c_name, c_othername from pr_emp_mst " & _
                "where c_rec_sta = 'A' and (d_dol is null or d_dol >= '" & Is_Date(vPayPeriodFrom, "S") & "') and " & _
                "c_rec_sta = 'A' and c_empno = '" & Trim(Va_Details.Text) & "'"
        rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
        If rsChk.RecordCount > 0 Then
           Va_Details.Col = 1
              Va_Details.Text = Is_Null(rsChk("c_empno").Value, False)
           Va_Details.Col = 2
              Va_Details.Text = Proper(Is_Null(rsChk("c_name").Value, False)) & " " & Proper(Is_Null(rsChk("c_othername").Value, False))
           
           If Not ChkPeriodOpen(vPeriod, "W") Then
              Check_Existing_Emp = False
              Exit Function
           End If
        Else
           Va_Details.Text = ""
           MsgBox "Employee details are not found. Press <F2> to Select", vbInformation, "Information"
           Check_Existing_Emp = False
           Exit Function
        End If
     Check_Existing_Emp = True
End Function

Private Sub Spread_Lock()
  Dim i As Integer
     
    For i = 1 To Va_Details.MaxCols
        Va_Details.Row = -1
        Va_Details.Col = i
        If i = 2 Or i = 7 Or i = 8 Or i = 10 Then
           Va_Details.Lock = True
        Else
           Va_Details.Lock = False
        End If
    Next i
    
    For i = 1 To Va_Leave.MaxCols
        Va_Leave.Row = -1
        Va_Leave.Col = i
          Va_Leave.Lock = True
    Next i
    
    Va_Details.Row = -1
    Va_Details.Col = 10
       Va_Details.ColHidden = True
End Sub

Private Sub Txtc_Month_KeyPress(KeyAscii As Integer)
    Call OnlyNumeric(Txtc_Month, KeyAscii, 2)
End Sub

Private Sub Txtc_Month_Validate(Cancel As Boolean)
   Call MakeMonthTwoDigits(Me)
   If Trim(Txtc_Month) <> "" Then
      If Not Val(Txtc_Month) > 0 Or Not Val(Txtc_Month) <= 12 Then
         MsgBox "Not a valid month", vbInformation, "Information"
         Cancel = True
      End If
   End If
   If Trim(Txtc_Month) <> "" And Trim(Txtc_Year) <> "" Then
      vPeriod = Is_Null(Format(Txtc_Year, "0000") & Format(Txtc_Month, "00"), True)
      If Not ChkPeriodOpen(vPeriod, "W") Then
         Txtc_Month.SetFocus
         Cancel = True
      End If
      Call Assign_PayPeriodDate
   End If
End Sub

Private Sub txtc_year_KeyPress(KeyAscii As Integer)
    Call OnlyNumeric(Txtc_Year, KeyAscii, 4)
End Sub

Private Sub txtc_year_Validate(Cancel As Boolean)
   If Trim(Txtc_Year) <> "" Then
      If Len(Trim(Txtc_Year)) <> 4 Then
         MsgBox "Not a valid year", vbInformation, "Information"
         Txtc_Year.SetFocus
         Cancel = True
      Else
         Txtc_Year = Format(Trim(Txtc_Year), "0000")
      End If
   End If
   If Trim(Txtc_Month) <> "" And Trim(Txtc_Year) <> "" Then
      vPeriod = Is_Null(Format(Txtc_Year, "0000") & Format(Txtc_Month, "00"), True)
      If Not ChkPeriodOpen(vPeriod, "W") Then
         Txtc_Year.SetFocus
         Cancel = True
      End If
      Call Assign_PayPeriodDate
   End If
End Sub

Private Sub Txtc_No_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim Search As New Search.MyClass, SerVar
  Dim tmpFilter As String
  
  If KeyCode = vbKeyF2 Then
     tmpFilter = InputBox("Please input period to view", "Filter Option")
     If Trim(tmpFilter) <> "" Then
        tmpFilter = " and n_period like '" & Trim(tmpFilter) & "%'"
     End If
  
     Search.Query = "select n_period Period, c_remarks Remarks, c_code Code " & _
                    "from pr_leaveentry_mst  where c_rec_sta='A' " & tmpFilter
     Search.CheckFields = "Code"
     Search.ReturnField = "Code"
     SerVar = Search.Search(, , CON)
     If Len(Search.col1) <> 0 Then
        Call CancelButtonClick
        Txtc_No = Search.col1
     End If
  End If
End Sub

Private Sub Txtc_No_Validate(Cancel As Boolean)
  Dim tmpStr As String, tmpArray
  Dim tmpYear As String, tmpNo As String
    
    If Trim(Txtc_No) = "" Then
       Exit Sub
    End If
    
    tmpArray = Split(Trim(Txtc_No), "/")
    If UBound(tmpArray) = 0 Then
       tmpYear = Right(Format(Year(g_CurrentDate), "0000"), 2) & Format(Month(g_CurrentDate), "00")
       tmpNo = Format(Trim(tmpArray(0)), "00000")
    Else
       tmpYear = Trim(tmpArray(0))
       tmpNo = Format(Trim(tmpArray(1)), "00000")
    End If
    tmpStr = tmpYear & "/" & tmpNo
    
    Call CancelButtonClick
    Txtc_Code = tmpStr
    Call Display_Records
    Txtc_No = tmpStr
    
End Sub

Private Function Check_Existing_Leave() As Boolean
  Dim rsChk As New ADODB.Recordset
  Dim tmpLeave As String
     
     Va_Details.Row = Va_Details.ActiveRow
     Va_Details.Col = 3
        tmpLeave = Trim(Right(Trim(Va_Details.Text), 7))
        
     Set rsChk = Nothing
     g_Sql = "select c_leavename, c_leave from pr_leave_mst where c_rec_sta = 'A' and c_leave = '" & Trim(tmpLeave) & "'"
     rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
     If rsChk.RecordCount > 0 Then
        Va_Details.Col = 3
        Va_Details.Text = Is_Null(rsChk("c_leavename").Value, False) & Space(100) & Is_Null(rsChk("c_leave").Value, False)
     Else
        Va_Details.Col = 3
        Va_Details.Text = ""
        MsgBox "Leave not found. Please press <F2> to select leave", vbInformation, "Information"
        Check_Existing_Leave = False
        Exit Function
     End If
     Check_Existing_Leave = True
End Function

Private Sub Va_Details_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
   If BlockCol = 3 Or BlockCol = 4 Or BlockCol = 5 Or BlockCol = 6 Then
      Call SpreadBlockCopy(Va_Details, BlockCol, BlockRow, BlockCol2, BlockRow2)
      Call Calculate_Days
   End If
End Sub

Private Sub Va_Details_Click(ByVal Col As Long, ByVal Row As Long)
   If Row = 0 Then
      Call SpreadColSort(Va_Details, Col, Row)
   Else
      If Col = 1 Or Col = 2 Or Col = 3 Then
         Call GetEmpLeaveDetails(Row)
      End If
   End If
End Sub

Private Sub Va_Details_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Search As New Search.MyClass, SerVar, SerArray
Dim tmpCol3 As String
  
  If vPeriod = 0 Then
     MsgBox "Please enter the Period", vbInformation, "Information"
     Txtc_Month.SetFocus
     Exit Sub
  End If
  
  If ((Shift And 1) = 1) And KeyCode = vbKeyInsert Then
     Call SpreadInsertRow(Va_Details, Va_Details.ActiveRow)
     Call Spread_Row_Height(Va_Details)

  ElseIf ((Shift And 1) = 1) And KeyCode = vbKeyDelete Then
     Call SpreadDeleteRow(Va_Details, Va_Details.ActiveRow)
  
  ElseIf KeyCode = vbKeyDelete Then
     Call SpreadCellDataClear(Va_Details, Va_Details.ActiveRow, Va_Details.ActiveCol)
  
  ElseIf (Va_Details.ActiveCol = 1 Or Va_Details.ActiveCol = 2) And KeyCode = vbKeyF2 Then
     Search.Query = "select c_empno EmpNo, c_name Name, c_othername Othername, c_branch Branch, c_dept Dept " & _
                    "from pr_emp_mst where d_dol is null and c_rec_sta = 'A'"
     Search.CheckFields = "EmpNo, Name, OtherName"
     Search.ReturnField = "EmpNo, Name, OtherName"
     SerVar = Search.Search(, , CON)
     SerArray = Split(SerVar, "~")
     If Len(Search.col1) <> 0 Then
        Va_Details.Row = Va_Details.ActiveRow
        Va_Details.Col = 1
           Va_Details.Text = Search.col1
        Va_Details.Col = 2
           Va_Details.Text = Proper(SerArray(1)) & " " & Proper(SerArray(2))
     End If
     
  ElseIf Va_Details.ActiveCol = 3 And KeyCode = vbKeyF2 Then
     Search.Query = "select c_leave Leave, c_leavename LeaveName from pr_leave_mst where c_rec_sta = 'A' "
     Search.CheckFields = "Leave, LeaveName "
     Search.ReturnField = "Leave, LeaveName "
     SerVar = Search.Search(, , CON)
     If Len(Search.col1) <> 0 Then
        Va_Details.Row = Va_Details.ActiveRow
        Va_Details.Col = 3
           Va_Details.Text = Search.col2 & Space(150) & Search.col1
     End If
     
  ElseIf KeyCode = vbKeyF3 Then
     Va_Details.Row = Va_Details.ActiveRow
     Va_Details.Col = 3
        tmpCol3 = Trim(Va_Details.Text)
        
  ElseIf KeyCode = vbKeyF4 Then
     If MsgBox("Do you want ot Past?", vbQuestion + vbYesNo) = vbYes Then
        Va_Details.Row = Va_Details.ActiveRow
        Va_Details.Col = 3
           Va_Details.Text = Trim(tmpCol3)
     End If
  End If
End Sub

Private Sub Va_Details_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
Dim tmpQty As Double
Dim tmpLeave As String

  If Col = 1 Or Col = 2 Then
     Va_Details.Row = Row
     Va_Details.Col = 1
        If Trim(Va_Details.Text) <> "" Then
           If vPeriod > 0 Then
              If Not Check_Existing_Emp Then
                 Cancel = True
              End If
           End If
           Call GetEmpLeaveBalance(Row)
        End If
  
  ElseIf Col = 3 Then
     Va_Details.Row = Row
     Va_Details.Col = 3
        If Trim(Va_Details.Text) <> "" Then
           If Not Check_Existing_Leave Then
              Cancel = True
           End If
           Call GetEmpLeaveBalance(Row)
        End If
  
  ElseIf Col = 5 Then
     Va_Details.Row = Row
     Va_Details.Col = 5
        If IsDate(Va_Details.Text) Then
           If CDate(Va_Details.Text) < vPayPeriodFrom Then
              Va_Details.Col = 5
                 Va_Details.Text = ""
              MsgBox "Leave from date should not less than pay period start date. The pay period start date is " & Is_Date(vPayPeriodFrom, "D"), vbInformation, "Information"
              Va_Details.SetFocus
              Cancel = True
           End If
        End If
     Call Calculate_Days(Row)
  
  ElseIf Col = 4 Or Col = 5 Or Col = 6 Then
     Call Calculate_Days(Row)
  End If
  
  If Col = 6 Then
     Va_Details.Row = Row
     Va_Details.Col = 3
        tmpLeave = Trim(Right(Trim(Va_Details.Text), 5))
     Va_Details.Col = 7
        tmpQty = Is_Null_D(Va_Details.Text, True)
     Va_Details.Col = 6
        If IsDate(Va_Details.Text) And tmpQty > 0 Then
           Va_Details.Col = 8
              If Is_Null_D(Va_Details.Text, True) < tmpQty And (tmpLeave = "CL" Or tmpLeave = "SL" Or tmpLeave = "VL") Then
                 Va_Details.Col = 6
                    Va_Details.Text = ""
                 MsgBox "Leave days should not be more than leave balance", vbInformation, "Information"
                 Va_Details.SetFocus
                 Cancel = True
              End If
        End If
  End If

  If (Col = 1 Or Col = 2 Or Col = 3) And Row = NewRow Then
     Call GetEmpLeaveDetails(Row)
  End If
End Sub

Private Sub GetEmpLeaveBalance(ByVal vRow As Long)
Dim rsChk As New ADODB.Recordset
Dim vEmpNo As String, vLeave As String
Dim tmpQty As Double

    Va_Details.Row = vRow
    Va_Details.Col = 1
       vEmpNo = Trim(Va_Details.Text)
    Va_Details.Col = 3
       vLeave = Trim(Right(Trim(Va_Details.Text), 7))
    Va_Details.Col = 10
       tmpQty = Is_Null_D(Va_Details.Text, True)
    
    If vEmpNo <> "" And vLeave <> "" And tmpQty = 0 Then
       Set rsChk = Nothing
       g_Sql = "select a.c_empno, a.c_name, a.c_othername, b.n_opbal, b.n_entitle, b.n_alloted, b.n_utilised, b.n_adjusted, b.n_clbal " & _
               "from pr_emp_mst a, pr_emp_leave_dtl b " & _
               "where a.c_empno = b.c_empno and a.c_rec_sta = 'A' and a.c_empno = '" & vEmpNo & "' and " & _
               "b.c_leave = '" & vLeave & "' and year(b.d_prfrom) = " & Val(Txtc_Year)
       rsChk.Open g_Sql, CON, adLockOptimistic, adLockReadOnly
       If rsChk.RecordCount > 0 Then
          Va_Details.Row = vRow
          Va_Details.Col = 8
             Va_Details.Text = Spread_NumFormat(Is_Null(rsChk("n_clbal").Value, True), True, 2)
             
          Lbl_EmpDtl = Is_Null(rsChk("c_empno").Value, False) & "   " & Proper(Is_Null(rsChk("c_name").Value, False)) & " " & Proper(Is_Null(rsChk("c_othername").Value, False))
          Txtn_Open = Format_Num(Is_Null(rsChk("n_opbal").Value, True))
          Txtn_Entitle = Format_Num(Is_Null(rsChk("n_entitle").Value, True))
          Txtn_Alloted = Format_Num(Is_Null(rsChk("n_alloted").Value, True))
          Txtn_Used = Format_Num(Is_Null(rsChk("n_utilised").Value, True))
          Txtn_Adj = Format_Num(Is_Null(rsChk("n_adjusted").Value, True))
          Txtn_Close = Format_Num(Is_Null(rsChk("n_clbal").Value, True))
       Else
           Va_Details.Row = vRow
           Va_Details.Col = 8
             Va_Details.Text = 0
       End If
    End If
End Sub

Private Sub GetEmpLeaveDetails(ByVal vRow As Long)
  Dim rsChk As New ADODB.Recordset
  Dim i As Integer
  Dim vEmpNo As String, vLeave As String, vName As String
  Dim vLeaveTot As Double
  
    Va_Details.Row = vRow
    Va_Details.Col = 1
       vEmpNo = Trim(Va_Details.Text)
    Va_Details.Col = 2
       vName = Trim(Va_Details.Text)
    Va_Details.Col = 3
       vLeave = Trim(Right(Trim(Va_Details.Text), 7))
       
    If vEmpNo <> "" Then
       Lbl_EmpDtl = vEmpNo & "   " & vName
       Txtn_Open = ""
       Txtn_Entitle = ""
       Txtn_Alloted = ""
       Txtn_Used = ""
       Txtn_Adj = ""
       Txtn_Close = ""
    End If
       
    If vEmpNo <> "" And vLeave <> "" Then
       Set rsChk = Nothing
       g_Sql = "select a.c_empno, a.c_name, a.c_othername, b.n_opbal, b.n_entitle, b.n_alloted, b.n_utilised, b.n_adjusted, b.n_clbal " & _
               "from pr_emp_mst a left outer join pr_emp_leave_dtl b  " & _
               "  on a.c_empno = b.c_empno and b.c_leave = '" & vLeave & "' and year(b.d_prfrom) = " & Val(Txtc_Year) & " " & _
               "where a.c_rec_sta = 'A' and a.c_empno = '" & vEmpNo & "' "
       rsChk.Open g_Sql, CON, adLockOptimistic, adLockReadOnly
       If rsChk.RecordCount > 0 Then
          Lbl_EmpDtl = Is_Null(rsChk("c_empno").Value, False) & "   " & Proper(Is_Null(rsChk("c_name").Value, False)) & " " & Proper(Is_Null(rsChk("c_othername").Value, False))
          Txtn_Open = Format_Num(Is_Null(rsChk("n_opbal").Value, True))
          Txtn_Entitle = Format_Num(Is_Null(rsChk("n_entitle").Value, True))
          Txtn_Alloted = Format_Num(Is_Null(rsChk("n_alloted").Value, True))
          Txtn_Used = Format_Num(Is_Null(rsChk("n_utilised").Value, True))
          Txtn_Adj = Format_Num(Is_Null(rsChk("n_adjusted").Value, True))
          Txtn_Close = Format_Num(Is_Null(rsChk("n_clbal").Value, True))
       End If
       
       
       Clear_Spread Va_Leave
       Set rsChk = Nothing
       
       g_Sql = "select a.c_empno, c.c_leavename,  b.d_date, b.n_present " & _
               "from pr_emp_mst a, pr_clock_emp b, pr_leave_mst c " & _
               "where a.c_empno = b.c_empno and b.c_presabs = c.c_leave and year(b.d_date) = " & Val(Txtc_Year) & " and " & _
               "a.c_rec_sta = 'A' and a.c_empno = '" & vEmpNo & "' and b.c_presabs = '" & vLeave & "' " & _
               "order by b.d_date "
       rsChk.Open g_Sql, CON, adLockOptimistic, adLockReadOnly
      
       Va_Leave.MaxRows = rsChk.RecordCount + 50
       vLeaveTot = 0
       For i = 1 To rsChk.RecordCount
          Va_Leave.Row = i
          Va_Leave.Col = 1
             Va_Leave.Text = Is_Date(rsChk("d_date").Value, "D")
          Va_Leave.Col = 2
             Va_Leave.Text = Is_Null(rsChk("c_leavename").Value, False)
          Va_Leave.Col = 3
             Va_Leave.Text = Spread_NumFormat(Is_Null(rsChk("n_present").Value, True), True)
             vLeaveTot = vLeaveTot + Is_Null(rsChk("n_present").Value, True)
          rsChk.MoveNext
       Next i
       
       If vLeaveTot > 0 Then
          Va_Leave.Row = i + 1
          Va_Leave.Col = 2
             Va_Leave.Text = "Total"
          Va_Leave.Col = 3
             Va_Leave.Text = Spread_NumFormat(vLeaveTot, True)
       End If
    End If
    Call Spread_Row_Height(Va_Leave)
End Sub

Private Sub Calculate_Days(Optional ByVal vRow As Long)
  Dim i As Long, vStartRow As Long, vEndRow As Long
  Dim vFrom As String, vTo As String, vType As String
  Dim vDays As Double
      
      If vRow = 0 Then
         vStartRow = 1:   vEndRow = Va_Details.DataRowCnt
      Else
         vStartRow = vRow: vEndRow = vRow
      End If

      For i = vStartRow To vEndRow
          vDays = 0: vFrom = "": vTo = ""
          Va_Details.Row = i
          Va_Details.Col = 4
             vType = Trim(Right(Trim(Va_Details.Text), 2))
          Va_Details.Col = 5
             If IsDate(Va_Details.Text) Then
                vFrom = Is_Date(Va_Details.Text, "S")
             End If
          Va_Details.Col = 6
             If IsDate(Va_Details.Text) Then
                vTo = Is_Date(Va_Details.Text, "S")
             End If
          Va_Details.Col = 7
             If Trim(vFrom) <> "" And Trim(vTo) <> "" Then
                vDays = DateDiff("d", CDate(vFrom), CDate(vTo)) + 1
             End If
             If vDays > 0 Then
                If vType = "AM" Or vType = "PM" Then
                   vDays = vDays - 0.5
                End If
                Va_Details.Text = Is_Null(vDays, True)
             Else
                Va_Details.Text = ""
             End If
      Next i

End Sub

Private Sub Disable_Controls()
    Txtc_No.Enabled = False
End Sub

Private Sub Assign_PayPeriodDate()
  Dim rsChk As New ADODB.Recordset

    Set rsChk = Nothing
    g_Sql = "select d_fromdate, d_todate from pr_payperiod_dtl where n_period = " & vPeriod & " and c_type = 'W' "
    rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If rsChk.RecordCount > 0 Then
       vPayPeriodFrom = rsChk("d_fromdate").Value
       vPayPeriodTo = rsChk("d_todate").Value
    End If
End Sub

Private Sub Load_Combo()
  Dim tmpStr As String
  
    tmpStr = "Full Day" & Space(50) & "FD" & Chr$(9) & _
             "First Half" & Space(50) & "AM" & Chr$(9) & _
             "Second Half" & Space(50) & "PM"
    
    If Len(tmpStr) > 0 Then
        Va_Details.Row = -1
       Va_Details.Col = 4
          Va_Details.TypeComboBoxList = tmpStr
    End If
End Sub

Private Sub Btn_GetEmp_Click()

    If vPeriod = 0 Then
       MsgBox "Period should not be empty", vbInformation, "Information"
       Txtc_Month.SetFocus
       Exit Sub
    End If

    Call LoadComboCompany(Me)
    Call LoadComboBranch(Me)
    Call LoadComboDept(Me)
    Call LoadComboDept(Me)
    Call LoadComboEmpType(Me)
    Call LoadComboShift(Me)
    Call LoadComboLeave(Me)
    
    Dtp_FromDate.Text = Is_Date(Null, "D")
    Dtp_ToDate.Text = Is_Date(Null, "D")
    
    Cmb_Type.Clear
    Cmb_Type.AddItem "Full Day" & Space(50) & "FD"
    Cmb_Type.AddItem "First Half" & Space(50) & "AM"
    Cmb_Type.AddItem "Second Half" & Space(50) & "PM"
    Cmb_Type.ListIndex = 0

    Frm_Filter.Left = 5300
    Frm_Filter.Top = 1460
    Frm_Filter.Visible = True
    Frm_Filter.ZOrder 0
End Sub


Private Sub Btn_Cancel_Click()
    Frm_Filter.Visible = False
End Sub

Private Sub Btn_Ok_Click()
  Dim rsChk As New ADODB.Recordset
  Dim i As Integer, j As Integer, vRow As Integer
  Dim vLeave As String
  
    If Trim(Cmb_Leave) = "" Then
       MsgBox "Leave should not be empty", vbInformation, "Information"
       Cmb_Leave.SetFocus
       Exit Sub
    ElseIf Trim(Cmb_Type) = "" Then
       MsgBox "Type should not be empty", vbInformation, "Information"
       Cmb_Type.SetFocus
       Exit Sub
    ElseIf Not IsDate(Dtp_FromDate.Text) Then
       MsgBox "From date should not be empty", vbInformation, "Information"
       Dtp_FromDate.SetFocus
       Exit Sub
    ElseIf Not IsDate(Dtp_ToDate.Text) Then
       MsgBox "To date should not be empty", vbInformation, "Information"
       Dtp_ToDate.SetFocus
       Exit Sub
    End If
  
    vRow = Va_Details.DataRowCnt
    vLeave = Trim(Right(Trim(Cmb_Leave), 7))
    
    Set rsChk = Nothing
    If vLeave = "CL" Or vLeave = "SL" Or vLeave = "VL" Then
       g_Sql = "select a.c_empno, a.c_name, a.c_othername, b.n_clbal " & _
               "from pr_emp_mst a left outer join pr_emp_leave_dtl b " & _
               "         on a.c_empno = b.c_empno and b.c_leave = '" & vLeave & "' and year(b.d_prfrom) = " & Val(Txtc_Year) & " " & _
               "where a.c_rec_sta = 'A' and a.d_dol is null "
    Else
       g_Sql = "select a.c_empno, a.c_name, a.c_othername, 0 n_clbal from pr_emp_mst a where a.c_rec_sta = 'A' and a.d_dol is null "
    End If
    
    If Trim(Cmb_Company) <> "" Then
       g_Sql = g_Sql & " and a.c_company = '" & Trim(Right(Trim(Cmb_Company), 7)) & "'"
    End If
    
    If Trim(Cmb_Branch) <> "" Then
       g_Sql = g_Sql & " and a.c_branch = '" & Trim(Cmb_Branch) & "'"
    End If
            
    If Trim(Cmb_Dept) <> "" Then
       g_Sql = g_Sql & " and a.c_dept = '" & Trim(Cmb_Dept) & "'"
    End If
    
    If Trim(Cmb_Desig) <> "" Then
       g_Sql = g_Sql & " and a.c_desig = '" & Trim(Cmb_Desig) & "'"
    End If
    
    If Trim(Cmb_EmpType) <> "" Then
       g_Sql = g_Sql & " and a.c_emptype = '" & Trim(Cmb_EmpType) & "'"
    End If
    
    If Trim(Cmb_Shift) <> "" Then
       g_Sql = g_Sql & " and a.c_shiftcode = '" & Trim(Right(Trim(Cmb_Shift), 7)) & "'"
    End If
    
    g_Sql = g_Sql & " order by a.c_empno "
    
    
    rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    Va_Details.MaxRows = rsChk.RecordCount + vRow + 25
    
    If rsChk.RecordCount > 0 Then
       rsChk.MoveFirst
       For i = 1 To rsChk.RecordCount
           Va_Details.Row = i + vRow
           Va_Details.Col = 1
              Va_Details.Text = Is_Null(rsChk("c_empno").Value, False)
           Va_Details.Col = 2
              Va_Details.Text = Proper(Is_Null(rsChk("c_name").Value, False)) & " " & Proper(Is_Null(rsChk("c_othername").Value, False))
           Va_Details.Col = 3
              Va_Details.Text = Trim(Cmb_Leave.Text)
           Va_Details.Col = 4
              j = SelComboString(Va_Details, Trim(Right(Trim(Cmb_Type), 2)), i + vRow, 4, True)
              Va_Details.TypeComboBoxIndex = j
           Va_Details.Col = 5
              Va_Details.Text = Is_DateSpread(Dtp_FromDate.Text)
           Va_Details.Col = 6
              Va_Details.Text = Is_DateSpread(Dtp_FromDate.Text)
           Va_Details.Col = 8
              Va_Details.Text = Spread_NumFormat(Is_Null(rsChk("n_clbal").Value, True), True, 2)
           rsChk.MoveNext
       Next i
        
       Call Calculate_Days
       Call Spread_Row_Height(Va_Details)
    End If
    
    Frm_Filter.Visible = False
End Sub
