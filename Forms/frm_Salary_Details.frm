VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Begin VB.Form frm_Salary_Details 
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3840
   ScaleWidth      =   6840
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Height          =   900
      Left            =   120
      TabIndex        =   42
      Top             =   -45
      Width           =   19155
      Begin VB.CheckBox Chk_Hide 
         Caption         =   "Hide EDF Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   210
         Index           =   1
         Left            =   10575
         TabIndex        =   8
         Top             =   375
         Width           =   2325
      End
      Begin VB.CheckBox Chk_Hide 
         Caption         =   "Hide Employee Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   210
         Index           =   0
         Left            =   10575
         TabIndex        =   7
         Top             =   135
         Width           =   2385
      End
      Begin VB.CheckBox Chk_Hide 
         Caption         =   "Hide Bank Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   210
         Index           =   2
         Left            =   10575
         TabIndex        =   9
         Top             =   600
         Width           =   2325
      End
      Begin VB.CheckBox Chk_Hide 
         Caption         =   "Hide Hours Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   210
         Index           =   3
         Left            =   13395
         TabIndex        =   10
         Top             =   120
         Width           =   2325
      End
      Begin VB.CheckBox Chk_Hide 
         Caption         =   "Hide Company Contributions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   210
         Index           =   5
         Left            =   13395
         TabIndex        =   12
         Top             =   600
         Width           =   2865
      End
      Begin VB.CheckBox Chk_Hide 
         Caption         =   "Hide Notes && Coins "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   210
         Index           =   6
         Left            =   16545
         TabIndex        =   13
         Top             =   135
         Width           =   2175
      End
      Begin VB.CheckBox Chk_Hide 
         Caption         =   "Hide Leave Pay Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   210
         Index           =   4
         Left            =   13395
         TabIndex        =   11
         Top             =   360
         Width           =   2325
      End
      Begin VB.CommandButton Btn_GridDefault 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Set as Your Default Grid"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   16530
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   420
         Width           =   2295
      End
      Begin VB.CommandButton Btn_Export 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   4425
         Picture         =   "frm_Salary_Details.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Hide 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   2685
         Picture         =   "frm_Salary_Details.frx":3647
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Restore 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   3420
         Picture         =   "frm_Salary_Details.frx":6D4C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Save 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   90
         Picture         =   "frm_Salary_Details.frx":A412
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Exit 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   5160
         Picture         =   "frm_Salary_Details.frx":DA9B
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Clear 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   945
         Picture         =   "frm_Salary_Details.frx":110FB
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Print 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   1680
         Picture         =   "frm_Salary_Details.frx":1476B
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   700
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   975
      Left            =   120
      TabIndex        =   32
      Top             =   1245
      Width           =   19155
      Begin VB.CommandButton Btn_Display 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   17790
         Picture         =   "frm_Salary_Details.frx":17D59
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   150
         Width           =   700
      End
      Begin VB.ComboBox Cmb_EmpType 
         Height          =   315
         Left            =   12900
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   180
         Width           =   1200
      End
      Begin VB.ComboBox Cmb_Nationality 
         Height          =   315
         Left            =   15360
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   180
         Width           =   1500
      End
      Begin VB.TextBox Txtc_YearFrom 
         Height          =   300
         Left            =   2310
         TabIndex        =   16
         Top             =   210
         Width           =   705
      End
      Begin VB.TextBox Txtc_MonthFrom 
         Height          =   300
         Left            =   1770
         TabIndex        =   15
         Top             =   210
         Width           =   525
      End
      Begin VB.ComboBox Cmb_Branch 
         Height          =   315
         Left            =   4710
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   525
         Width           =   2370
      End
      Begin VB.ComboBox Cmb_Dept 
         Height          =   315
         Left            =   8760
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   210
         Width           =   2370
      End
      Begin VB.ComboBox Cmb_Company 
         Height          =   315
         Left            =   4710
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   210
         Width           =   2370
      End
      Begin VB.TextBox Txtc_YearTo 
         Height          =   300
         Left            =   2310
         TabIndex        =   18
         Top             =   510
         Width           =   705
      End
      Begin VB.TextBox Txtc_MonthTo 
         Height          =   300
         Left            =   1770
         TabIndex        =   17
         Top             =   525
         Width           =   525
      End
      Begin VB.ComboBox Cmb_Desig 
         Height          =   315
         Left            =   8760
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   525
         Width           =   2370
      End
      Begin VB.TextBox Txtc_EmployeeName 
         Height          =   300
         Left            =   12900
         TabIndex        =   25
         Top             =   510
         Width           =   3960
      End
      Begin VB.Label Label3 
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
         Left            =   12405
         TabIndex        =   41
         Top             =   225
         Width           =   405
      End
      Begin VB.Label Label1 
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
         Left            =   14460
         TabIndex        =   40
         Top             =   225
         Width           =   840
      End
      Begin VB.Label Lbl_Period 
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
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   645
         TabIndex        =   39
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label Label9 
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
         Left            =   4050
         TabIndex        =   38
         Top             =   570
         Width           =   570
      End
      Begin VB.Label Label4 
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
         Left            =   8295
         TabIndex        =   37
         Top             =   255
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
         Left            =   3840
         TabIndex        =   36
         Top             =   255
         Width           =   780
      End
      Begin VB.Label Label5 
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
         Left            =   8220
         TabIndex        =   35
         Top             =   570
         Width           =   465
      End
      Begin VB.Label Label8 
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
         Left            =   870
         TabIndex        =   34
         Top             =   570
         Width           =   795
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   12375
         TabIndex        =   33
         Top             =   540
         Width           =   465
      End
   End
   Begin VB.TextBox Txtn_NoRows 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   1365
      TabIndex        =   28
      Top             =   9570
      Width           =   975
   End
   Begin FPUSpreadADO.fpSpread Va_Details 
      Height          =   7290
      Left            =   120
      TabIndex        =   27
      Top             =   2235
      Width           =   19125
      _Version        =   458752
      _ExtentX        =   33734
      _ExtentY        =   12859
      _StockProps     =   64
      AutoClipboard   =   0   'False
      ColsFrozen      =   2
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
      MaxCols         =   123
      MaxRows         =   50
      ProcessTab      =   -1  'True
      SpreadDesigner  =   "frm_Salary_Details.frx":1B497
   End
   Begin MSComDlg.CommonDialog ComDialog 
      Left            =   19545
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "No. of Rows"
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
      Left            =   240
      TabIndex        =   31
      Top             =   9600
      Width           =   975
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
      Left            =   17640
      TabIndex        =   30
      Top             =   975
      Width           =   360
   End
   Begin VB.Label lbl_scr_name 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Salary Details"
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
      Left            =   240
      TabIndex        =   29
      Top             =   975
      Width           =   1095
   End
   Begin VB.Shape shp_scr_name 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   300
      Left            =   120
      Top             =   945
      Width           =   19140
   End
End
Attribute VB_Name = "frm_Salary_Details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As New ADODB.Recordset
Private tmpFrmActivateCtr As Integer
Private vPayPeriodFrom As Long, vPayPeriodTo As Long
Private strFilter As String
Private SelFor As String, SelFilFor As String, RepHead As String, RepHeadDate As String
Private vF1, vF2, vF3, vF4, vF5 As String
Private vFromMonthName As String, vToMonthName As String

Private Sub Form_Activate()
    Txtc_MonthFrom.SetFocus
End Sub

Private Sub Form_Load()
    lbl_scr_name = "Salary Details"
    lbl_date.Caption = Format(Date, "dd-mmm-yyyy")
    Enable_Controls Me, True

    If Trim(Txtc_YearFrom) = "" Then
       Txtc_YearFrom = Year(g_CurrentDate)
    End If
    Chk_Hide(0).Value = 1
    Chk_Hide(1).Value = 1
    Chk_Hide(2).Value = 1
    Chk_Hide(3).Value = 1
    Chk_Hide(4).Value = 1
    Chk_Hide(5).Value = 1
    Chk_Hide(6).Value = 1
    Call Spread_Hide_Check
    tmpFrmActivateCtr = tmpFrmActivateCtr + 1

    Clear_Spread Va_Details

    Call Combo_Load
    Call Spread_Lock
    Call TGControlProperty(Me)
    Call SpreadHeaderFont(Va_Details, "Arial", 7, False)
    Call Spread_Header_Name
    
    Btn_Save.Visible = False
'    Btn_View.Visible = False
'    Btn_Delete.Visible = False
    Cmb_Company.ListIndex = 1
End Sub

Private Sub Btn_Exit_Click()
    Unload Me
End Sub

Private Sub Btn_Clear_Click()
  Dim i As Integer
  
    Call Clear_Spread(Va_Details)
    Va_Details.MaxRows = 0
    Va_Details.MaxRows = 50
    
    Cmb_Branch.ListIndex = 0
    Cmb_Dept.ListIndex = 0
    Cmb_Desig.ListIndex = 0
    Cmb_EmpType.ListIndex = 0
    Txtc_EmployeeName = ""

    For i = 1 To Va_Details.MaxRows
        Va_Details.Row = i
        Va_Details.Col = 1
           Va_Details.Value = False
    Next i
    
End Sub

Private Sub Btn_Print_Click()
On Error GoTo Err_Print
  Dim RsChk As New ADODB.Recordset
  Dim SelFor As String
  Dim tmpStr As String, RepOpt As String, RepOptSub As String
  Dim RepTitle As String, RepDate As String, tmpCompName As String
  Dim vF1, vF2, vF3, vF4, vF5 As String
  
    SelFor = "": RepTitle = "": RepDate = "": tmpCompName = ""
    vF1 = "": vF2 = "": vF3 = "": vF4 = "": vF5 = ""
    RepOpt = "": RepOptSub = ""
  
    Call Assign_PayPeriod
    
    RepDate = MakeReportSubHead(Me)
 
    Set RsChk = Nothing
    g_Sql = "select min(d_fromdate) d_fromdate, max(d_todate) d_todate, sum(n_days) n_days from pr_payperiod_dtl " & _
            "where n_period >= " & vPayPeriodFrom & " and n_period <= " & vPayPeriodTo & " and c_type = 'W'"
    RsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If RsChk.RecordCount > 0 Then
       RepDate = RepDate & Space(15) & "Pay period form " & Format(RsChk("d_fromdate").Value, "dd/mm/yyyy") & _
                 " to " & Format(RsChk("d_todate").Value, "dd/mm/yyyy")
    End If
            
    SelFor = SalaryMaster_RepFilter
    
    tmpStr = "1. Paysheet Report " & vbCrLf & _
             "2. Paysheet Summary " & vbCrLf & _
             "3. Payment Details (Check List) " & vbCrLf & vbCrLf & _
             "4. Pay Slip " & vbCrLf & vbCrLf & _
             "5. Travel Allowance Details "

    RepOpt = InputBox(tmpStr, "Choose your option", "1")
    
    If Val(RepOpt) = 0 Then
       Exit Sub
    End If
    
    If Val(RepOpt) = 1 Then
       RepTitle = "Pay Sheet Details - " & getPeriodTitle
       RepTitle = MakeReportHeadShort(Me, RepTitle, False)
       Call Print_Rpt(SelFor, "Pr_Salary_Dtl.rpt")
    
    ElseIf Val(RepOpt) = 2 Then
       Call Print_Rpt(SelFor, "Pr_PaySheet_Summary.rpt")
    
    ElseIf Val(RepOpt) = 3 Then
       tmpStr = "1. Payment Details " & vbCrLf & _
                "2. Payment Details - All type "
       RepOpt = InputBox(tmpStr, "Choose your option", "1")
                       
       If Val(RepOpt) = 1 Then
          SelFor = SelFor & " AND {PR_SALARY_DTL.C_SALARY} Like ['SAL*'] "
          SelFor = SelFor & " AND {PR_SALARY_DTL.C_SALTYPE} <> 'O' "
          Call Print_Rpt(SelFor, "Pr_Sal_Emp_Details.rpt")
       ElseIf Val(RepOpt) = 2 Then
          Call Print_Rpt(SelFor, "Pr_Sal_Emp_ChkLst.rpt")
       Else
          Exit Sub
       End If
    
    ElseIf Val(RepOpt) = 4 Then
       Call Pr_Payslip_Emp_Gen
       SelFor = ""
       Call Print_Rpt(SelFor, "Pr_Salary_PaySlip.rpt")
    
    ElseIf Val(RepOpt) = 5 Then
       SelFor = SelFor & " AND {PR_SALARY_DTL.C_SALARY}='SAL0010' "
       Call Print_Rpt(SelFor, "Pr_Sal_TA_Details.rpt")
    
    Else
       Exit Sub
    End If
    
   If Trim(RepTitle) <> "" Then
      Mdi_Ta_HrPay.CRY1.Formulas(1) = "ReportHead='" & Trim(RepTitle) & "'"
   End If

   If Trim(RepDate) <> "" Then
      Mdi_Ta_HrPay.CRY1.Formulas(2) = "RepHeadDate='" & Trim(RepDate) & "'"
   End If
   Mdi_Ta_HrPay.CRY1.Action = 1
  
  Exit Sub

Err_Print:
    MsgBox "Error while Generating - " + Err.Description, vbInformation, "Information"
End Sub

Private Sub Chk_Hide_Click(Index As Integer)
    Call Spread_ColHide
End Sub

Private Sub Btn_Display_Click()
    
  If Val(Txtc_MonthFrom) = 0 Then
     MsgBox "Please enter period", vbInformation, "Information"
     Txtc_MonthFrom.SetFocus
     Exit Sub
  ElseIf Val(Txtc_YearFrom) = 0 Then
     MsgBox "Please enter period", vbInformation, "Information"
     Txtc_YearFrom.SetFocus
     Exit Sub
  End If
     
  If Va_Details.DataRowCnt > 0 Then
     If MsgBox("The Data will be refesh?. Do you want to display?", vbYesNo + vbDefaultButton2, "Confirmation") = vbNo Then
        Exit Sub
     Else
        Va_Details.MaxRows = 0
        Va_Details.MaxRows = 50
     End If
  End If
  
  Call Display_Records

End Sub

Private Sub Display_Records()
On Error GoTo Err_Display
  Dim DyDisp As New ADODB.Recordset
  Dim i, j, tmpTotOt  As Double
  Dim tmpEmpNo As String, tmpPeriod As Long
  
  Dim vTotBasicWrkHrs, vOtHrs15, vOtHrs2, vOtHrs3, vSunPhHrs2, vSunPhHrs3, vOtHrs20, vOtHrs30, vTotOTHrs
  Dim vBasic, vEarnFR2, vEarnFR3, vEarnFR4, vEarnFR5, vOt15, vOt2, vOt3, vSunPh2, vSunPh3, vOt20, vOt30, vTotOT
  Dim vBonus, vPresBonus, vMealAllow, vNightAllow, vTravel, vTravelTax, vTotLeave, vCashAdd, vEOYBonus, vOthIncome, vTotIncome
  Dim vLop, vLate, vEmpNps, vEmpEwf, vEmpEpz, vEmpLevy, vLoan, vPayee, vCashDeduct, vOthDeduction, vTotDeduction, vNet
  Dim vComNps, vComEwf, vComEpz, vComLevy
  Dim v1000, v500, v200, v100, v50, v10
  
  Call Assign_PayPeriod
  
  g_Sql = "Select a.c_empno, a.n_period, b.c_name, b.c_othername, b.c_itno, a.c_company, e.c_companyname, a.c_branch, a.c_dept, a.c_desig, " & _
          "a.n_basic, a.n_fixedrate2, a.n_fixedrate3, a.n_fixedrate4, a.n_fixedrate5, " & _
          "a.c_emptype, a.c_stafftype, a.c_salarytype, a.c_paytype, a.c_bank, d.c_shortname c_bankshortname, a.c_bankcode, a.c_acctno, " & _
          "a.c_edfcat, a.n_edfamount, a.n_eduamount, a.n_intamount, a.n_preamount, a.n_othamount, a.n_edftotal, " & _
          "a.n_wrkhrs, a.n_earnbasic, a.n_earnfr2, a.n_earnfr3, a.n_earnfr4, a.n_earnfr5, a.n_ot15, a.n_ot20, a.n_ot30, a.n_sunph20, a.n_sunph30, a.n_ot, " & _
          "a.n_bonus, a.n_presbonus, a.n_mealallow, a.n_nightallow, a.n_travel, a.n_travel_tax, a.n_ph, a.n_local, a.n_sick, a.n_injury, a.n_prolong, " & _
          "a.n_wedding, a.n_maternity, a.n_paternity, a.n_compoff, a.n_others, a.n_totleave, a.n_cashadd, a.n_othincome, a.n_lop, a.n_late, " & _
          "a.n_empnps, a.n_empepz, a.n_empewf, a.n_emplevy, a.n_comnps, a.n_comepz, a.n_comewf, a.n_comlevy, " & _
          "a.n_loan, a.n_paye, a.n_cashdeduct, a.n_othdeduction, a.n_totincome, a.n_totdeduction, a.n_net, " & _
          "c.n_lophrs, c.n_latehrs, c.n_othrs15, c.n_othrs20, c.n_othrs30, c.n_sunphhrs20, c.n_sunphhrs30, c.n_totothrs, c.n_publicholiday, c.n_localleave, c.n_sickleave, " & _
          "c.n_injuryleave, c.n_prolongleave, c.n_weddingleave, c.n_matleave, c.n_patleave, c.n_compleave, c.n_othleave, c.n_lopdays, c.n_sl_fullday, " & _
          "c.n_no_travelallow , c.n_no_mealallow, c.n_no_nightallow, c.n_noweek, " & _
          "a.n_bonusincome, a.n_bonusdeduction, a.n_bonuspaye, a.n_eoybonus, a.n_carbenefit, a.n_1000, a.n_500, a.n_200, a.n_100, a.n_50, a.n_10 " & _
          "from pr_salary_mst a left outer join pr_workhrs_dtl c on a.c_empno = c.c_empno and a.n_period = c.n_period " & _
          "                     left outer join pr_bankmast d on a.c_bank = d.c_code, " & _
          "     pr_emp_mst b, pr_company_mst e " & _
          "where a.c_empno = b.c_empno and a.c_company = e.c_company and a.c_rec_sta = 'A' "
          
  g_Sql = g_Sql & " and a.n_period >= " & vPayPeriodFrom & " and a.n_period <= " & vPayPeriodTo
  
  If Trim(Cmb_Company) <> "" Then
     g_Sql = g_Sql & " and b.c_company = '" & Right(Trim(Cmb_Company), 7) & "'"
  End If
  If Trim(Cmb_Branch) <> "" Then
     g_Sql = g_Sql & " and b.c_branch = '" & Trim(Cmb_Branch) & "'"
  End If
  If Trim(Cmb_Dept) <> "" Then
     g_Sql = g_Sql & " and b.c_dept = '" & Trim(Cmb_Dept) & "'"
  End If
  If Trim(Cmb_Desig) <> "" Then
     g_Sql = g_Sql & " and b.c_desig = '" & Trim(Cmb_Desig) & "'"
  End If
  If Trim(Cmb_EmpType) <> "" Then
     g_Sql = g_Sql & " and b.c_emptype = '" & Right(Trim(Cmb_EmpType), 1) & "'"
  End If
  If Trim(Cmb_Nationality) <> "" Then
     If Trim(Cmb_Nationality) = "All Expatriate" Then
        g_Sql = g_Sql & " and b.c_expatriate = 'Y' "
     Else
        g_Sql = g_Sql & " and b.c_nationality = '" & Trim(Cmb_Nationality) & "'"
     End If
  End If
  If Trim(Txtc_EmployeeName) <> "" Then
     g_Sql = g_Sql & " and a.c_empno = '" & Trim(Right(Trim(Txtc_EmployeeName), 7)) & "'"
  End If
  
  g_Sql = g_Sql & " order by a.c_empno, a.n_period"
  
  Call Clear_Spread(Va_Details)
  
  vTotBasicWrkHrs = 0:  vOtHrs15 = 0: vOtHrs2 = 0: vOtHrs3 = 0:  vSunPhHrs2 = 0: vSunPhHrs3 = 0: vOtHrs20 = 0: vOtHrs30 = 0: vTotOTHrs = 0:
  vBasic = 0: vEarnFR2 = 0: vEarnFR3 = 0: vEarnFR4 = 0: vEarnFR5 = 0:  vOt15 = 0: vOt2 = 0: vOt3 = 0:  vSunPh2 = 0: vSunPh3 = 0: vOt20 = 0: vOt30 = 0: vTotOT = 0
  vBonus = 0: vPresBonus = 0: vMealAllow = 0: vNightAllow = 0: vTravel = 0: vTravelTax = 0: vTotLeave = 0: vCashAdd = 0: vEOYBonus = 0: vOthIncome = 0: vTotIncome = 0
  vLop = 0: vLate = 0: vEmpNps = 0: vEmpEwf = 0: vEmpEpz = 0: vEmpLevy = 0: vLoan = 0: vPayee = 0: vCashDeduct = 0: vOthDeduction = 0: vTotDeduction = 0: vNet = 0
  vComNps = 0: vComEwf = 0: vComEpz = 0: vComLevy = 0
  v1000 = 0: v500 = 0: v200 = 0: v100 = 0: v50 = 0: v10 = 0
  
  Set DyDisp = Nothing
  DyDisp.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
  Va_Details.MaxRows = DyDisp.RecordCount * 2
  
  j = 0
  If DyDisp.RecordCount > 0 Then
     DyDisp.MoveFirst
     For i = 1 To DyDisp.RecordCount
         j = j + 1
         If tmpEmpNo <> Is_Null(DyDisp("c_empno").Value, False) And tmpPeriod <> Is_Null(DyDisp("n_period").Value, True) Then
            If j > 1 Then
               j = j + 1
            End If
         End If
         
         tmpTotOt = 0
         
         Va_Details.Row = j
         Va_Details.Col = 1
            Va_Details.Text = Is_Null(DyDisp("c_empno").Value, False)
         Va_Details.Col = 2
            Va_Details.Text = GetMonthYear(Is_Null(DyDisp("n_period").Value, True))
         Va_Details.Col = 3
            Va_Details.Text = Proper(Is_Null(DyDisp("c_name").Value, False)) & " " & Proper(Is_Null(DyDisp("c_othername").Value, False))
         Va_Details.Col = 4
            Va_Details.Text = Proper(Is_Null(DyDisp("c_companyname").Value, False))
         Va_Details.Col = 5
            Va_Details.Text = Proper(Is_Null(DyDisp("c_branch").Value, False))
         Va_Details.Col = 6
            Va_Details.Text = Proper(Is_Null(DyDisp("c_dept").Value, False))
         Va_Details.Col = 7
            Va_Details.Text = Proper(Is_Null(DyDisp("c_desig").Value, False))
         Va_Details.Col = 8
            If Is_Null(DyDisp("c_emptype").Value, False) = "S" Then
               Va_Details.Text = "Staff"
            Else
               Va_Details.Text = "Worker"
            End If
         Va_Details.Col = 9
            If Is_Null(DyDisp("c_stafftype").Value, False) = "F" Then
               Va_Details.Text = "Flat"
            Else
               Va_Details.Text = "OT"
            End If
         Va_Details.Col = 10
            If Is_Null(DyDisp("c_salarytype").Value, False) = "ML" Then
               Va_Details.Text = "Monthly"
            Else
               Va_Details.Text = "Hourly"
            End If
         Va_Details.Col = 11
            Va_Details.Text = ""
         
         ' fixed rate in master
         Va_Details.Col = 12
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_basic").Value, True), True)
         Va_Details.Col = 13
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_fixedrate2").Value, True), True)
         Va_Details.Col = 14
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_fixedrate3").Value, True), True)
         Va_Details.Col = 15
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_fixedrate4").Value, True), True)
         Va_Details.Col = 16
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_fixedrate5").Value, True), True)
         Va_Details.Col = 17
            Va_Details.Text = ""
         
         ' edf details
         Va_Details.Col = 18
            Va_Details.Text = Is_Null(DyDisp("c_edfcat").Value, False)
         Va_Details.Col = 19
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_edfamount").Value, True), True)
         Va_Details.Col = 20
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_preamount").Value, True), True)
         Va_Details.Col = 21
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_eduamount").Value, True), True)
         Va_Details.Col = 22
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_intamount").Value, True), True)
         Va_Details.Col = 23
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_othamount").Value, True), True)
         Va_Details.Col = 24
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_edftotal").Value, True), True)
         Va_Details.Col = 25
            Va_Details.Text = ""
            
         ' bank details
         Va_Details.Col = 26
            If Is_Null(DyDisp("c_paytype").Value, False) = "BA" Then
               Va_Details.Text = "Bank"
            Else
               Va_Details.Text = "Cash"
            End If
          Va_Details.Col = 27
            Va_Details.Text = Is_Null(DyDisp("c_bankshortname").Value, False)
          Va_Details.Col = 28
            Va_Details.Text = Is_Null(DyDisp("c_bankcode").Value, False)
          Va_Details.Col = 29
            Va_Details.Text = Is_Null(DyDisp("c_acctno").Value, False)
          Va_Details.Col = 30
            Va_Details.Text = Is_Null(DyDisp("c_itno").Value, False)
         Va_Details.Col = 31
            Va_Details.Text = ""
            
         ' hours
         Va_Details.Col = 32
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_wrkhrs").Value, True), True)
            vTotBasicWrkHrs = vTotBasicWrkHrs + Is_Null(DyDisp("n_wrkhrs").Value, True)
         Va_Details.Col = 33
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_lopdays").Value, True), True)
         Va_Details.Col = 34
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_latehrs").Value, True), True)
         Va_Details.Col = 35
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_othrs15").Value, True), True)
            vOtHrs15 = vOtHrs15 + Is_Null(DyDisp("n_othrs15").Value, True)
         Va_Details.Col = 36
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_othrs20").Value, True), True)
            vOtHrs2 = vOtHrs2 + Is_Null(DyDisp("n_othrs20").Value, True)
         Va_Details.Col = 37
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_othrs30").Value, True), True)
            vOtHrs3 = vOtHrs3 + Is_Null(DyDisp("n_othrs30").Value, True)
         Va_Details.Col = 38
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_sunphhrs20").Value, True), True)
            vSunPhHrs2 = vSunPhHrs2 + Is_Null(DyDisp("n_sunphhrs20").Value, True)
         Va_Details.Col = 39
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_sunphhrs30").Value, True), True)
            vSunPhHrs3 = vSunPhHrs3 + Is_Null(DyDisp("n_sunphhrs30").Value, True)
         Va_Details.Col = 40
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_othrs20").Value, True) + Is_Null(DyDisp("n_sunphhrs20").Value, True), True)
            vOtHrs20 = vOtHrs20 + Is_Null(DyDisp("n_othrs20").Value, True)
         Va_Details.Col = 41
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_othrs30").Value, True) + Is_Null(DyDisp("n_sunphhrs30").Value, True), True)
            vOtHrs30 = vOtHrs30 + Is_Null(DyDisp("n_othrs30").Value, True)
         Va_Details.Col = 42
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_totothrs").Value, True), True)
            vTotOTHrs = vTotOTHrs + Is_Null(DyDisp("n_totothrs").Value, True)
         Va_Details.Col = 43
            Va_Details.Text = ""
         
         ' leave hrs
         Va_Details.Col = 44
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_lophrs").Value, True), True)
         Va_Details.Col = 45
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_publicholiday").Value, True), True)
         Va_Details.Col = 46
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_localleave").Value, True), True)
         Va_Details.Col = 47
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_sickleave").Value, True), True)
         Va_Details.Col = 48
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_injuryleave").Value, True), True)
         Va_Details.Col = 49
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_prolongleave").Value, True), True)
         Va_Details.Col = 50
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_weddingleave").Value, True), True)
         Va_Details.Col = 51
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_matleave").Value, True), True)
         Va_Details.Col = 52
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_patleave").Value, True), True)
         Va_Details.Col = 53
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_compleave").Value, True), True)
         Va_Details.Col = 54
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_othleave").Value, True), True)
         Va_Details.Col = 55
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_no_travelallow").Value, True), True)
         Va_Details.Col = 56
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_no_mealallow").Value, True), True)
         Va_Details.Col = 57
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_no_nightallow").Value, True), True)
         Va_Details.Col = 58
            Va_Details.Text = ""
         
         ' salary amount
         Va_Details.Col = 59
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_earnbasic").Value, True), True)
            vBasic = vBasic + Is_Null(DyDisp("n_basic").Value, True)
         Va_Details.Col = 60
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_earnfr2").Value, True), True)
            vEarnFR2 = vEarnFR2 + Is_Null(DyDisp("n_earnfr2").Value, True)
         Va_Details.Col = 61
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_earnfr3").Value, True), True)
            vEarnFR3 = vEarnFR3 + Is_Null(DyDisp("n_earnfr3").Value, True)
         Va_Details.Col = 62
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_earnfr4").Value, True), True)
            vEarnFR4 = vEarnFR4 + Is_Null(DyDisp("n_earnfr4").Value, True)
         Va_Details.Col = 63
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_earnfr5").Value, True), True)
            vEarnFR5 = vEarnFR5 + Is_Null(DyDisp("n_earnfr5").Value, True)
         Va_Details.Col = 64
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_ot15").Value, True), True)
            vOt15 = vOt15 + Is_Null(DyDisp("n_ot15").Value, True)
         Va_Details.Col = 65
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_ot20").Value, True), True)
            vOt2 = vOt2 + Is_Null(DyDisp("n_ot20").Value, True)
         Va_Details.Col = 66
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_ot30").Value, True), True)
            vOt3 = vOt3 + Is_Null(DyDisp("n_ot30").Value, True)
         Va_Details.Col = 67
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_sunph20").Value, True), True)
            vSunPh2 = vSunPh2 + Is_Null(DyDisp("n_sunph20").Value, True)
         Va_Details.Col = 68
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_sunph30").Value, True), True)
            vSunPh3 = vSunPh3 + Is_Null(DyDisp("n_sunph30").Value, True)
         Va_Details.Col = 69
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_ot20").Value, True) + Is_Null(DyDisp("n_sunph20").Value, True), True)
            vOt20 = vOt20 + Is_Null(DyDisp("n_ot20").Value, True)
         Va_Details.Col = 70
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_ot30").Value, True) + Is_Null(DyDisp("n_sunph30").Value, True), True)
            vOt30 = vOt30 + Is_Null(DyDisp("n_ot30").Value, True)
         Va_Details.Col = 71
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_ot").Value, True), True)
            vTotOT = vTotOT + Is_Null(DyDisp("n_ot").Value, True)
         Va_Details.Col = 72
            Va_Details.Text = ""
         
         ' additional amount
         Va_Details.Col = 73
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_bonus").Value, True), True)
            vBonus = vBonus + Is_Null(DyDisp("n_bonus").Value, True)
         Va_Details.Col = 74
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_presbonus").Value, True), True)
            vPresBonus = vPresBonus + Is_Null(DyDisp("n_presbonus").Value, True)
         Va_Details.Col = 75
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_mealallow").Value, True), True)
            vMealAllow = vMealAllow + Is_Null(DyDisp("n_mealallow").Value, True)
         Va_Details.Col = 76
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_nightallow").Value, True), True)
            vNightAllow = vNightAllow + Is_Null(DyDisp("n_nightallow").Value, True)
         Va_Details.Col = 77
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_travel").Value, True), True)
            vTravel = vTravel + Is_Null(DyDisp("n_travel").Value, True)
         Va_Details.Col = 78
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_travel_tax").Value, True), True)
            vTravelTax = vTravelTax + Is_Null(DyDisp("n_travel_tax").Value, True)
         
         Va_Details.Col = 79
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_ph").Value, True), True)
         Va_Details.Col = 80
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_local").Value, True), True)
         Va_Details.Col = 81
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_sick").Value, True), True)
         Va_Details.Col = 82
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_injury").Value, True), True)
         Va_Details.Col = 83
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_prolong").Value, True), True)
         Va_Details.Col = 84
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_wedding").Value, True), True)
         Va_Details.Col = 85
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_maternity").Value, True), True)
         Va_Details.Col = 86
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_paternity").Value, True), True)
         Va_Details.Col = 87
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_compoff").Value, True), True)
         Va_Details.Col = 88
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_othleave").Value, True), True)
         Va_Details.Col = 89
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_totleave").Value, True), True)
            vTotLeave = vTotLeave + Is_Null(DyDisp("n_totleave").Value, True)
         
         Va_Details.Col = 90
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_cashadd").Value, True), True)
            vCashAdd = vCashAdd + Is_Null(DyDisp("n_cashadd").Value, True)
         Va_Details.Col = 91
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_eoybonus").Value, True), True)
            vEOYBonus = vEOYBonus + Is_Null(DyDisp("n_eoybonus").Value, True)
         Va_Details.Col = 92
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_othincome").Value, True), True)
            vOthIncome = vOthIncome + Is_Null(DyDisp("n_othincome").Value, True)
         Va_Details.Col = 93
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_totincome").Value, True), True)
            vTotIncome = vTotIncome + Is_Null(DyDisp("n_totincome").Value, True)
         Va_Details.Col = 94
            Va_Details.Text = ""
         
         ' deductions
         Va_Details.Col = 95
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_lop").Value, True), True)
            vLop = vLop + Is_Null(DyDisp("n_lop").Value, True)
         Va_Details.Col = 96
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_late").Value, True), True)
            vLate = vLate + Is_Null(DyDisp("n_late").Value, True)
         Va_Details.Col = 97
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_empnps").Value, True), True)
            vEmpNps = vEmpNps + Is_Null(DyDisp("n_empnps").Value, True)
         Va_Details.Col = 98
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_empepz").Value, True), True)
            vEmpEpz = vEmpEpz + Is_Null(DyDisp("n_empepz").Value, True)
         Va_Details.Col = 99
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_empewf").Value, True), True)
            vEmpEwf = vEmpEwf + Is_Null(DyDisp("n_empewf").Value, True)
         Va_Details.Col = 100
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_emplevy").Value, True), True)
            vEmpLevy = vEmpLevy + Is_Null(DyDisp("n_emplevy").Value, True)
         Va_Details.Col = 101
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_loan").Value, True), True)
            vLoan = vLoan + Is_Null(DyDisp("n_loan").Value, True)
         Va_Details.Col = 102
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_paye").Value, True), True)
            vPayee = vPayee + Is_Null(DyDisp("n_paye").Value, True)
         Va_Details.Col = 103
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_cashdeduct").Value, True), True)
            vCashDeduct = vCashDeduct + Is_Null(DyDisp("n_cashdeduct").Value, True)
         Va_Details.Col = 104
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_othdeduction").Value, True), True)
            vOthDeduction = vOthDeduction + Is_Null(DyDisp("n_othdeduction").Value, True)
         Va_Details.Col = 105
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_totdeduction").Value, True), True)
            vTotDeduction = vTotDeduction + Is_Null(DyDisp("n_totdeduction").Value, True)
         Va_Details.Col = 106
            Va_Details.Text = ""
         
         ' net
         Va_Details.Col = 107
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_net").Value, True), True)
            vNet = vNet + Is_Null(DyDisp("n_net").Value, True)
         Va_Details.Col = 108
            Va_Details.Text = ""
         
         ' company contrib
         Va_Details.Col = 109
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_carbenefit").Value, True), True)
         Va_Details.Col = 110
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_comnps").Value, True), True)
            vComNps = vComNps + Is_Null(DyDisp("n_comnps").Value, True)
         Va_Details.Col = 111
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_comepz").Value, True), True)
            vComEpz = vComEpz + Is_Null(DyDisp("n_comepz").Value, True)
         Va_Details.Col = 112
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_comewf").Value, True), True)
            vComEwf = vComEwf + Is_Null(DyDisp("n_comewf").Value, True)
         Va_Details.Col = 113
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_comlevy").Value, True), True)
            vComLevy = vComLevy + Is_Null(DyDisp("n_comlevy").Value, True)
         Va_Details.Col = 114
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_bonusincome").Value, True), True)
         Va_Details.Col = 115
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_bonusdeduction").Value, True), True)
         Va_Details.Col = 116
            Va_Details.Text = ""
         
         ' notes & coin
         Va_Details.Col = 117
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_1000").Value, True), True)
            v1000 = v1000 + Is_Null(DyDisp("n_1000").Value, True)
         Va_Details.Col = 118
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_500").Value, True), True)
            v500 = v500 + Is_Null(DyDisp("n_500").Value, True)
         Va_Details.Col = 119
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_200").Value, True), True)
            v200 = v200 + Is_Null(DyDisp("n_200").Value, True)
         Va_Details.Col = 120
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_100").Value, True), True)
            v100 = v100 + Is_Null(DyDisp("n_100").Value, True)
         Va_Details.Col = 121
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_50").Value, True), True)
            v50 = v50 + Is_Null(DyDisp("n_50").Value, True)
         Va_Details.Col = 122
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_10").Value, True), True)
            v10 = v10 + Is_Null(DyDisp("n_10").Value, True)
         Va_Details.Col = 123
            Va_Details.Text = ""
         
       
        tmpEmpNo = Is_Null(DyDisp("c_empno").Value, False)
        tmpPeriod = Is_Null(DyDisp("n_period").Value, True)
        
        DyDisp.MoveNext
     Next i
     Txtn_NoRows = DyDisp.RecordCount
     Va_Details.MaxRows = DyDisp.RecordCount + 3
     
        Va_Details.Row = Va_Details.DataRowCnt + 2
        Va_Details.Col = 3
           Va_Details.Text = "Total"
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 32
           Va_Details.Text = Spread_NumFormat(vTotBasicWrkHrs, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 35
           Va_Details.Text = Spread_NumFormat(vOtHrs15, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 36
           Va_Details.Text = Spread_NumFormat(vOtHrs2, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 37
           Va_Details.Text = Spread_NumFormat(vOtHrs3, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 38
           Va_Details.Text = Spread_NumFormat(vSunPhHrs2, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 39
           Va_Details.Text = Spread_NumFormat(vSunPhHrs3, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 40
           Va_Details.Text = Spread_NumFormat(vOtHrs20, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 41
           Va_Details.Text = Spread_NumFormat(vOtHrs30, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 42
           Va_Details.Text = Spread_NumFormat(vTotOTHrs, True)
           Va_Details.ForeColor = vbBlue
       
        Va_Details.Col = 59
           Va_Details.Text = Spread_NumFormat(vBasic, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 60
           Va_Details.Text = Spread_NumFormat(vEarnFR2, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 61
           Va_Details.Text = Spread_NumFormat(vEarnFR3, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 62
           Va_Details.Text = Spread_NumFormat(vEarnFR4, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 63
           Va_Details.Text = Spread_NumFormat(vEarnFR5, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 64
           Va_Details.Text = Spread_NumFormat(vOt15, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 65
           Va_Details.Text = Spread_NumFormat(vOt2, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 66
           Va_Details.Text = Spread_NumFormat(vOt3, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 67
           Va_Details.Text = Spread_NumFormat(vSunPh2, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 68
           Va_Details.Text = Spread_NumFormat(vSunPh3, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 69
           Va_Details.Text = Spread_NumFormat(vOt20, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 70
           Va_Details.Text = Spread_NumFormat(vOt30, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 71
           Va_Details.Text = Spread_NumFormat(vTotOT, True)
           Va_Details.ForeColor = vbBlue
       
        Va_Details.Col = 73
           Va_Details.Text = Spread_NumFormat(vBonus, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 74
           Va_Details.Text = Spread_NumFormat(vPresBonus, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 75
           Va_Details.Text = Spread_NumFormat(vMealAllow, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 76
           Va_Details.Text = Spread_NumFormat(vNightAllow, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 77
           Va_Details.Text = Spread_NumFormat(vTravel, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 78
           Va_Details.Text = Spread_NumFormat(vTravelTax, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 89
           Va_Details.Text = Spread_NumFormat(vTotLeave, True)
           Va_Details.ForeColor = vbBlue

        Va_Details.Col = 90
           Va_Details.Text = Spread_NumFormat(vCashAdd, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 91
           Va_Details.Text = Spread_NumFormat(vEOYBonus, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 92
           Va_Details.Text = Spread_NumFormat(vOthIncome, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 93
           Va_Details.Text = Spread_NumFormat(vTotIncome, True)
           Va_Details.ForeColor = vbBlue

        Va_Details.Col = 95
           Va_Details.Text = Spread_NumFormat(vLop, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 96
           Va_Details.Text = Spread_NumFormat(vLate, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 97
           Va_Details.Text = Spread_NumFormat(vEmpNps, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 98
           Va_Details.Text = Spread_NumFormat(vEmpEwf, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 99
           Va_Details.Text = Spread_NumFormat(vEmpEpz, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 100
           Va_Details.Text = Spread_NumFormat(vEmpLevy, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 101
           Va_Details.Text = Spread_NumFormat(vLoan, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 102
           Va_Details.Text = Spread_NumFormat(vPayee, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 103
           Va_Details.Text = Spread_NumFormat(vCashDeduct, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 104
           Va_Details.Text = Spread_NumFormat(vOthDeduction, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 105
           Va_Details.Text = Spread_NumFormat(vTotDeduction, True)
           Va_Details.ForeColor = vbBlue
         
        Va_Details.Col = 107
           Va_Details.Text = Spread_NumFormat(vNet, True)
           Va_Details.ForeColor = vbBlue
     
        Va_Details.Col = 110
           Va_Details.Text = Spread_NumFormat(vComNps, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 111
           Va_Details.Text = Spread_NumFormat(vComEwf, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 112
           Va_Details.Text = Spread_NumFormat(vComEpz, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 113
           Va_Details.Text = Spread_NumFormat(vComLevy, True)
           Va_Details.ForeColor = vbBlue
     
        Va_Details.Col = 117
           Va_Details.Text = Spread_NumFormat(v1000, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 118
           Va_Details.Text = Spread_NumFormat(v500, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 119
           Va_Details.Text = Spread_NumFormat(v200, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 120
           Va_Details.Text = Spread_NumFormat(v100, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 121
           Va_Details.Text = Spread_NumFormat(v50, True)
           Va_Details.ForeColor = vbBlue
        Va_Details.Col = 122
           Va_Details.Text = Spread_NumFormat(v10, True)
           Va_Details.ForeColor = vbBlue
     
  
  Else
     MsgBox "No details found", vbInformation, "Information"
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

Private Sub Txtc_EmployeeName_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim Search As New Search.MyClass, SerVar
 
   If KeyCode = vbKeyF2 Then
      Search.Query = "select c_empno EmpNo, c_name Name, c_othername OtherName, c_dept Dept, " & _
                   "c_desig Desig, c_branch Branch, c_emptype Type " & _
                   "from pr_emp_mst "
      Search.CheckFields = "EmpNo, Name"
      Search.ReturnField = "EmpNo, Name"
      SerVar = Search.Search(, , CON)
      If Len(Search.col1) <> 0 Then
         Txtc_EmployeeName = Search.col2 & Space(100) & Search.col1
      End If
   End If
End Sub

Private Sub Txtc_EmployeeName_Validate(Cancel As Boolean)
 Dim RsChk As New ADODB.Recordset
 Dim i As Integer
  If Trim(Txtc_EmployeeName) <> "" Then
     Set RsChk = Nothing
     g_Sql = "select c_empno, c_name, c_othername, c_company, c_branch, c_dept, c_desig, c_emptype, " & _
             "d_doj, d_dol, c_stafftype, c_daywork from pr_emp_mst " & _
             "where c_rec_sta = 'A' and c_empno = '" & Trim(Right(Trim(Txtc_EmployeeName), 7)) & "'"
     RsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
     If RsChk.RecordCount > 0 Then
        Txtc_EmployeeName = Is_Null(RsChk("c_name").Value, False) & " " & Is_Null(RsChk("c_othername").Value, False) & Space(100) & Is_Null(RsChk("c_empno").Value, False)
        Call DisplayComboCompany(Me, Is_Null(RsChk("c_company").Value, False))
        Call DisplayComboBranch(Me, Is_Null(RsChk("c_branch").Value, False))
        Call DisplayComboDept(Me, Is_Null(RsChk("c_dept").Value, False))
        Call DisplayComboDesig(Me, Is_Null(RsChk("c_desig").Value, False))
        Call DisplayComboEmpType(Me, Is_Null(RsChk("c_emptype").Value, False))
     Else
        MsgBox "Employee not found. Press <F2> to select.", vbInformation, "Information"
        Cancel = True
     End If
  End If
End Sub

Private Sub Spread_Lock()
 Dim i As Long
   
   For i = 1 To Va_Details.MaxCols
       Va_Details.Row = -1
       Va_Details.Col = i
          Va_Details.Lock = True
   Next i
End Sub

Private Sub Txtc_MonthFrom_KeyPress(KeyAscii As Integer)
    Call OnlyNumeric(Txtc_MonthFrom, KeyAscii, 2)
End Sub

Private Sub Txtc_MonthFrom_Validate(Cancel As Boolean)
   If Trim(Txtc_MonthFrom) <> "" Then
      If Len(Trim(Txtc_MonthFrom)) < 2 Then
         Txtc_MonthFrom = "0" & Trim(Txtc_MonthFrom)
      End If
      If Val(Txtc_MonthFrom) <= 0 Or Val(Txtc_MonthFrom) > 13 Then
         MsgBox "Not a valid month", vbInformation, "Information"
         Txtc_MonthFrom.SetFocus
         Cancel = True
      End If
   End If
   Call Assign_PayPeriod
End Sub

Private Sub Txtc_YearFrom_KeyPress(KeyAscii As Integer)
    Call OnlyNumeric(Txtc_YearFrom, KeyAscii, 4)
End Sub

Private Sub Txtc_YearFrom_Validate(Cancel As Boolean)
   If Len(Txtc_YearFrom) > 0 Then
      If Len(Txtc_YearFrom) <> 4 Then
         MsgBox "Not a valid year", vbInformation, "Information"
         Txtc_YearFrom.SetFocus
         Cancel = True
      End If
   End If
   Call Assign_PayPeriod
End Sub

Private Sub Txtc_MonthTo_KeyPress(KeyAscii As Integer)
    Call OnlyNumeric(Txtc_MonthTo, KeyAscii, 2)
End Sub

Private Sub Txtc_MonthTo_Validate(Cancel As Boolean)
   If Trim(Txtc_MonthTo) <> "" Then
      If Len(Trim(Txtc_MonthTo)) < 2 Then
         Txtc_MonthTo = "0" & Trim(Txtc_MonthTo)
      End If
      If Val(Txtc_MonthTo) <= 0 Or Val(Txtc_MonthTo) > 13 Then
         MsgBox "Not a valid month", vbInformation, "Information"
         Txtc_MonthTo.SetFocus
         Cancel = True
      End If
   End If
   Call Assign_PayPeriod
End Sub

Private Sub Txtc_YearTo_KeyPress(KeyAscii As Integer)
    Call OnlyNumeric(Txtc_YearTo, KeyAscii, 4)
End Sub

Private Sub Txtc_YearTo_Validate(Cancel As Boolean)
   If Len(Txtc_YearTo) > 0 Then
      If Len(Txtc_YearTo) <> 4 Then
         MsgBox "Not a valid year", vbInformation, "Information"
         Txtc_YearTo.SetFocus
         Cancel = True
      End If
   End If
   Call Assign_PayPeriod
End Sub

Private Sub Assign_PayPeriod()
    If Val(Txtc_MonthTo) = 0 Then
       Txtc_MonthTo = Txtc_MonthFrom
    End If
    
    If Val(Txtc_YearTo) = 0 Or Val(Txtc_YearTo) < Val(Txtc_YearFrom) Then
       Txtc_YearTo = Txtc_YearFrom
    End If
    
    If Trim(Txtc_MonthFrom) <> "" And Trim(Txtc_YearFrom) <> "" Then
       vPayPeriodFrom = Is_Null(Format(Trim(Txtc_YearFrom), "0000") & Trim(Txtc_MonthFrom), True)
    End If
    If Trim(Txtc_MonthTo) <> "" And Trim(Txtc_YearTo) <> "" Then
       vPayPeriodTo = Is_Null(Format(Trim(Txtc_YearTo), "0000") & Trim(Txtc_MonthTo), True)
    End If
    
    If vPayPeriodFrom = vPayPeriodTo Then
       vFromMonthName = MonthName(Val(Txtc_MonthFrom), True)
       vToMonthName = ""
    Else
       vFromMonthName = MonthName(Val(Txtc_MonthFrom), True)
       vToMonthName = MonthName(Val(Txtc_MonthTo), True)
    End If
    
    
End Sub

Private Sub Combo_Load()
  Dim rsCombo As New ADODB.Recordset
  Dim i As Integer
    
    Call LoadComboCompany(Me)
    Call LoadComboBranch(Me)
    Call LoadComboDept(Me)
    Call LoadComboDesig(Me)
    Call LoadComboEmpType(Me)
    
    
    Set rsCombo = Nothing
    g_Sql = "select distinct c_nationality from pr_emp_mst where c_rec_sta = 'A' and c_nationality is not null order by c_nationality "
    rsCombo.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    Cmb_Nationality.Clear
    Cmb_Nationality.AddItem ""
    Cmb_Nationality.AddItem "All Expatriate"
    For i = 0 To rsCombo.RecordCount - 1
        Cmb_Nationality.AddItem rsCombo("c_nationality").Value
        rsCombo.MoveNext
    Next i
End Sub

Private Function GetMonthYear(ByVal vPeriod As Long) As String
  Dim vMonth As Integer, vYear As String
      
    vMonth = Right(Trim(Str(vPeriod)), 2)
    vYear = Right(Left(Trim(Str(vPeriod)), 4), 2)
    
    If vMonth = 13 Then
       GetMonthYear = "Bonus - " & Trim(vYear)
    Else
       GetMonthYear = MonthName(vMonth, True) & " - " & Trim(vYear)
    End If
End Function

Private Sub Btn_GridDefault_Click()
On Error GoTo ErrSave
  Dim vGridView As String
  Dim i As Integer
  
    For i = 1 To Va_Details.MaxCols
        Va_Details.Row = -1
        Va_Details.Col = i
           If Va_Details.ColHidden = True Then
              vGridView = vGridView & Trim(Str(i)) & ","
           End If
    Next i
    
    CON.BeginTrans
    
    Set rs = Nothing
    g_Sql = "select * from pr_user_grid_set where c_user_id = '" & g_UserName & "' and c_screen_id = '" & g_SalaryScr & "'"
    rs.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
    If rs.RecordCount = 0 Then
       rs.AddNew
       rs("c_user_id").Value = g_UserName
       rs("c_screen_id").Value = g_SalaryScr
       rs("c_gridview").Value = vGridView
       rs.Update
    Else
       rs("c_gridview").Value = vGridView
       rs.Update
    End If
    
    CON.CommitTrans
    
    MsgBox "Grid column display is set successfully. The grid will be display as you set for next time when you open this screen.", vbInformation, "Information"
    
  Exit Sub
     
ErrSave:
     CON.RollbackTrans
     Screen.MousePointer = vbDefault
     MsgBox "Error while Saving - " + Err.Description, vbCritical, "Critical"
End Sub


Private Sub Spread_Hide_Check()
  Dim RsChk As New ADODB.Recordset
  Dim vGridView As String
    
    Set RsChk = Nothing
    g_Sql = "select c_gridview from pr_user_grid_set where c_user_id = '" & g_UserName & "' and c_screen_id = '" & g_SalaryScr & "'"
    RsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If RsChk.RecordCount = 0 Then
       Call Spread_ColHide
    Else
       vGridView = Is_Null(RsChk("c_gridview").Value, False)
       If Trim(vGridView) = "" Then
          Call Spread_ColHide
       Else
          Call Spread_ColUser_Hide(vGridView)
       End If
    End If
    
End Sub

Private Sub Spread_ColUser_Hide(ByVal vGridView As String)
  Dim i As Integer
    
    For i = 1 To Va_Details.MaxCols
        Va_Details.Row = -1
        Va_Details.Col = i
           Va_Details.ColHidden = IsColHide(vGridView, i)
    Next i
  
End Sub

Private Function IsColHide(ByVal vGridView As String, ByVal vCol As Integer) As Boolean
  Dim i As Integer
  Dim vColHide As Variant
  
    vColHide = Split(vGridView, ",")
    If UBound(vColHide) < 0 Then Exit Function
       
    For i = 0 To UBound(vColHide)
        If vColHide(i) = vCol Then
           IsColHide = True
           Exit For
        End If
    Next i
End Function

Private Sub Spread_ColHide()
  Dim i As Integer, j As Integer
    ' always hide
    Va_Details.Row = -1
    Va_Details.Col = 11
       Va_Details.ColHidden = True
    Va_Details.Col = 17
       Va_Details.ColHidden = True
    Va_Details.Col = 25
       Va_Details.ColHidden = True

    ' hrs
    Va_Details.Col = 31
       Va_Details.ColHidden = True
    Va_Details.Col = 33
       Va_Details.ColHidden = True
    Va_Details.Col = 36
       Va_Details.ColHidden = True
    Va_Details.Col = 37
       Va_Details.ColHidden = True
    Va_Details.Col = 38
       Va_Details.ColHidden = True
    Va_Details.Col = 39
       Va_Details.ColHidden = True
    Va_Details.Col = 42
       Va_Details.ColHidden = True
    Va_Details.Col = 43
       Va_Details.ColHidden = True

    ' leave hrs
    Va_Details.Col = 58
       Va_Details.ColHidden = True

    ' amount
    Va_Details.Col = 61
       Va_Details.ColHidden = True
    Va_Details.Col = 62
       Va_Details.ColHidden = True
    Va_Details.Col = 63
       Va_Details.ColHidden = True
    Va_Details.Col = 65
       Va_Details.ColHidden = True
    Va_Details.Col = 66
       Va_Details.ColHidden = True
    Va_Details.Col = 67
       Va_Details.ColHidden = True
    Va_Details.Col = 68
       Va_Details.ColHidden = True
    Va_Details.Col = 71
       Va_Details.ColHidden = True
    Va_Details.Col = 72
       Va_Details.ColHidden = True
    
    ' deduction
    Va_Details.Col = 94
       Va_Details.ColHidden = True
       
    Va_Details.Col = 106
       Va_Details.ColHidden = True
    Va_Details.Col = 108
       Va_Details.ColHidden = True
    Va_Details.Col = 116
       Va_Details.ColHidden = True

    For j = 1 To 5
        Va_Details.Row = -1
        Va_Details.Col = 12 + j
           Va_Details.ColHidden = True
    Next j


    ' first - employee
    For j = 0 To 6
        Va_Details.Row = -1
        Va_Details.Col = 6 + j
           If Chk_Hide(0).Value = 1 Then
              Va_Details.ColHidden = True
           Else
              Va_Details.ColHidden = False
           End If

           If j = 4 Then
              Va_Details.Col = 8
                 Va_Details.ColHidden = False
           End If
    Next j

    ' second - edf
    For j = 0 To 5
        Va_Details.Row = -1
        Va_Details.Col = 19 + j
           If Chk_Hide(1).Value = 1 Then
              Va_Details.ColHidden = True
           Else
              Va_Details.ColHidden = False
           End If
    Next j

    ' thired - bank
    For j = 0 To 3
        Va_Details.Row = -1
        Va_Details.Col = 26 + j
           If Chk_Hide(2).Value = 1 Then
              Va_Details.ColHidden = True
           Else
              Va_Details.ColHidden = False
           End If
    Next j

    ' forth  - hours
    For j = 0 To 13
        Va_Details.Row = -1
        Va_Details.Col = 44 + j
           If Chk_Hide(3).Value = 1 Then
              Va_Details.ColHidden = True
           Else
              Va_Details.ColHidden = False
           End If
    Next j

    ' fifth - leave pay
    For j = 0 To 9
        Va_Details.Row = -1
        Va_Details.Col = 79 + j
           If Chk_Hide(4).Value = 1 Then
              Va_Details.ColHidden = True
           Else
              Va_Details.ColHidden = False
           End If
    Next j
    
    ' six - comp contrib
    For j = 0 To 6
        Va_Details.Row = -1
        Va_Details.Col = 109 + j
           If Chk_Hide(5).Value = 1 Then
              Va_Details.ColHidden = True
           Else
              Va_Details.ColHidden = False
           End If
    Next j

    ' sevanth - notes
    For j = 0 To 6
        Va_Details.Row = -1
        Va_Details.Col = 117 + j
           If Chk_Hide(6).Value = 1 Then
              Va_Details.ColHidden = True
           Else
              Va_Details.ColHidden = False
           End If
    Next j
End Sub

Private Sub Va_Details_DblClick(ByVal Col As Long, ByVal Row As Long)
   Call SpreadColSort(Va_Details, Col, Row)
End Sub

Private Sub Btn_Hide_Click()
  Dim i As Integer
  
    For i = Va_Details.SelBlockCol To Va_Details.SelBlockCol2
        If i > 2 Then
           Va_Details.Row = -1
           Va_Details.Col = i
             Va_Details.ColHidden = True
        End If
    Next i
End Sub

Private Sub Btn_Restore_Click()
  Dim i As Integer
    For i = 1 To Va_Details.MaxCols
        Va_Details.Row = -1
        Va_Details.Col = i
           Va_Details.ColHidden = False
    Next i
    Call Spread_ColHide
End Sub

Private Sub Btn_Export_Click()
On Error GoTo Err_Export
  Dim SelFor As String
  Dim vFileName As String
  
     If Va_Details.DataRowCnt = 0 Then
        Exit Sub
     End If
  
     If MsgBox("Do you want to Export to CSV file?", vbYesNo + vbDefaultButton2, "Confirmation") = vbNo Then
        Exit Sub
     End If
  
     Screen.MousePointer = vbHourglass
    
     CON.BeginTrans
        Save_Pr_Export_Csv_Rep
     CON.CommitTrans
     
     
     ComDialog.FileName = "EmpDtl.CSV"
     ComDialog.ShowSave
     vFileName = Trim(ComDialog.FileName)
     
     If vFileName = "SalDtl.CSV" Then  'user cancel the export to csv file in save dialog box.
        Screen.MousePointer = vbDefault
        Exit Sub
     ElseIf Trim(Right(Trim(vFileName), 4)) <> ".CSV" Then
        vFileName = vFileName & ".CSV"
     End If

     Call CsvFileExport_Process(vFileName)

     Screen.MousePointer = vbDefault
     
     MsgBox "Transfered Successfully", vbInformation, "Information"
     
  Exit Sub

Err_Export:
    CON.RollbackTrans
    Screen.MousePointer = vbDefault
    MsgBox "Error while Generating - " + Err.Description, vbInformation, "Information"
End Sub


Private Sub Save_Pr_Export_Csv_Rep()
  Dim i As Long, j As Long, Ctr As Long
  Dim vStr As String, vEmpNo As String
  
    Ctr = 0
    g_Sql = "truncate table pr_export_csv "
    CON.Execute (g_Sql)
    
    Set rs = Nothing
    g_Sql = "select * from pr_export_csv where 1=2"
    rs.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
    
    For i = 0 To Va_Details.DataRowCnt
        Va_Details.Row = i
        Va_Details.Col = 1
           If Trim(Va_Details.Text) <> "" Then
              vStr = "": vEmpNo = ""
              Ctr = Ctr + 1
              Va_Details.Row = i
              Va_Details.Col = 1
                 vEmpNo = Trim(Va_Details.Text)
                 vStr = Trim(Va_Details.Text)
              
              For j = 2 To Va_Details.MaxCols
                  Va_Details.Row = i
                  Va_Details.Col = j
                     If Va_Details.ColHidden = False Then
                        If j = 0 Then 'combo
                           vStr = vStr & "," & Trim(Left(Trim(Va_Details.Text), 50))

                        ElseIf j = 0 Then 'option box
                          If i = 0 Then
                              vStr = vStr & "," & Trim(Va_Details.Text)
                           Else
                              vStr = vStr & "," & IIf(Va_Details.Value = 1, "Yes", "No")
                           End If

                        Else
                           vStr = vStr & "," & Trim(Replace(Va_Details.Text, ",", ""))
                        End If
                     End If
              Next j
              
              rs.AddNew
              rs("n_seq").Value = Ctr
              rs("c_refno").Value = vEmpNo
              rs("c_csv").Value = vStr
              rs("c_usr_id").Value = g_UserName
              rs("d_created").Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
              rs.Update
           End If
    Next i
End Sub

Private Sub Spread_Header_Name()
   Call Get_FixedRate_Name
   Va_Details.Row = 0
   Va_Details.Col = 13
      Va_Details.Text = g_FixedRateName2
   Va_Details.Col = 14
      Va_Details.Text = g_FixedRateName3
   Va_Details.Col = 15
      Va_Details.Text = g_FixedRateName4
   Va_Details.Col = 16
      Va_Details.Text = g_FixedRateName5

   Va_Details.Col = 60
      Va_Details.Text = g_FixedRateName2
   Va_Details.Col = 61
      Va_Details.Text = g_FixedRateName3
   Va_Details.Col = 62
      Va_Details.Text = g_FixedRateName4
   Va_Details.Col = 63
      Va_Details.Text = g_FixedRateName5
End Sub

Private Sub Pr_Payslip_Emp_Gen()
  Dim RsChk As New ADODB.Recordset
  Dim tmpStr As String, vEmpNo1 As String, vEmpNo2 As String, vPeriod As String
  Dim i As Integer, j As Integer
  
    g_Sql = "Truncate Table PR_SALARY_PAYSLIP_REP "
    CON.Execute g_Sql
    
    tmpStr = " and a.n_period >= " & vPayPeriodFrom & " and a.n_period <= " & vPayPeriodTo

    If Trim(Cmb_Company) <> "" Then
       tmpStr = tmpStr & " and a.c_company = '" & Right(Trim(Cmb_Company), 7) & "'"
    End If
    If Trim(Cmb_Branch) <> "" Then
       tmpStr = tmpStr & " and a.c_branch = '" & Trim(Cmb_Branch) & "'"
    End If
    If Trim(Cmb_Dept) <> "" Then
       tmpStr = tmpStr & " and a.c_dept = '" & Trim(Cmb_Dept) & "'"
    End If
    If Trim(Cmb_Desig) <> "" Then
       tmpStr = tmpStr & " and a.c_desig = '" & Trim(Cmb_Desig) & "'"
    End If
    If Trim(Cmb_EmpType) <> "" Then
       tmpStr = tmpStr & " and a.c_emptype = '" & Right(Trim(Cmb_EmpType), 1) & "'"
    End If
    If Trim(Cmb_Nationality) <> "" Then
       If Trim(Cmb_Nationality) = "All Expatriate" Then
          tmpStr = tmpStr & " and b.c_expatriate = 'Y' "
       Else
          tmpStr = tmpStr & " and b.c_nationality = '" & Trim(Cmb_Nationality) & "'"
       End If
    End If
    If Trim(Txtc_EmployeeName) <> "" Then
       tmpStr = tmpStr & " and a.c_empno = '" & Trim(Right(Trim(Txtc_EmployeeName), 7)) & "'"
    End If
    
    g_Sql = "select a.c_empno, a.n_period, a.c_paytype from pr_salary_mst a, pr_emp_mst b " & _
            "where a.c_empno = b.c_empno " & tmpStr & _
            "order by a.n_period, a.c_branch, a.c_dept, a.c_empno "
    RsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If RsChk.RecordCount = 0 Then
       Exit Sub
    End If
    
    j = 0
    RsChk.MoveFirst
    Do While Not RsChk.EOF
        vEmpNo1 = "": vEmpNo2 = "": vPeriod = ""
        
        vPeriod = Is_Null(RsChk("n_period").Value, False)
        vEmpNo1 = Is_Null(RsChk("c_empno").Value, False)
        RsChk.MoveNext
        
        If Not RsChk.EOF Then
           vEmpNo2 = Is_Null(RsChk("c_empno").Value, False)
           RsChk.MoveNext
        Else
           vEmpNo2 = vEmpNo1
        End If
        
        j = j + 1
        g_Sql = "Insert into Pr_Salary_Payslip_Rep (n_slno, n_period, c_empno1, c_empno2) " & _
                "values (" & j & ", '" & vPeriod & "', '" & vEmpNo1 & "', '" & vEmpNo2 & "') "
        CON.Execute g_Sql
        
        If RsChk.EOF Then
           Exit Sub
        End If
    Loop
  
End Sub

Private Function SalaryMaster_RepFilter() As String
  Dim vF1 As String, vF2 As String, vF3 As String, vF4 As String, vF5 As String
  Dim SelFor As String
  
     vF1 = "": vF2 = "": vF3 = "": vF4 = "": vF5 = ""
     SelFor = ""
     
     If Right(Trim(Cmb_Company), 7) <> "" Then
        vF1 = "{V_PR_SALARY_MST.C_COMPANY}='" & Right(Trim(Cmb_Company), 7) & "'"
     End If
     
     If Trim(Cmb_Branch) <> "" Then
        vF2 = "{V_PR_SALARY_MST.C_BRANCH}='" & Trim(Cmb_Branch) & "'"
     End If
     
     If Trim(Cmb_Dept) <> "" Then
        vF3 = "{V_PR_SALARY_MST.C_DEPT}='" & Trim(Cmb_Dept) & "'"
     End If
     
     If Trim(Cmb_Desig) <> "" Then
        vF4 = "{V_PR_SALARY_MST.C_DESIG}='" & Trim(Cmb_Desig) & "'"
     End If
     
     If Trim(Cmb_EmpType) <> "" Then
        vF5 = "{V_PR_SALARY_MST.C_EMPTYPE}='" & Trim(Left(Trim(Cmb_EmpType), 10)) & "'"
     End If
     
     SelFor = ReportFilterOption(vF1, vF2, vF3, vF4, vF5)
     vF1 = "": vF2 = "": vF3 = "": vF4 = "": vF5 = ""
     
     If Trim(Txtc_EmployeeName) <> "" Then
        vF1 = "{V_PR_SALARY_MST.C_EMPNO}='" & Trim(Right(Trim(Txtc_EmployeeName), 7)) & "'"
     End If
     
     If Trim(Cmb_Nationality) <> "" Then
        vF2 = "{V_PR_EMP_MST.C_NATIONALITY}='" & Trim(Cmb_Nationality) & "'"
     End If
     
     vF3 = "{V_PR_SALARY_MST.N_PERIOD} >= " & vPayPeriodFrom & " AND {V_PR_SALARY_MST.N_PERIOD} <= " & vPayPeriodTo
     
     SalaryMaster_RepFilter = ReportFilterOption(SelFor, vF1, vF2, vF3)

End Function

Private Function getPeriodTitle() As String
  Dim tmpStr As String
  
    tmpStr = vFromMonthName & " " & Format(Txtc_YearFrom, "0000")
    
    If vPayPeriodFrom <> vPayPeriodTo Then
       tmpStr = tmpStr & "  to  " & vToMonthName & " " & Format(Txtc_YearTo, "0000")
    End If
    getPeriodTitle = tmpStr
End Function
