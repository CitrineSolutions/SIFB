VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.MDIForm Mdi_Ta_HrPay 
   BackColor       =   &H8000000C&
   Caption         =   "HRPTA"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   4680
   Icon            =   "Mdi_Ta_HrPay.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "Mdi_Ta_HrPay.frx":424A
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CRY1 
      Left            =   600
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   600
      Top             =   480
   End
   Begin MSComctlLib.StatusBar sts_bar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   2820
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8819
            MinWidth        =   8819
            Text            =   "Sugar Insurance Fund Board  -  Time and Attendance System"
            TextSave        =   "Sugar Insurance Fund Board  -  Time and Attendance System"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2646
            MinWidth        =   2646
            TextSave        =   "19/02/2018"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "6:39 PM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnu_PayMas 
      Caption         =   "Masters"
      Begin VB.Menu mnu_PayMasSub 
         Caption         =   "Company Master"
         Index           =   1
      End
      Begin VB.Menu mnu_PayMasSub 
         Caption         =   "Bank Master"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_PayMasSub 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnu_PayMasSub 
         Caption         =   "Clock Card Shift"
         Index           =   4
      End
      Begin VB.Menu mnu_PayMasSub 
         Caption         =   "Shift Structure"
         Index           =   5
      End
      Begin VB.Menu mnu_PayMasSub 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnu_PayMasSub 
         Caption         =   "Pay Period"
         Index           =   7
      End
      Begin VB.Menu mnu_PayMasSub 
         Caption         =   "Pay Structure"
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_PayMasSub 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnu_PayMasSub 
         Caption         =   "Leave Master"
         Index           =   10
      End
      Begin VB.Menu mnu_PayMasSub 
         Caption         =   "Leave Entitle Slabs"
         Index           =   11
      End
      Begin VB.Menu mnu_PayMasSub 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnu_PayMasSub 
         Caption         =   "Public Holiday"
         Index           =   13
      End
      Begin VB.Menu mnu_PayMasSub 
         Caption         =   "EDF Structure"
         Index           =   14
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnu_HR 
      Caption         =   "HR"
      Begin VB.Menu mnu_HRSub 
         Caption         =   "Employee Master"
         Index           =   1
      End
      Begin VB.Menu mnu_HRSub 
         Caption         =   "Employee Master Multi-Updates"
         Index           =   2
      End
      Begin VB.Menu mnu_HRSub 
         Caption         =   "-"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_HRSub 
         Caption         =   "Transport Planning"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_HRSub 
         Caption         =   "-"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_HRSub 
         Caption         =   "Loan/Advance Details"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_HRSub 
         Caption         =   "Additional Income/Deduction Details"
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_HRSub 
         Caption         =   "-"
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_HRSub 
         Caption         =   "Payroll Process"
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_HRSub 
         Caption         =   "Salary Details"
         Index           =   10
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_HRSub 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu mnu_HRSub 
         Caption         =   "Leave Entry Details"
         Index           =   12
      End
      Begin VB.Menu mnu_HRSub 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnu_HRSub 
         Caption         =   "Leave Encash / Adjustment"
         Index           =   14
      End
   End
   Begin VB.Menu mnu_Clock 
      Caption         =   "Clocking"
      Begin VB.Menu mnu_ClockSub 
         Caption         =   "Data Capture && Process"
         Index           =   1
      End
      Begin VB.Menu mnu_ClockSub 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnu_ClockSub 
         Caption         =   "Shift Change"
         Index           =   3
      End
      Begin VB.Menu mnu_ClockSub 
         Caption         =   "Attendance Details"
         Index           =   4
      End
   End
   Begin VB.Menu mnu_Report 
      Caption         =   "Reports"
      Begin VB.Menu mnu_ReportSub 
         Caption         =   "Employee Information"
         Index           =   1
         Begin VB.Menu mnu_EmpInfoRepSub 
            Caption         =   "List"
            Index           =   1
         End
         Begin VB.Menu mnu_EmpInfoRepSub 
            Caption         =   "New Entry Details"
            Index           =   2
         End
         Begin VB.Menu mnu_EmpInfoRepSub 
            Caption         =   "Left Details"
            Index           =   3
         End
      End
      Begin VB.Menu mnu_ReportSub 
         Caption         =   "Leave Info"
         Index           =   2
      End
      Begin VB.Menu mnu_ReportSub 
         Caption         =   "-"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_ReportSub 
         Caption         =   "Loan Details"
         Index           =   4
         Visible         =   0   'False
         Begin VB.Menu mnu_LoanRepSub 
            Caption         =   "List"
            Index           =   1
         End
         Begin VB.Menu mnu_LoanRepSub 
            Caption         =   "Outstanding"
            Index           =   2
         End
      End
      Begin VB.Menu mnu_ReportSub 
         Caption         =   "Additional Income/Dedtn "
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_ReportSub 
         Caption         =   "-"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_ReportSub 
         Caption         =   "Salary Details"
         Index           =   7
         Visible         =   0   'False
         Begin VB.Menu mnu_RepSalarySub 
            Caption         =   "Payment Mode Details"
            Index           =   1
         End
         Begin VB.Menu mnu_RepSalarySub 
            Caption         =   "Pay Component Details"
            Index           =   2
         End
         Begin VB.Menu mnu_RepSalarySub 
            Caption         =   "Note && Coin Analysis"
            Index           =   3
         End
         Begin VB.Menu mnu_RepSalarySub 
            Caption         =   "Cash Payment Register"
            Index           =   4
         End
      End
      Begin VB.Menu mnu_ReportSub 
         Caption         =   "-"
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_ReportSub 
         Caption         =   "Social Security"
         Index           =   9
         Visible         =   0   'False
         Begin VB.Menu mnu_RepSocSecSub 
            Caption         =   "Contribution Details"
            Index           =   1
         End
         Begin VB.Menu mnu_RepSocSecSub 
            Caption         =   "PAYE Details"
            Index           =   2
         End
         Begin VB.Menu mnu_RepSocSecSub 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnu_RepSocSecSub 
            Caption         =   "PAYE CSV File Format"
            Index           =   4
         End
         Begin VB.Menu mnu_RepSocSecSub 
            Caption         =   "NPF CSV File Format"
            Index           =   5
         End
         Begin VB.Menu mnu_RepSocSecSub 
            Caption         =   "EPZ CSV File Format"
            Index           =   6
         End
         Begin VB.Menu mnu_RepSocSecSub 
            Caption         =   "EPZ Loan CSV File Format"
            Index           =   7
         End
         Begin VB.Menu mnu_RepSocSecSub 
            Caption         =   "-"
            Index           =   8
         End
         Begin VB.Menu mnu_RepSocSecSub 
            Caption         =   "Emoulment Statement && CSV File Format"
            Index           =   9
         End
      End
      Begin VB.Menu mnu_ReportSub 
         Caption         =   "-"
         Index           =   10
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_ReportSub 
         Caption         =   "Bank Transfer"
         Index           =   11
         Visible         =   0   'False
         Begin VB.Menu mnu_BankTranSub 
            Caption         =   "MCB"
            Index           =   1
         End
         Begin VB.Menu mnu_BankTranSub 
            Caption         =   "SBM"
            Index           =   2
         End
         Begin VB.Menu mnu_BankTranSub 
            Caption         =   "Barclays"
            Index           =   3
         End
      End
      Begin VB.Menu mnu_ReportSub 
         Caption         =   "Statement of Emoluments"
         Index           =   12
         Visible         =   0   'False
         Begin VB.Menu mnu_RepSubStaEmoStaff 
            Caption         =   "Old Format"
            Index           =   1
         End
         Begin VB.Menu mnu_RepSubStaEmoStaff 
            Caption         =   "New Format (MRA)"
            Index           =   2
         End
      End
   End
   Begin VB.Menu mnu_Admin 
      Caption         =   "Admin"
      Begin VB.Menu mnu_AdminSub 
         Caption         =   "Pay Period Close"
         Index           =   1
      End
      Begin VB.Menu mnu_AdminSub 
         Caption         =   "Pay Period Re-Open"
         Index           =   2
      End
      Begin VB.Menu mnu_AdminSub 
         Caption         =   "-"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_AdminSub 
         Caption         =   "Clocking  Period Open"
         Index           =   4
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnu_Util 
      Caption         =   "&Utilities"
      Begin VB.Menu mnu_Utility_Sub 
         Caption         =   "Login"
         Index           =   1
      End
      Begin VB.Menu mnu_Utility_Sub 
         Caption         =   "Change Password"
         Index           =   2
      End
      Begin VB.Menu mnu_Utility_Sub 
         Caption         =   "User Creation"
         Index           =   3
      End
      Begin VB.Menu mnu_Utility_Sub 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnu_Utility_Sub 
         Caption         =   "A&bout The Software"
         Index           =   5
      End
   End
   Begin VB.Menu mnu_exit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "Mdi_Ta_HrPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private vCloseMsg As Boolean

Private Sub MDIForm_Load()
  Dim tmpFlag As String
  If App.PrevInstance Then
     MsgBox "This Application is already running", vbInformation, "Information"
     tmpFlag = "T"
     End
  Else
    Call StartDB
    sts_bar1.Panels(2).Text = g_ClientCode
    sts_bar1.Panels(3).Text = g_Database
  End If
  
  If tmpFlag <> "T" Then
     vCloseMsg = True
  End If

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  If (MsgBox("Do You want to exit the Application", vbYesNo + vbQuestion, "Exit") = vbYes) Then
     End
  Else
     Cancel = True
  End If
End Sub


Private Sub mnu_AdminSub_Click(Index As Integer)
  Dim Clock_DataProcess As New frm_Clock_DataProcess
  
   If Index = 1 Then
      If ChkUsrRight Then
         frm_Period_Closed.mnuOption = "C"
         frm_Period_Closed.Show
      End If
   ElseIf Index = 2 Then
      If ChkUsrRight Then
         frm_Period_Closed.mnuOption = "O"
         frm_Period_Closed.Show
      End If
      
   ElseIf Index = 4 Then
       If ChkUsrRight Then
          Clock_DataProcess.mnuOption = "CPO"
          Clock_DataProcess.Show
       End If
   End If
  
End Sub

Private Sub mnu_BankTranSub_Click(Index As Integer)
  Dim Report_Generate As New frm_Report_Generate
    
    If ChkUsrRight Then
       If Index = 1 Then
          Report_Generate.RepName = "salmcbpaytr"
          Report_Generate.Show
       ElseIf Index = 2 Then
          Report_Generate.RepName = "salscbpaytr"
          Report_Generate.Show
       ElseIf Index = 3 Then
          Report_Generate.RepName = "salbarpaytr"
          Report_Generate.Show
       End If
    End If
End Sub


Private Sub mnu_ClockSub_Click(Index As Integer)
  Dim Clock_Emp As New frm_Clock_Emp
  Dim Clock_DataProcess As New frm_Clock_DataProcess
  
    If Index = 1 Then
       If ChkScreenRight("frm_Clock_DataProcess_DWH") Then
          Clock_DataProcess.mnuOption = "DWH"
          Clock_DataProcess.Show
       End If
    ElseIf Index = 3 Then
       Clock_Emp.mnuOption = "S"
       Clock_Emp.Show
    ElseIf Index = 4 Then
       Clock_Emp.mnuOption = "A"
       Clock_Emp.Show
    End If
End Sub

Private Sub mnu_exit_Click()
   Unload Me
End Sub

Private Sub mnu_HRSub_Click(Index As Integer)
  Dim Clock_DataProcess As New frm_Clock_DataProcess
  Dim Emp_Master As New frm_Emp_Master
  Dim Emp_Upd As New frm_Emp_Upd
'  Dim AddPay_Details As New frm_AddPay_Details
'  Dim Salary_Details As New frm_Salary_Details
  Dim Leave_Adj As New frm_Leave_Adj
  Dim Leave_Entry As New frm_Leave_Entry
  
    If Index = 1 Then
       If ChkScreenRight("frm_emp_master") Then
          Emp_Master.Show
       End If
       
    ElseIf Index = 2 Then
       If ChkScreenRight("frm_Emp_Upd") Then
          Emp_Upd.mnuOption = "E"
          Emp_Upd.Show
       End If
    
    ElseIf Index = 4 Then
       If ChkScreenRight("frm_Emp_Upd") Then
          Emp_Upd.mnuOption = "T"
          Emp_Upd.Show
       End If
       
    ElseIf Index = 6 Then
'       If ChkScreenRight("frm_Loan") Then
'          frm_Loan.Show
'       End If
    ElseIf Index = 7 Then
'       If ChkScreenRight("frm_AddPay_Details") Then
'          AddPay_Details.Show
'       End If
    
    ElseIf Index = 9 Then
       If ChkScreenRight("frm_Clock_DataProcess_PRP") Then
          Clock_DataProcess.mnuOption = "PRP"
          Clock_DataProcess.Show
       End If
    ElseIf Index = 10 Then
'       If ChkScreenRight("frm_Salary_Details") Then
'          Salary_Details.Show
'       End If
       
    
    ElseIf Index = 12 Then
       If ChkScreenRight("frm_Leave_Entry") Then
          frm_Leave_Entry.Show
       End If
    
    ElseIf Index = 14 Then
       If ChkScreenRight("frm_Leave_Adj") Then
          frm_Leave_Adj.Show
       End If
    End If
End Sub

Private Sub mnu_MisSub_Click(Index As Integer)

End Sub

Private Sub mnu_LoanRepSub_Click(Index As Integer)
  Dim Report_Filter As New frm_Report_Filter

    If Not g_Admin Then
       If ChkScreenRight("frm_Loan") Then
          Exit Sub
       End If
    End If

    If Index = 1 Then
       Report_Filter.RepName = "loandtl"
       Report_Filter.Show
    ElseIf Index = 2 Then
       Report_Filter.RepName = "loanosdtl"
       Report_Filter.Show
    End If
End Sub

Private Sub mnu_PayMasSub_Click(Index As Integer)

    If Index = 1 Then
       If ChkScreenRight("frm_Company") Then
          frm_Company.Show
       End If
    ElseIf Index = 2 Then
'       If ChkScreenRight("frm_BankMaster") Then
'          frm_BankMaster.Show
'       End If
    ElseIf Index = 4 Then
       If ChkScreenRight("frm_Shifts") Then
          frm_Shifts.Show
       End If
    ElseIf Index = 5 Then
       If ChkScreenRight("frm_Shift_Structure") Then
          frm_Shift_Structure.Show
       End If
        
    ElseIf Index = 7 Then
       If ChkScreenRight("frm_PayPeriod") Then
          frm_PayPeriod.Show
       End If
    ElseIf Index = 8 Then
'       If ChkScreenRight("frm_Pay_Struture") Then
'          frm_Pay_Struture.Show
'       End If
    
    ElseIf Index = 10 Then
       If ChkScreenRight("frm_LeaveMaster") Then
          frm_LeaveMaster.Show
       End If
    ElseIf Index = 11 Then
       If ChkScreenRight("frm_Leave_Allotment") Then
          frm_Leave_Allotment.Show
       End If
    
    ElseIf Index = 13 Then
       If ChkScreenRight("frm_Holiday") Then
          frm_Holiday.Show
       End If
    ElseIf Index = 14 Then
'       If ChkScreenRight("frm_EDF_Master") Then
'          frm_EDF_Master.Show
'       End If
    End If
End Sub

Private Sub mnu_ReportSub_Click(Index As Integer)
   Dim Report_Filter  As New frm_Report_Filter

    If Index = 2 Then
       Report_Filter.RepName = "leavedtl"
       Report_Filter.Show
    
    ElseIf Index = 5 Then
       Report_Filter.RepName = "addincded"
       Report_Filter.Show
    End If
End Sub

Private Sub mnu_RepSalarySub_Click(Index As Integer)
  Dim Report_Filter As New frm_Report_Filter
  
    If Index = 1 Then
       Report_Filter.RepName = "salpaymodedtl"
       Report_Filter.Show
    ElseIf Index = 2 Then
       Report_Filter.RepName = "salpaycompdtl"
       Report_Filter.Show
    ElseIf Index = 3 Then
       Report_Filter.RepName = "salnotecoin"
       Report_Filter.Show
    ElseIf Index = 4 Then
       Report_Filter.RepName = "salcashreg"
       Report_Filter.Show
    End If
End Sub

Private Sub mnu_RepSocSecSub_Click(Index As Integer)
   Dim Report_Filter As New frm_Report_Filter
   Dim Report_Generate As New frm_Report_Generate

      If Index = 1 Then
         Report_Filter.RepName = "salsocsecdtl"
         Report_Filter.Show
      ElseIf Index = 2 Then
         Report_Filter.RepName = "salpayedtl"
         Report_Filter.Show
      
      ElseIf Index = 4 Then
         Report_Generate.RepName = "salpayecsv"
         Report_Generate.Show
      ElseIf Index = 5 Then
         Report_Generate.RepName = "salnpfcsv"
         Report_Generate.Show
      ElseIf Index = 6 Then
         Report_Generate.RepName = "salepzcsv"
         Report_Generate.Show
      ElseIf Index = 7 Then
         Report_Generate.RepName = "salepzloancsv"
         Report_Generate.Show

      ElseIf Index = 9 Then
         Report_Generate.RepName = "salemotran"
         Report_Generate.Show
      End If
End Sub

Private Sub mnu_Utility_Sub_Click(Index As Integer)
    If Index = 1 Then
       frm_Login.Show 1
       sts_bar1.Panels(2).Text = g_ClientCode
       sts_bar1.Panels(3).Text = g_Database
       sts_bar1.Panels(4).Text = UCase(g_UserName)
    
    ElseIf Index = 2 Then
       frm_ChangePwd.mnuOption = "UC"
       frm_ChangePwd.Show 1
    
    ElseIf Index = 3 Then
       If ChkUsrRight Then
          frm_Users.Show
       End If
       
    ElseIf Index = 5 Then
       frm_About.Show
    End If
    
End Sub

Private Sub mnu_EmpInfoRepSub_Click(Index As Integer)
   Dim Report_Filter  As New frm_Report_Filter

     If Index = 1 Then
       Report_Filter.RepName = "empmstlist"
       Report_Filter.Show
    ElseIf Index = 2 Then
       Report_Filter.RepName = "empmstdoj"
       Report_Filter.Show
    ElseIf Index = 3 Then
       Report_Filter.RepName = "empmstdol"
       Report_Filter.Show
     End If
End Sub
Private Sub Timer1_Timer()
    If Timer1.tag <> "D" And vCloseMsg Then
       frm_Login.Show 1, Mdi_Ta_HrPay
       Timer1.tag = "D"
       sts_bar1.Panels(2).Text = g_ClientCode
       sts_bar1.Panels(3).Text = g_Database
       sts_bar1.Panels(4).Text = UCase(g_UserName)
       Timer1.Interval = 0
       Timer1.Enabled = False
    End If
End Sub
