VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{C3A136DA-B937-492B-968D-A437638F7AAB}#1.0#0"; "CS_DateControl.ocx"
Begin VB.Form frm_Clock_DataProcess 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12315
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   12315
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
      Height          =   1665
      Left            =   375
      TabIndex        =   16
      Top             =   645
      Width           =   11415
      Begin VB.ComboBox Cmb_Desig 
         Height          =   315
         Left            =   8610
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   390
         Width           =   2475
      End
      Begin VB.ComboBox Cmb_EmpType 
         Height          =   315
         Left            =   8610
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   720
         Width           =   2460
      End
      Begin VB.ComboBox Cmb_Company 
         Height          =   315
         Left            =   4425
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   390
         Width           =   2505
      End
      Begin VB.ComboBox Cmb_Branch 
         Height          =   315
         Left            =   4425
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   705
         Width           =   2505
      End
      Begin VB.TextBox Txtc_EmployeeName 
         Height          =   300
         Left            =   8610
         TabIndex        =   9
         Top             =   1050
         Width           =   2445
      End
      Begin VB.ComboBox Cmb_Dept 
         Height          =   315
         Left            =   4425
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
         Width           =   525
      End
      Begin VB.TextBox Txtc_Year 
         Height          =   300
         Left            =   1980
         TabIndex        =   1
         Top             =   405
         Width           =   705
      End
      Begin CS_DateControl.DateControl Txtd_FromDate 
         Height          =   345
         Left            =   1440
         TabIndex        =   2
         Top             =   780
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   609
      End
      Begin CS_DateControl.DateControl Txtd_ToDate 
         Height          =   345
         Left            =   1440
         TabIndex        =   3
         Top             =   1110
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   609
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
         Left            =   7485
         TabIndex        =   25
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
         Left            =   7485
         TabIndex        =   24
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
         Left            =   3300
         TabIndex        =   23
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
         Left            =   3765
         TabIndex        =   22
         Top             =   750
         Width           =   570
      End
      Begin VB.Label Label1 
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
         TabIndex        =   21
         Top             =   1155
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
         Left            =   7710
         TabIndex        =   20
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
         Left            =   3300
         TabIndex        =   19
         Top             =   1050
         Width           =   1035
      End
      Begin VB.Label Label2 
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
         TabIndex        =   18
         Top             =   840
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
         TabIndex        =   17
         Top             =   435
         Width           =   570
      End
   End
   Begin VB.Frame Fme_Generate 
      Height          =   2145
      Left            =   375
      TabIndex        =   14
      Top             =   2280
      Width           =   11430
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
         Height          =   450
         Left            =   4245
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1380
         Width           =   1155
      End
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
         Left            =   5865
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1380
         Width           =   1155
      End
      Begin ComctlLib.ProgressBar ProBar 
         Height          =   300
         Left            =   2535
         TabIndex        =   10
         Top             =   705
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   529
         _Version        =   327682
         Appearance      =   1
      End
      Begin MSComDlg.CommonDialog comDialog 
         Left            =   285
         Top             =   1320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Lbl_Info 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Text"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   11190
      End
   End
   Begin VB.Label lbl_scr_name 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Date Capture && Process"
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
      Left            =   540
      TabIndex        =   13
      Top             =   270
      Width           =   11070
   End
   Begin VB.Shape shp_scr_name 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   360
      Top             =   240
      Width           =   11415
   End
End
Attribute VB_Name = "frm_Clock_DataProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mnuOption As String
Private vPayPeriod As Long, vPrevPayPeriod As Long
Private vPayPeriodFrom As Date, vPayPeriodTo As Date
Private vF1 As String, vF2 As String, vF3 As String, vF4 As String, vF5 As String, vF6 As String
Private vNetNorHrs, vNetOT1, vNetOT2, vNetOT3 As Double
Private vProBarValue As Integer

Private Sub Form_Load()
   
   If mnuOption = "CPO" Then
      lbl_scr_name.Caption = "Clocking Period Open Process"
      
   ElseIf mnuOption = "DWH" Then
      lbl_scr_name.Caption = "Daily Work Hours Process"
   
   ElseIf mnuOption = "PRP" Then
      Txtd_FromDate.Enabled = False
      Txtd_ToDate.Enabled = False
      lbl_scr_name.Caption = "Payroll Process"
   End If
   
   Lbl_Info.Caption = ""
      
   Call Combo_Load
   Call TGControlProperty(Me)
   ProBar.Value = 0
End Sub

Private Sub Btn_Ok_Click()
 Dim i As Long
 
   Lbl_Info.Caption = ""
   ProBar.Value = 0
   
   If Not Gen_Check() Then
      Exit Sub
   End If
    
   If mnuOption = "CPO" Then
      Call Clocking_Period_Open_Process
      
   ElseIf mnuOption = "DWH" Then
      Call Daily_Work_Hours_Process
      
   ElseIf mnuOption = "PRP" Then
      Call WeeklyHours_Payroll_Process
      
   End If
    
   Btn_Ok.Enabled = True
   Btn_Exit.Enabled = True
   Screen.MousePointer = vbDefault
   Txtd_FromDate.SetFocus
   
End Sub

Private Sub Btn_Exit_Click()
  Unload Me
End Sub

Private Sub Clocking_Period_Open_Process()
On Error GoTo Err_Flag
  
    vProBarValue = 0
    
    If MsgBox("Do you want to Process? It may take time to complete the process. ", vbYesNo, "Confirmation") = vbYes Then
       Screen.MousePointer = vbHourglass
              
       Lbl_Info.Caption = "Leave entitle, earned and allotment process"
       ProBar.Value = 5
       DoEvents
      
       g_Sql = "HR_LEAVE_ALLOT_PROC " & Val(Txtc_Year)
       CON.Execute g_Sql
       ProBar.Value = 15
       DoEvents
             
       Lbl_Info.Caption = "Leave re-update process."
       DoEvents
       g_Sql = "HR_LEAVE_REUPDATE_PROC " & Val(Txtc_Year)
       CON.Execute g_Sql
       ProBar.Value = 30
       
             
       Lbl_Info.Caption = "Clocking period open process."
       DoEvents
       If Trim(Txtc_EmployeeName) = "" Then
         g_Sql = "HR_CLOCK_EMP_UPD '" & Is_Date(Txtd_FromDate.Text, "S") & "', '" & Is_Date(Txtd_ToDate.Text, "S") & "', 'A'"
       Else
         g_Sql = "HR_CLOCK_EMP_UPD '" & Is_Date(Txtd_FromDate.Text, "S") & "', '" & Is_Date(Txtd_ToDate.Text, "S") & "', '" & Trim(Right(Trim(Txtc_EmployeeName), 7)) & "'"
       End If
       CON.Execute g_Sql
       ProBar.Value = 100
       
       Lbl_Info.Caption = "Process completed"
       DoEvents
    
       Screen.MousePointer = vbDefault
       MsgBox "Successfully Completed", vbInformation, "Information"
    End If
  
  Exit Sub

Err_Flag:
    Screen.MousePointer = vbDefault
    MsgBox "Error while Processing " & vbCrLf & Err.Description
End Sub

Private Sub Daily_Work_Hours_Process()
On Error GoTo Err_Flag
  Dim rsChk As New ADODB.Recordset
  Dim i As Integer
  Dim vPathFileName As String, tmpPeriod As String
  Dim tmpDate As Date
  
    vProBarValue = 0
    
    If MsgBox("Do you want to Process? It may take time to complete the process. ", vbYesNo, "Confirmation") = vbYes Then
       Screen.MousePointer = vbHourglass
              
       vProBarValue = 0: ProBar.Value = 0
       DoEvents
            
       vF1 = Is_Date(Txtd_FromDate.Text, "S")
       vF2 = Is_Date(Txtd_ToDate.Text, "S")
       vF3 = IIf(Trim(Cmb_Company) <> "", Trim(Right(Trim(Cmb_Company), 7)), "A")
       vF4 = IIf(Trim(Cmb_Branch) <> "", Trim(Cmb_Branch), "A")
       vF5 = IIf(Trim(Cmb_Dept) <> "", Trim(Cmb_Dept), "A")
       vF6 = IIf(Trim(Txtc_EmployeeName) <> "", Trim(Right(Trim(Txtc_EmployeeName), 7)), "A")
      
       Lbl_Info.Caption = "Clocking period open process"
       DoEvents
          g_Sql = "HR_CLOCK_EMP_UPD '" & vF1 & "', '" & vF2 & "', '" & vF3 & "', '" & vF4 & "', '" & vF5 & "', '" & vF6 & "'"
          CON.Execute g_Sql
       ProBar.Value = 10
       DoEvents
       
             
       Lbl_Info.Caption = "Clocking In && Out Timing process"
       DoEvents
          Call Update_Clock_InOut(tmpDate)
       ProBar.Value = 30
       DoEvents
       
       
       For tmpDate = CDate(Txtd_FromDate.Text) To CDate(Txtd_ToDate.Text)
           Lbl_Info.Caption = "Alternative shift and next day clock out process on " & Format(tmpDate, "dd/mm/yyyy")
           DoEvents
           
           Call Update_Alternate_Shift(tmpDate)
           
           Call Update_Clock_NextDay_Out(tmpDate)
           
           vProBarValue = ProBar.Value + Round(100 / (CDate(Txtd_ToDate.Text) - CDate(Txtd_FromDate.Text) + 1))
           vProBarValue = IIf(vProBarValue <= 100, vProBarValue, 100)
           ProBar.Value = vProBarValue
           DoEvents
       
       Next tmpDate

       vProBarValue = 0: ProBar.Value = 0
       DoEvents
       

       For tmpDate = CDate(Txtd_FromDate.Text) To CDate(Txtd_ToDate.Text)
           Lbl_Info.Caption = "Daily work hour process of " & Format(tmpDate, "dd/mm/yyyy")
           DoEvents
          
           Call Update_WorkHrs_Details(tmpDate)
          
           vProBarValue = ProBar.Value + Round(100 / (CDate(Txtd_ToDate.Text) - CDate(Txtd_FromDate.Text) + 1))
           vProBarValue = IIf(vProBarValue <= 100, vProBarValue, 100)
           ProBar.Value = vProBarValue
           DoEvents
       Next tmpDate
       
       
       ' to update Presabs flag. P, WO, PH
       ' to update WO before join and after left
       Lbl_Info.Caption = "Processing Staus Flag Updation "
       ProBar.Value = 70
       DoEvents
          vF1 = Is_Date(Txtd_FromDate.Text, "S")
          vF2 = Is_Date(Txtd_ToDate.Text, "S")
          g_Sql = "HR_CLOCK_PRESABS_UPD '" & vF1 & "', '" & vF2 & "', '" & vF3 & "', '" & vF4 & "', '" & vF5 & "', '" & vF6 & "'"
          CON.Execute g_Sql
       ProBar.Value = 75
       DoEvents

       
       Lbl_Info.Caption = "Leave entitle, earned and allotment process."
       DoEvents
          g_Sql = "HR_LEAVE_ALLOT_PROC " & Val(Txtc_Year)
          CON.Execute g_Sql
       ProBar.Value = 85
       DoEvents
       
       
       Lbl_Info.Caption = "Leave entry update process."
       DoEvents
          g_Sql = "HR_LEAVEENTRY_DTL_UPD "
          CON.Execute g_Sql
       ProBar.Value = 90
       DoEvents

       
       Lbl_Info.Caption = "Leave re-update process."
       DoEvents
          g_Sql = "HR_LEAVE_REUPDATE_PROC " & Val(Txtc_Year)
          CON.Execute g_Sql
       ProBar.Value = 100
       DoEvents
             
       Lbl_Info.Caption = "Process completed"
       DoEvents
       
    Else
       ProBar.Value = 0
       Exit Sub
    End If
    
    Screen.MousePointer = vbDefault
    MsgBox "Successfully Completed", vbInformation, "Information"
  Exit Sub

Err_Flag:
    Screen.MousePointer = vbDefault
    MsgBox "Error while Processing " & vbCrLf & Err.Description
End Sub

Private Sub Update_Clock_InOut(ByVal vDate As Date)
  Dim rsChk As New ADODB.Recordset
  Dim tmpEmpNo As String, tmpDate As Date
  Dim i As Integer, Ctr As Integer, NoClock As Integer, vClMin As Integer
  Dim Time As Double, tmpTime As Double
  Dim aTime(12) As Double
  
  Dim tmpTimeDiff
  
    i = 0: Ctr = 0: NoClock = 0: Time = 0: tmpTime = 0
    vClMin = 0
    
    Set rsChk = Nothing
    g_Sql = "select max(clmin) clmin from pr_clock_shift"
    rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If rsChk.RecordCount > 0 Then
       If Is_Null(rsChk("clmin").Value, True) = 0 Then
          vClMin = 10
       Else
          vClMin = Is_Null(rsChk("clmin").Value, True)
       End If
    Else
      vClMin = 10
    End If
    
    Set rsChk = Nothing
    g_Sql = "Select a.c_empno, a.d_date,  replace(hitclock.timelog,':','.') n_time " & _
            "from pr_clock_emp a, hitfpta.dbo.employee hitemp, hitfpta.dbo.personallog hitclock, pr_emp_mst c " & _
            "where a.c_empno = hitemp.employeeid and hitemp.fingerprintid = hitclock.fingerprintid and " & _
            "a.d_date = hitclock.datelog and a.c_empno = c.c_empno and c.c_rec_sta = 'A' " ' and a.d_date = '" & Is_Date(vDate, "S") & "'"
            
    g_Sql = g_Sql & " and a.d_date >= '" & Is_Date(Txtd_FromDate.Text, "S") & "' and a.d_date <= '" & Is_Date(Txtd_ToDate.Text, "S") & "'"
            
    If Trim(Cmb_Company) <> "" Then
       g_Sql = g_Sql & " and c.c_company = '" & Trim(Right(Trim(Cmb_Company), 7)) & "' "
    End If
    If Trim(Cmb_Branch) <> "" Then
       g_Sql = g_Sql & " and c.c_branch = '" & Trim(Cmb_Branch) & "' "
    End If
    If Trim(Cmb_Dept) <> "" Then
       g_Sql = g_Sql & " and c.c_dept = '" & Trim(Cmb_Dept) & "' "
    End If
    If Trim(Cmb_Desig) <> "" Then
       g_Sql = g_Sql & " and c.c_desig = '" & Trim(Cmb_Desig) & "' "
    End If
    If Trim(Cmb_EmpType) <> "" Then
       g_Sql = g_Sql & " and c.c_emptype = '" & Trim(Cmb_EmpType) & "' "
    End If
    If Trim(Txtc_EmployeeName) <> "" Then
       g_Sql = g_Sql & " and c.c_empno = '" & Trim(Right(Trim(Txtc_EmployeeName), 7)) & "' "
    End If
    
    g_Sql = g_Sql & " order by a.c_empno, a.d_date, hitclock.timelog "
    
    rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If rsChk.RecordCount <= 0 Then
       Exit Sub
    End If
    
    rsChk.MoveFirst
    
    Do While Not rsChk.EOF
       tmpEmpNo = Trim(rsChk("c_empno").Value)
       tmpDate = rsChk("d_date").Value
       
       For i = 0 To 12
           aTime(i) = 0
       Next i
       
       Do While tmpEmpNo = Trim(rsChk("c_empno").Value) And tmpDate = rsChk("d_date").Value
          Time = Is_Null(rsChk("n_time").Value, True)
         
          If aTime(1) = 0 Then
             aTime(1) = Time
             i = 1
          Else
             For i = 1 To 12
                If aTime(i) = 0 Then
                   aTime(i) = Time
                   Exit For
                ElseIf aTime(i) = Time Then
                   Exit For
                ElseIf Abs(TimeToMins(aTime(i)) - TimeToMins(Time)) <= vClMin Then '10 mins check
                   Exit For
                ElseIf aTime(i) > Time Then
                   tmpTime = aTime(i)
                   aTime(i) = Time
                   Time = tmpTime
                End If
             Next i
          End If
          
          rsChk.MoveNext
          If rsChk.EOF Then
             Exit Do
          End If
       Loop
       
       NoClock = i

       g_Sql = "Update pr_clock_emp Set " & _
               "n_noclock = " & NoClock & ", n_noclock_day = " & NoClock & ", " & _
               "n_time1 = " & aTime(1) & ", n_time2 = " & aTime(2) & ", n_time3 = " & aTime(3) & ", " & _
               "n_time4 = " & aTime(4) & ", n_time5 = " & aTime(5) & ", n_time6 = " & aTime(6) & ", " & _
               "n_time7 = " & aTime(7) & ", n_time8 = " & aTime(8) & ", n_time9 = " & aTime(9) & ", " & _
               "n_time10 = " & aTime(10) & ", n_time11 = " & aTime(11) & ", n_time12 = " & aTime(12) & " " & _
               "where c_empno = '" & tmpEmpNo & "' and d_date = '" & Is_Date(tmpDate, "S") & "' "
       
       CON.Execute g_Sql
    Loop
  
End Sub

Private Sub Update_Alternate_Shift(ByVal vDate As Date)
  Dim rsChk As New ADODB.Recordset
  Dim i As Integer
  Dim Time, StartHrs, EndHrs, AltStartHrs, AltEndHrs As Double
  Dim ShiftDiff, AltShiftDiff As Double
  Dim vShiftOld, vShift As String, UpdFlag As Boolean
    
    Set rsChk = Nothing
    g_Sql = "select a.c_empno, b.d_date, b.c_shift, b.n_time1, " & _
            "f.c_shiftcode, f.starthrs, f.endhrs, " & _
            "c.c_wocode, d.starthrs wostarthrs, d.endhrs woendhrs, " & _
            "c.c_altcode, e.starthrs altstarthrs, e.endhrs altendhrs " & _
            "from pr_emp_mst a, pr_clock_emp b, pr_shiftstructure_dtl c " & _
            "    left outer join pr_clock_shift d on c.c_wocode = d.c_shiftcode " & _
            "    left outer join pr_clock_shift e on c.c_altcode = e.c_shiftcode, " & _
            "pr_clock_shift f " & _
            "where a.c_empno = b.c_empno and a.c_shiftcode = c.c_code and " & _
            "b.n_wkday = c.n_wkday and c.c_shiftcode = f.c_shiftcode and a.c_rec_sta = 'A' and " & _
            "b.c_flag = 'B' and (b.n_workhrs > 0 or b.n_time1 > 0) and " & _
            "IsNull(b.c_sh_flag,'A') <> 'U' and b.d_date = '" & Is_Date(vDate, "S") & "' "
    
    If Trim(Cmb_Company) <> "" Then
       g_Sql = g_Sql & " and a.c_company  = '" & Trim(Right(Trim(Cmb_Company), 7)) & "' "
    End If
    If Trim(Cmb_Branch) <> "" Then
       g_Sql = g_Sql & " and a.c_branch = '" & Trim(Cmb_Branch) & "' "
    End If
    If Trim(Cmb_Dept) <> "" Then
       g_Sql = g_Sql & " and a.c_dept = '" & Trim(Cmb_Dept) & "' "
    End If
    If Trim(Cmb_Desig) <> "" Then
       g_Sql = g_Sql & " and a.c_desig = '" & Trim(Cmb_Desig) & "' "
    End If
    If Trim(Cmb_EmpType) <> "" Then
       g_Sql = g_Sql & " and a.c_emptype = '" & Trim(Cmb_EmpType) & "' "
    End If
    If Trim(Txtc_EmployeeName) <> "" Then
       g_Sql = g_Sql & " and a.c_empno = '" & Trim(Right(Trim(Txtc_EmployeeName), 7)) & "' "
    End If
    
    g_Sql = g_Sql & " order by a.c_empno, b.d_date "
    
    rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If rsChk.RecordCount <= 0 Then
       Exit Sub
    End If
    
    rsChk.MoveFirst
    Do While Not rsChk.EOF
       Time = 0: StartHrs = 0: EndHrs = 0:  AltStartHrs = 0: AltEndHrs = 0
       ShiftDiff = 0:  AltShiftDiff = 0
       vShiftOld = "": vShift = "": UpdFlag = False
       
       Time = TimeToMins(Is_Null(rsChk("n_time1").Value, True))
       StartHrs = Is_Null(rsChk("starthrs").Value, True)
       EndHrs = Is_Null(rsChk("endhrs").Value, True)
       vShiftOld = Is_Null(rsChk("c_shift").Value, False)
       vShift = Is_Null(rsChk("c_shift").Value, False)
       
       If vShift = "" Then ' Assign Defult shift code as Shift.
          vShift = Is_Null(rsChk("c_shiftcode").Value, False)
          UpdFlag = True
       ElseIf vShift = "WO" Then ' Assign WO shift code as Shift.
          vShift = Is_Null(rsChk("c_wocode").Value, False)
          StartHrs = Is_Null(rsChk("wostarthrs").Value, True)
          EndHrs = Is_Null(rsChk("woendhrs").Value, True)
          UpdFlag = True
       End If
       
       If Abs(Time - StartHrs) > 480 Then  'Basic Condition. Shift start hrs and arr time should be above 8 Hrs
          AltStartHrs = Is_Null(rsChk("altstarthrs").Value, True)
          AltEndHrs = Is_Null(rsChk("altendhrs").Value, True)
          
          ShiftDiff = Abs(StartHrs - Time)
          If AltStartHrs > 0 Then
             AltShiftDiff = Abs(AltStartHrs - Time)
          End If
          
          If AltShiftDiff > 0 Then
             If AltShiftDiff < ShiftDiff Then
                vShift = Is_Null(rsChk("c_altcode").Value, False)
                UpdFlag = True
             End If
          End If
       End If
       
       If vShift = vShiftOld Then
          UpdFlag = False
       End If
       
       If UpdFlag Then
          g_Sql = "Update pr_clock_emp Set c_shift = '" & vShift & "' " & _
                  "Where c_empno = '" & Is_Null(rsChk("c_empno").Value, False) & "' and " & _
                  "d_date = '" & Is_Date(vDate, "S") & "' and " & _
                  "IsNull(c_flag,'A') = 'B' and IsNull(c_sh_flag,'A') <> 'U' "
          CON.Execute g_Sql
       End If
       
       rsChk.MoveNext
    Loop
End Sub

Private Sub Update_Clock_NextDay_Out(ByVal vDate As Date)
On Error GoTo Err_Chk
  Dim rsChk As New ADODB.Recordset
  Dim i As Integer, Ctr As Integer
  Dim tmpEmpNo As String, tmpDate As Date
  Dim NoClock As Integer, vOutTime As Double
  Dim aTime(12) As Double, aCurDayTime(12), aNextDayTime(12) As Double
  
  Dim vShStartHrs, vShEndHrs, vShMaxHrs As Double
  
    i = 0: Ctr = 0: NoClock = 0: vOutTime = 0
    
    vShStartHrs = 0: vShEndHrs = 0: vShMaxHrs = 0
    
    Set rsChk = Nothing
    g_Sql = "Select a.c_empno, a.d_date, a.n_noclock, " & _
            "a.n_time1, a.n_time2, a.n_time3, a.n_time4, a.n_time5, a.n_time6, " & _
            "a.n_time7, a.n_time8, a.n_time9, a.n_time10, a.n_time11, a.n_time12, " & _
            "b.starthrs, b.endhrs, b.maxhrs " & _
            "from pr_clock_emp a, pr_clock_shift b, pr_emp_mst c " & _
            "where a.c_shift = b.c_shiftcode and a.c_empno = c.c_empno and c.c_rec_sta = 'A' and " & _
            "a.d_date >= '" & Is_Date(vDate, "S") & "' and " & _
            "a.d_date <= '" & Is_Date(CDate(vDate) + 1, "S") & "' "
    
    If Trim(Cmb_Company) <> "" Then
       g_Sql = g_Sql & " and c.c_company = '" & Trim(Right(Trim(Cmb_Company), 7)) & "' "
    End If
    If Trim(Cmb_Branch) <> "" Then
       g_Sql = g_Sql & " and c.c_branch = '" & Trim(Cmb_Branch) & "' "
    End If
    If Trim(Cmb_Dept) <> "" Then
       g_Sql = g_Sql & " and c.c_dept = '" & Trim(Cmb_Dept) & "' "
    End If
    If Trim(Cmb_Desig) <> "" Then
       g_Sql = g_Sql & " and c.c_desig = '" & Trim(Cmb_Desig) & "' "
    End If
    If Trim(Cmb_EmpType) <> "" Then
       g_Sql = g_Sql & " and c.c_emptype = '" & Trim(Cmb_EmpType) & "' "
    End If
    If Trim(Txtc_EmployeeName) <> "" Then
       g_Sql = g_Sql & " and c.c_empno = '" & Trim(Right(Trim(Txtc_EmployeeName), 7)) & "' "
    End If
    
    g_Sql = g_Sql & " order by a.c_empno, a.d_date "
    
    rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If rsChk.RecordCount <= 0 Then
       Exit Sub
    End If
    
    rsChk.MoveFirst
    
    Do While Not rsChk.EOF
       tmpEmpNo = Trim(rsChk("c_empno").Value)
       
       Do While tmpEmpNo = Trim(rsChk("c_empno").Value)
          NoClock = Is_Null(rsChk("n_noclock").Value, True)
          If NoClock > 0 Then
            vShStartHrs = Is_Null(rsChk("starthrs").Value, True)
            vShEndHrs = Is_Null(rsChk("endhrs").Value, True)
            vShMaxHrs = Is_Null(rsChk("maxhrs").Value, True)
            
            For i = 0 To 12
                aCurDayTime(i) = 0
                If i > 0 Then
                   aCurDayTime(i) = Is_Null(rsChk("n_time" & Trim(Str(i))).Value, True)
                End If
            Next i
            For i = 1 To 12
                If aCurDayTime(i) = 0 Then
                   Ctr = i - 1
                   Exit For
                End If
            Next i
            
            ' To Cross check
'            If NoClock <> Ctr Then
'               Screen.MousePointer = vbDefault
'               MsgBox "Critical Error in Clocking for the Emp No.  - " & RsChk("c_empno").Value, vbInformation, "Information"
'            End If
            
            rsChk.MoveNext
            If rsChk.EOF Then
               Exit Do
            End If
            If tmpEmpNo = Trim(rsChk("c_empno").Value) And Is_Null(rsChk("n_noclock").Value, True) > 0 Then
               tmpDate = rsChk("d_date").Value
               vOutTime = TimeToMins(rsChk("n_time1").Value)
               If (vOutTime + 1440) <= vShMaxHrs Then
                  For i = 0 To 12
                      aTime(i) = 0
                      aNextDayTime(i) = 0
                      If i > 0 Then
                         aNextDayTime(i) = Is_Null(rsChk("n_time" & Trim(Str(i))).Value, True)
                      End If
                  Next i
                  
                  Ctr = 0
                  For i = 1 To 12
                     vOutTime = TimeToMins(rsChk("n_time" & Trim(Str(i))).Value)
                     If (vOutTime + 1440) <= vShMaxHrs Then
                        Ctr = Ctr + 1
                        aTime(i) = Is_Null(rsChk("n_time" & Trim(Str(i))).Value, True)
                     Else
                        Exit For
                     End If
                  Next i
                  
                  For i = (Ctr + 1) To 12
                      aNextDayTime(i - Ctr) = aNextDayTime(i)
                  Next i
                  For i = (12 - Ctr) + 1 To 12
                      aNextDayTime(i) = 0
                  Next i
                  
                  Ctr = 0
                  For i = 1 To 12
                      If aNextDayTime(i) <> 0 Then
                         Ctr = Ctr + 1
                      End If
                  Next i
                  
                  g_Sql = "Update pr_clock_emp " & _
                          "Set n_noclock = " & Ctr & ", " & _
                          "n_time1 = " & aNextDayTime(1) & ", n_time2 = " & aNextDayTime(2) & ", " & _
                          "n_time3 = " & aNextDayTime(3) & ", n_time4 = " & aNextDayTime(4) & ", " & _
                          "n_time5 = " & aNextDayTime(5) & ", n_time6 = " & aNextDayTime(6) & ", " & _
                          "n_time7 = " & aNextDayTime(7) & ", n_time8 = " & aNextDayTime(8) & ", " & _
                          "n_time9 = " & aNextDayTime(8) & ", n_time10 = " & aNextDayTime(10) & ", " & _
                          "n_time11 = " & aNextDayTime(11) & ", n_time12 = " & aNextDayTime(12) & " " & _
                          "where c_empno = '" & tmpEmpNo & "' and d_date = '" & Is_Date(tmpDate, "S") & "' "
                  CON.Execute g_Sql
                  
                  ' Previouse Day Working
                  rsChk.MovePrevious
                  tmpDate = rsChk("d_date").Value
                  Ctr = 0
                  
                  For i = 1 To 12
                      If aCurDayTime(i) <> 0 Then
                         Ctr = Ctr + 1
                      End If
                  Next i
                  For i = Ctr + 1 To 12
                      If aTime(i - Ctr) > 0 Then
                         aCurDayTime(i) = aTime(i - Ctr) + 24
                      Else
                         aCurDayTime(i) = aTime(i - Ctr)
                      End If
                  Next i
                  
                  Ctr = 0
                  For i = 1 To 12
                      If aCurDayTime(i) <> 0 Then
                         Ctr = Ctr + 1
                      End If
                  Next i
                  
                  g_Sql = "Update pr_clock_emp " & _
                          "Set n_noclock = " & Ctr & ", " & _
                          "n_time1 = " & aCurDayTime(1) & ", n_time2 = " & aCurDayTime(2) & ", " & _
                          "n_time3 = " & aCurDayTime(3) & ", n_time4 = " & aCurDayTime(4) & ", " & _
                          "n_time5 = " & aCurDayTime(5) & ", n_time6 = " & aCurDayTime(6) & ", " & _
                          "n_time7 = " & aCurDayTime(7) & ", n_time8 = " & aCurDayTime(8) & ", " & _
                          "n_time9 = " & aCurDayTime(8) & ", n_time10 = " & aCurDayTime(10) & ", " & _
                          "n_time11 = " & aCurDayTime(11) & ", n_time12 = " & aCurDayTime(12) & " " & _
                          "where c_empno = '" & tmpEmpNo & "' and d_date = '" & Is_Date(tmpDate, "S") & "' "
                  CON.Execute g_Sql
                  
                  rsChk.MoveNext
                  
               End If
            End If
            rsChk.MovePrevious
          End If
          
          rsChk.MoveNext
          If rsChk.EOF Then
             Exit Do
          End If
       Loop
    Loop
    
 Exit Sub
 
Err_Chk:
    MsgBox "Error While Process - " & Err.Description & " - " & rsChk("c_empno").Value
  
End Sub

Private Sub Update_WorkHrs_Details(ByVal vDate As Date)
  Dim rsChk As New ADODB.Recordset, rsClock As New ADODB.Recordset
  Dim i As Long, j As Long, vCanteen As Integer, vTravel As Integer, vWkDay As Integer
  Dim vChk As String, vShift As String, vDayWork As String, vDept As String, vStaffType As String, vOtInStr As String, vOtOutStr As String
  Dim vArrTime, vDepTime, vLateHrs, vEarlyHrs, vWorkHrs, vOverTime, vPermHrs As Double
  Dim vShStartHrs, vShEndHrs, vShWorkHrs, vShLateHrs, vShBreakHrs As Double
  
  Dim vShBr1, vShBrMin1, vShBr2, vShBrMin2, vShBr3, vShBrMin3 As Double
  Dim vNoClock, vShClMin, vShMaxHrs, vShPermHrs, vShCutOffHrs As Double
  Dim vOtIn As Double, vOtOut As Double, vOt15 As Double, vOt20 As Double, vOt30 As Double
  
  Dim vPerArrTime, vPerDepTime, vActPerArrTime, vActPerDepTime As Double
  Dim vAddBrTime As Double
  Dim vActArrTime, vActDepTime As Double
  Dim vWorkHrs_FP As Double, vWorkHrs_Diff As Double
  Dim vPerAfterShift As Boolean, vPubHoliday As Boolean
  Dim tmpVal As Double
  
    vChk = "": vShift = "": vDayWork = "": vDept = "": vStaffType = "": vOtInStr = "": vOtOutStr = ""
    vArrTime = 0: vDepTime = 0: vLateHrs = 0: vEarlyHrs = 0: vWorkHrs = 0: vOverTime = 0: vPermHrs = 0
    
    vShStartHrs = 0: vShEndHrs = 0: vShWorkHrs = 0: vShLateHrs = 0: vShBreakHrs = 0
    vShBr1 = 0: vShBrMin1 = 0: vShBr2 = 0: vShBrMin2 = 0: vShBr3 = 0: vShBrMin3 = 0
    vNoClock = 0: vShClMin = 0: vShPermHrs = 0: vShMaxHrs = 0: vShCutOffHrs = 0
    
    vPerArrTime = 0: vPerDepTime = 0: vActPerArrTime = 0: vActPerDepTime = 0
    vAddBrTime = 0: vCanteen = 0: vTravel = 0
    
    vActArrTime = 0: vActDepTime = 0:    vWorkHrs_FP = 0: vWorkHrs_Diff = 0
    vOtIn = 0: vOtOut = 0: vOt15 = 0: vOt20 = 0: vOt30 = 0
    tmpVal = 0
    
    vPerAfterShift = False: vPubHoliday = False
    
    Set rsChk = Nothing
    g_Sql = "select d_phdate from pr_holiday_mst where c_rec_sta = 'A' and " & _
            "d_phdate = '" & Is_Date(vDate, "S") & "'"
    rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If rsChk.RecordCount > 0 Then
       vPubHoliday = True
    End If
    
    CON.Execute "SET DATEFIRST 1"
    
    Set rsChk = Nothing
    g_Sql = "Select a.c_empno, a.d_date, a.n_wkday, a.n_noclock, a.n_permhrs, a.c_shift, " & _
            "a.n_time1, a.n_time2, a.n_time3, a.n_time4, a.n_time5, a.n_time6, " & _
            "a.n_time7, a.n_time8, a.n_time9, a.n_time10, a.n_time11, a.n_time12, " & _
            "b.starthrs, b.endhrs, b.shifthrs, b.breakhrs, b.latehrs, b.permhrs, " & _
            "b.break1, b.mins1, b.break2, b.mins2, b.break3, b.mins3, b.clmin, " & _
            "b.maxhrs, b.cutoffhrs, a.n_canteen, a.c_flag, a.n_otin_str, a.n_otout_str, " & _
            "c.c_dept, c.c_daywork, c.c_stafftype, c.c_mealallow  " & _
            "from pr_clock_emp a, pr_clock_shift b, pr_emp_mst c " & _
            "where a.c_shift = b.c_shiftcode and a.c_empno = c.c_empno and c.c_rec_sta = 'A' and " & _
            "a.c_flag = 'B' and a.d_date = '" & Is_Date(vDate, "S") & "' "
  
    If Trim(Cmb_Branch) <> "" Then
       g_Sql = g_Sql & " and c.c_branch = '" & Trim(Right(Trim(Cmb_Branch), 7)) & "' "
    End If
    If Trim(Cmb_Dept) <> "" Then
       g_Sql = g_Sql & " and c.c_dept = '" & Trim(Right(Trim(Cmb_Dept), 7)) & "' "
    End If
    If Trim(Cmb_Desig) <> "" Then
       g_Sql = g_Sql & " and c.c_desig = '" & Trim(Cmb_Desig) & "' "
    End If
    If Trim(Cmb_EmpType) <> "" Then
       g_Sql = g_Sql & " and c.c_emptype = '" & Trim(Cmb_EmpType) & "' "
    End If
    If Trim(Txtc_EmployeeName) <> "" Then
       g_Sql = g_Sql & " and c.c_empno = '" & Trim(Right(Trim(Txtc_EmployeeName), 6)) & "' "
    End If
            
    g_Sql = g_Sql & " order by a.c_empno, a.d_date "
    rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    
    For i = 1 To rsChk.RecordCount
        vChk = "": vShift = "": vDayWork = "": vDept = "": vStaffType = "": vOtInStr = "": vOtOutStr = ""
        vArrTime = 0: vDepTime = 0: vLateHrs = 0: vEarlyHrs = 0: vWorkHrs = 0: vOverTime = 0: vPermHrs = 0
        
        vShStartHrs = 0: vShEndHrs = 0: vShWorkHrs = 0: vShLateHrs = 0: vShBreakHrs = 0
        vShBr1 = 0: vShBrMin1 = 0: vShBr2 = 0: vShBrMin2 = 0: vShBr3 = 0: vShBrMin3 = 0
        vNoClock = 0: vShClMin = 0: vShPermHrs = 0: vShMaxHrs = 0: vShCutOffHrs = 0
        
        vPerArrTime = 0: vPerDepTime = 0: vActPerArrTime = 0: vActPerDepTime = 0
        vAddBrTime = 0: vCanteen = 0: vTravel = 0
        
        vActArrTime = 0: vActDepTime = 0:    vWorkHrs_FP = 0: vWorkHrs_Diff = 0
        vOtIn = 0: vOtOut = 0: vOt15 = 0: vOt20 = 0: vOt30 = 0
        tmpVal = 0
       
        
        vDept = rsChk("c_dept").Value
        vStaffType = rsChk("c_stafftype").Value
        vDayWork = Is_Null(rsChk("c_daywork").Value, False)
        
        vWkDay = rsChk("n_wkday").Value
        vShift = rsChk("c_shift").Value
        vShStartHrs = rsChk("starthrs").Value
        vShEndHrs = rsChk("endhrs").Value
        vShWorkHrs = rsChk("shifthrs").Value
        vShLateHrs = rsChk("latehrs").Value
        vShBreakHrs = rsChk("breakhrs").Value
        
        vShBr1 = rsChk("break1").Value
        vShBrMin1 = rsChk("mins1").Value
        vShBr2 = rsChk("break2").Value
        vShBrMin2 = rsChk("mins2").Value
        vShBr3 = rsChk("break3").Value
        vShBrMin3 = rsChk("mins3").Value
        
        vNoClock = Is_Null(rsChk("n_noclock").Value, True)
        vShClMin = rsChk("clmin").Value
        vShPermHrs = rsChk("permhrs").Value
        vShMaxHrs = rsChk("maxhrs").Value
        vShCutOffHrs = rsChk("cutoffhrs").Value
        
        vOtInStr = Is_Null(rsChk("n_otin_str").Value, False)
        vOtOutStr = Is_Null(rsChk("n_otout_str").Value, False)
              
       ' Arr Time
        vArrTime = TimeToMins(Is_Null(rsChk("n_time1").Value, True))
        If vArrTime = 0 Then
           vActArrTime = 0
        ElseIf vArrTime < (vShStartHrs - vShLateHrs) Then
           vActArrTime = vArrTime
        ElseIf vArrTime >= (vShStartHrs - vShLateHrs) And vArrTime <= (vShStartHrs + vShLateHrs) Then
           vActArrTime = vShStartHrs
        ElseIf vArrTime > (vShStartHrs + vShLateHrs) And vArrTime <= (vShBr1 + vShBrMin1 + vShLateHrs) Then
           If vArrTime < (vShBr1 - vShLateHrs) Then
              vActArrTime = vArrTime
           ElseIf vArrTime >= (vShBr1 - vShLateHrs) Then
              vActArrTime = vShBr1 + vShBrMin1: vShBreakHrs = vShBreakHrs - vShBrMin1
           End If
        ElseIf vArrTime > (vShBr1 + vShBrMin1 + vShLateHrs) And vArrTime <= (vShBr2 + vShBrMin2 + vShLateHrs) Then
           If vArrTime < (vShBr2 - vShLateHrs) Then
              vActArrTime = vArrTime: vShBreakHrs = vShBreakHrs - vShBrMin1
           ElseIf vArrTime >= (vShBr2 - vShLateHrs) Then
              vActArrTime = vShBr2 + vShBrMin2: vShBreakHrs = vShBreakHrs - (vShBrMin1 + vShBrMin2)
           End If
        ElseIf vArrTime > (vShBr2 + vShBrMin2 + vShLateHrs) And vArrTime <= (vShBr3 + vShBrMin3 + vShLateHrs) Then
           If vArrTime < (vShBr3 - vShLateHrs) Then
              vActArrTime = vArrTime: vShBreakHrs = vShBreakHrs - (vShBrMin1 + vShBrMin2)
           ElseIf vArrTime >= (vShBr3 - vShLateHrs) Then
              vActArrTime = vShBr3 + vShBrMin3: vShBreakHrs = 0
           End If
        Else
           vActArrTime = vArrTime: vShBreakHrs = 0
        End If
          
        ' OT In time
        If vActArrTime > 0 And vActArrTime < vShStartHrs Then
           vOtIn = vShStartHrs - vActArrTime
        End If
       
        
        ' Dep Time
        If vNoClock > 1 Then 'And vNoClock Mod 2 = 0 Then
           vDepTime = TimeToMins(Is_Null(rsChk("n_time" & Trim(Str(rsChk("n_noclock").Value))).Value, True))
           If vNoClock Mod 2 = 1 Then
              vChk = "#"
           End If
        Else
           If vNoClock > 0 Then
              vChk = "*"   ' single clocking
           End If
        End If
        
        If vDepTime > 0 Then
           If vDepTime <= (vShStartHrs + vShLateHrs) Then
              vActDepTime = 0: vShBreakHrs = 0
           ElseIf vDepTime > (vShStartHrs + vShLateHrs) And vDepTime <= (vShBr1 + vShBrMin1 + vShLateHrs) Then
              If vDepTime < vShBr1 Then
                 vActDepTime = vDepTime: vShBreakHrs = 0
              ElseIf vDepTime >= vShBr1 Then
                 vActDepTime = vShBr1: vShBreakHrs = 0
              End If
           ElseIf vDepTime > (vShBr1 + vShBrMin1 + vShLateHrs) And vDepTime <= (vShBr2 + vShBrMin2 + vShLateHrs) Then
              If vDepTime < vShBr2 Then
                 vActDepTime = vDepTime: vShBreakHrs = vShBrMin1
              ElseIf vDepTime >= vShBr2 Then
                 vActDepTime = vShBr2: vShBreakHrs = vShBrMin1
              End If
           ElseIf vDepTime > (vShBr2 + vShBrMin2 + vShLateHrs) And vDepTime <= (vShBr3 + vShBrMin3 + vShLateHrs) Then
              If vDepTime < vShBr3 Then
                 vActDepTime = vDepTime: vShBreakHrs = vShBreakHrs - vShBrMin3
              ElseIf vDepTime >= vShBr3 Then
                 vActDepTime = vShBr3: vShBreakHrs = vShBreakHrs - vShBrMin3
              End If
           ElseIf vDepTime > (vShBr3 + vShBrMin3 + vShLateHrs) And vDepTime <= vShEndHrs + vShLateHrs Then
              If vDepTime < (vShEndHrs - vShLateHrs) Then
                 vActDepTime = vDepTime
              Else
                 vActDepTime = vShEndHrs
              End If
           Else
              vActDepTime = vDepTime
           End If
        End If
        
        ' OT Out time
        If vActDepTime > 0 And vActDepTime > vShEndHrs Then
           vOtOut = vActDepTime - vShEndHrs
        End If
        
        
        ' Late Hrs
        If vActArrTime > vShStartHrs And vActArrTime < vShEndHrs Then
           If vActArrTime < vShBr1 Then
              vLateHrs = vActArrTime - vShStartHrs
           ElseIf vActArrTime >= vShBr1 And vActArrTime < vShBr2 Then
              vLateHrs = (vActArrTime - vShBrMin1) - vShStartHrs
           ElseIf vActArrTime >= vShBr2 And vActArrTime < vShBr3 Then
              vLateHrs = (vActArrTime - (vShBrMin1 + vShBrMin2)) - vShStartHrs
           Else
              vLateHrs = (vActArrTime - (vShBrMin1 + vShBrMin2 + vShBrMin3)) - vShStartHrs
           End If
        End If
        If vWkDay = 6 Or vWkDay = 7 Or vPubHoliday Then
           vLateHrs = 0
        End If
        
      ' Early Hrs
        If vActDepTime > vShStartHrs And vActDepTime < vShEndHrs Then
           If vActDepTime <= vShBr1 Then
              vEarlyHrs = (vShEndHrs - vActDepTime) - (vShBrMin1 + vShBrMin2 + vShBrMin3)
           ElseIf vActDepTime > vShBr1 And vActDepTime <= vShBr2 Then
              vEarlyHrs = (vShEndHrs - vActDepTime) - (vShBrMin2 + vShBrMin3)
           ElseIf vActDepTime > vShBr2 And vActDepTime <= vShBr3 Then
              vEarlyHrs = (vShEndHrs - vActDepTime) - vShBrMin3
           Else
              vEarlyHrs = vShEndHrs - vActDepTime
           End If
        End If
        If vWkDay = 6 Or vWkDay = 7 Or vPubHoliday Then
           vEarlyHrs = 0
        End If
        
       
        If vArrTime = 0 Then vLateHrs = 0
        If vDepTime = 0 Then vEarlyHrs = 0
        If vActArrTime >= vActDepTime Then vLateHrs = 0
        If vActDepTime <= vActArrTime Then vEarlyHrs = 0
          
        
        ' Permission Hrs
        If vChk = "" And Is_Null(rsChk("n_noclock").Value, True) > 2 Then
           For j = 2 To vNoClock - 1 Step 2
               ' Permission Out
               vPerArrTime = TimeToMins(Is_Null(rsChk("n_time" & Trim(Str(j))).Value, True))
               If vPerArrTime < vShStartHrs Then
                  vActPerArrTime = vShStartHrs
               ElseIf vPerArrTime >= vShStartHrs And vPerArrTime < vShBr1 Then
                  vActPerArrTime = vPerArrTime
               ElseIf vPerArrTime >= vShBr1 And vPerArrTime < vShBr2 Then
                  If vPerArrTime <= vShBr1 + vShBrMin1 Then
                     vActPerArrTime = vShBr1 + vShBrMin1
                  Else
                     vActPerArrTime = vPerArrTime
                  End If
               ElseIf vPerArrTime >= vShBr2 And vPerArrTime < vShBr3 Then
                  If vPerArrTime <= vShBr2 + vShBrMin2 Then
                     vActPerArrTime = vShBr2 + vShBrMin2
                  Else
                     vActPerArrTime = vPerArrTime
                  End If
               Else
                  vActPerArrTime = vPerArrTime
               End If
           
               ' Permission In
               vPerDepTime = TimeToMins(Is_Null(rsChk("n_time" & Trim(Str(j + 1))).Value, True))
               If vPerDepTime < vShStartHrs Then
                  vActPerDepTime = vShStartHrs
               ElseIf vPerDepTime >= vShStartHrs And vPerDepTime < vShBr1 Then
                  vActPerDepTime = vPerDepTime
               ElseIf vPerDepTime >= vShBr1 And vPerDepTime < vShBr2 Then
                  If vPerDepTime <= vShBr1 + vShBrMin1 Then
                     vActPerDepTime = vShBr1
                  Else
                     vActPerDepTime = vPerDepTime - vShBrMin1
                  End If
               ElseIf vPerDepTime >= vShBr2 And vPerDepTime < vShBr3 Then
                  If vPerDepTime <= vShBr2 + vShBrMin2 Then
                     vActPerDepTime = vShBr2
                  Else
                     vActPerDepTime = vPerDepTime - vShBrMin2
                  End If
               Else
                  vActPerDepTime = vPerDepTime
               End If
               
               If (vActPerDepTime - vActPerArrTime) > vShPermHrs Then
                  If (vShEndHrs - vShLateHrs) < vActPerArrTime Then
                     vPerAfterShift = True
                  End If
                  vPermHrs = vPermHrs + (vActPerDepTime - vActPerArrTime)
               End If
           Next j
        End If
        If vWkDay = 6 Or vWkDay = 7 Or vPubHoliday Then
           vPermHrs = 0
        End If
        vPermHrs = 0
        
        ' Work hrs from Finger Print  - Actual
        If vChk = "" Or vChk = "#" Then
           vWorkHrs_FP = (vActDepTime - vActArrTime) - (vShBreakHrs + vPermHrs)
        End If
        If vWorkHrs_FP < 0 Then
           vWorkHrs_FP = 0
        ElseIf vWorkHrs_FP > 1800 Then ' 30 hrs
           vWorkHrs_FP = 0
        End If
       
        vWorkHrs = vWorkHrs_FP
        
             
        ' OT break-up
        If vWkDay = 7 Or vPubHoliday Then   ' sun & ph
           vOtIn = vWorkHrs
           vOtOut = 0
           If vOtIn > 0 And vOtInStr = "$" Then
              vOt20 = FloorTo15Mins(vOtIn)
              If vOt20 > 480 Then
                 vOt30 = vOt20 - 480
                 vOt20 = 480
              End If
           End If
        
        ElseIf vWkDay = 6 Then  ' sat
          vOtIn = vWorkHrs
          vOtOut = 0
           If vOtIn > 0 And vOtInStr = "$" Then
              If vActArrTime >= 360 Then  '6.00
                 vOt15 = FloorTo15Mins(vOtIn)
              Else
                 vOt20 = FloorTo15Mins(360 - vActArrTime)
                 vOt15 = FloorTo15Mins(vOtIn) - vOt20
              End If
              
              If vActDepTime <= 1380 Then  '23.00
                 vOt15 = vOt15 + FloorTo15Mins(vOtOut)
              Else
                 If vOt20 > 0 Then
                    tmpVal = vOt20
                 End If
                 vOt20 = vOt20 + FloorTo15Mins(vActDepTime - 1380)
                 vOt15 = vOt15 + (FloorTo15Mins(vOtOut) - (vOt20 - tmpVal))
              End If
           End If
           
        Else
           If vOtIn > 0 And vOtInStr = "$" Then
              If vActArrTime >= 360 Then  '6.00
                 vOt15 = FloorTo15Mins(vOtIn)
              Else
                 vOt20 = FloorTo15Mins(360 - vActArrTime)
                 vOt15 = FloorTo15Mins(vOtIn) - vOt20
              End If
           End If
            
           If vOtOut > 0 And vOtOutStr = "$" Then
              If vActDepTime <= 1380 Then  '23.00
                 vOt15 = vOt15 + FloorTo15Mins(vOtOut)
              Else
                 If vOt20 > 0 Then
                    tmpVal = vOt20
                 End If
                 vOt20 = vOt20 + FloorTo15Mins(vActDepTime - 1380)
                 vOt15 = vOt15 + (FloorTo15Mins(vOtOut) - (vOt20 - tmpVal))
              End If
           End If
        End If
          
        ' OT hrs
        If vWkDay = 6 Or vWkDay = 7 Or vPubHoliday Then
           vOverTime = vWorkHrs
        Else
           vOverTime = vOtIn + vOtOut
        End If

        If vWorkHrs <= 0 Then vOverTime = 0
        If vOverTime <= 0 Then vOverTime = 0
        
        If vOverTime = 0 Then
           vOtIn = 0: vOtOut = 0:  vOt15 = 0:  vOt20 = 0:  vOt30 = 0
        End If
        
        vWorkHrs_Diff = Abs(vWorkHrs_FP - vWorkHrs)
        
        'Convert Mins to Time
        vArrTime = MinsToTime(vArrTime)
        vLateHrs = MinsToTime(vLateHrs)
        vDepTime = MinsToTime(vDepTime)
        vEarlyHrs = MinsToTime(vEarlyHrs)
        vPermHrs = MinsToTime(vPermHrs)
        vWorkHrs = MinsToTime(vWorkHrs)
        vOverTime = MinsToTime(vOverTime)
        vWorkHrs_FP = MinsToTime(vWorkHrs_FP)
        vOtIn = MinsToTime(vOtIn)
        vOtOut = MinsToTime(vOtOut)
        vOt15 = MinsToTime(vOt15)
        vOt20 = MinsToTime(vOt20)
        vOt30 = MinsToTime(vOt30)
        
        g_Sql = "update pr_clock_emp set " & _
                "n_period = " & vPayPeriod & ", " & _
                "n_arrtime = " & vArrTime & ", n_arrtime_dec = " & MinsToDecimal(vArrTime) & ", n_arrtime_min = " & TimeToMins(vArrTime) & ", " & _
                "n_latehrs = " & vLateHrs & ", n_latehrs_dec = " & MinsToDecimal(vLateHrs) & ", n_latehrs_min = " & TimeToMins(vLateHrs) & ", " & _
                "n_deptime = " & vDepTime & ", n_deptime_dec = " & MinsToDecimal(vDepTime) & ", n_deptime_min = " & TimeToMins(vDepTime) & ", " & _
                "n_earlhrs = " & vEarlyHrs & ", n_earlhrs_dec = " & MinsToDecimal(vEarlyHrs) & ", n_earlhrs_min = " & TimeToMins(vEarlyHrs) & ", " & _
                "n_permhrs = " & vPermHrs & ", n_permhrs_dec = " & MinsToDecimal(vPermHrs) & ", n_permhrs_min = " & TimeToMins(vPermHrs) & ", " & _
                "n_workhrs = " & vWorkHrs & ", n_workhrs_dec = " & MinsToDecimal(vWorkHrs) & ", n_workhrs_min = " & TimeToMins(vWorkHrs) & ", " & _
                "n_overtime = " & vOverTime & ", n_overtime_dec = " & MinsToDecimal(vOverTime) & ", n_overtime_min = " & TimeToMins(vOverTime) & ", " & _
                "n_otin = " & vOtIn & ", n_otin_dec = " & MinsToDecimal(vOtIn) & ", n_otin_min = " & TimeToMins(vOtIn) & ", " & _
                "n_otout = " & vOtOut & ", n_otout_dec = " & MinsToDecimal(vOtOut) & ", n_otout_min = " & TimeToMins(vOtOut) & ", " & _
                "n_ot15 = " & vOt15 & ", n_ot15_dec = " & MinsToDecimal(vOt15) & ", n_ot15_min = " & TimeToMins(vOt15) & ", " & _
                "n_ot20 = " & vOt20 & ", n_ot20_dec = " & MinsToDecimal(vOt20) & ", n_ot20_min = " & TimeToMins(vOt20) & ", " & _
                "n_ot30 = " & vOt30 & ", n_ot30_dec = " & MinsToDecimal(vOt30) & ", n_ot30_min = " & TimeToMins(vOt30) & ", " & _
                "n_canteen = " & vCanteen & ", c_chq = '" & Trim(vChk) & "', " & _
                "n_workhrs_fp = " & vWorkHrs_FP & ", n_workhrs_diff = " & vWorkHrs_Diff & " " & _
                "where c_empno = '" & Is_Null(rsChk("c_empno").Value, False) & "' and " & _
                "d_date = '" & Is_Date(rsChk("d_date").Value, "S") & "' and " & _
                "c_flag = 'B'"

        CON.Execute g_Sql
        
' this process is slow compare with update statement.
'        Set rsClock = Nothing
'        g_Sql = "select * from pr_clock_emp where c_empno = '" & Is_Null(rsChk("c_empno").Value, False) & "' and " & _
'                "d_date = '" & Is_Date(rsChk("d_date").Value, "S") & "' and c_flag = 'B'"
'        rsClock.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
'        If rsClock.RecordCount > 0 Then
'           rsClock("n_period").Value = vPayPeriod
'           rsClock("n_arrtime").Value = vArrTime
'           rsClock("n_latehrs").Value = vLateHrs
'           rsClock("n_deptime").Value = vDepTime
'           rsClock("n_earlhrs").Value = vEarlyHrs
'           rsClock("n_permhrs").Value = vPermHrs
'           rsClock("n_workhrs").Value = vWorkHrs
'           rsClock("n_overtime").Value = vOverTime
'           rsClock("n_otin").Value = vOtIn
'           rsClock("n_otout").Value = vOtOut
'           rsClock("n_ot15").Value = vOt15
'           rsClock("n_ot20").Value = vOt20
'           rsClock("n_ot30").Value = vOt30
'
'           rsClock("n_arrtime_dec").Value = MinsToDecimal(vArrTime)
'           rsClock("n_latehrs_dec").Value = MinsToDecimal(vLateHrs)
'           rsClock("n_deptime_dec").Value = MinsToDecimal(vDepTime)
'           rsClock("n_earlhrs_dec").Value = MinsToDecimal(vEarlyHrs)
'           rsClock("n_permhrs_dec").Value = MinsToDecimal(vPermHrs)
'           rsClock("n_workhrs_dec").Value = MinsToDecimal(vWorkHrs)
'           rsClock("n_overtime_dec").Value = MinsToDecimal(vOverTime)
'           rsClock("n_otin_dec").Value = MinsToDecimal(vOtIn)
'           rsClock("n_otout_dec").Value = MinsToDecimal(vOtOut)
'           rsClock("n_ot15_dec").Value = MinsToDecimal(vOt15)
'           rsClock("n_ot20_dec").Value = MinsToDecimal(vOt20)
'           rsClock("n_ot30_dec").Value = MinsToDecimal(vOt30)
'
'           rsClock("n_arrtime_min").Value = TimeToMins(vArrTime)
'           rsClock("n_latehrs_min").Value = TimeToMins(vLateHrs)
'           rsClock("n_deptime_min").Value = TimeToMins(vDepTime)
'           rsClock("n_earlhrs_min").Value = TimeToMins(vEarlyHrs)
'           rsClock("n_permhrs_min").Value = TimeToMins(vPermHrs)
'           rsClock("n_workhrs_min").Value = TimeToMins(vWorkHrs)
'           rsClock("n_overtime_min").Value = TimeToMins(vOverTime)
'           rsClock("n_otin_min").Value = TimeToMins(vOtIn)
'           rsClock("n_otout_min").Value = TimeToMins(vOtOut)
'           rsClock("n_ot15_min").Value = TimeToMins(vOt15)
'           rsClock("n_ot20_min").Value = TimeToMins(vOt20)
'           rsClock("n_ot30_min").Value = TimeToMins(vOt30)
'
'           rsClock("n_canteen").Value = vCanteen
'           rsClock("c_chq").Value = Trim(vChk)
'           rsClock("n_workhrs_fp").Value = vWorkHrs_FP
'           rsClock("n_workhrs_diff").Value = vWorkHrs_Diff
'
'           rsClock.Update
'        End If
                
        rsChk.MoveNext
    Next i
            
End Sub

Private Sub Combo_Load()
 
    Call LoadComboCompany(Me)
    Call LoadComboBranch(Me)
    Call LoadComboDept(Me)
    Call LoadComboDesig(Me)
    Call LoadComboEmpType(Me)

End Sub

Private Sub Txtc_EmployeeName_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim Search As New Search.MyClass, SerVar
 
   If KeyCode = vbKeyDelete Then
      Txtc_EmployeeName = ""
   End If
   
   If KeyCode = vbKeyF2 Then
      Search.Query = "select c_empno EmpNo, c_name Name, c_othername OtherName, c_dept Dept, " & _
                     "c_designame Desig, c_branch Branch, c_emptype Type " & _
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
       If mnuOption = "DWH" Then
          If (Val(Txtc_Month) <= 0 Or Val(Txtc_Month) > 12) Then
             MsgBox "Not a valid month", vbInformation, "Information"
             Txtc_Month.SetFocus
             Cancel = True
             Exit Sub
          End If
       Else
          If (Val(Txtc_Month) <= 0 Or Val(Txtc_Month) > 13) Then
             MsgBox "Not a valid month", vbInformation, "Information"
             Txtc_Month.SetFocus
             Cancel = True
             Exit Sub
          End If
       End If
    End If
    
    If Trim(Txtc_Month) <> "" And Trim(Txtc_Year) <> "" Then
       vPayPeriod = Is_Null(Format(Trim(Txtc_Year), "0000") & Trim(Txtc_Month), True)
       If Not ChkPeriodOpen(vPayPeriod, "W") Then
          Txtc_Month.SetFocus
          Cancel = True
          Exit Sub
       End If
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
     vPayPeriod = Format(Trim(Txtc_Year), "0000") & Format(Trim(Txtc_Month), "00")
     If Not ChkPeriodOpen(Is_Null(vPayPeriod, True), "W") Then
        Txtc_Year.SetFocus
        Cancel = True
     End If
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
    
    If Val(Txtc_Month) = 1 Then
       vPrevPayPeriod = Is_Null(Format(Str(Val(Txtc_Year) - 1), "0000") & "12", True)
    Else
       vPrevPayPeriod = vPayPeriod - 1
    End If
End Sub

Private Function Gen_Check() As Boolean
    
    If Trim(Txtc_Month) = "" Then
       MsgBox "Month should not be empty", vbInformation, "Information"
       Txtc_Month.SetFocus
       Exit Function
    ElseIf Trim(Txtc_Year) = "" Then
       MsgBox "Year should not be empty", vbInformation, "Information"
       Txtc_Year.SetFocus
       Exit Function
    ElseIf Not IsDate(Txtd_FromDate.Text) Then
       MsgBox "Date Should not be Empty", vbInformation, "Information"
       Txtd_FromDate.Text.SetFocus
       Exit Function
    ElseIf Not IsDate(Txtd_ToDate.Text) Then
       MsgBox "Date Should not be Empty", vbInformation, "Information"
       Txtd_ToDate.SetFocus
       Exit Function
    End If
    
    If CDate(Txtd_FromDate.Text) < vPayPeriodFrom Or CDate(Txtd_FromDate.Text) > vPayPeriodTo Then
       MsgBox "From Date is out side the Pay Period. Please check the date", vbInformation, "Information"
       Txtd_FromDate.SetFocus
       Exit Function
    End If
    If CDate(Txtd_ToDate.Text) < vPayPeriodFrom Or CDate(Txtd_ToDate.Text) > vPayPeriodTo Then
       MsgBox "To Date is out side the Pay Period. Please check the date", vbInformation, "Information"
       Txtd_ToDate.SetFocus
       Exit Function
    End If
     
    vPayPeriod = Is_Null(Format(Trim(Txtc_Year), "0000") & Format(Trim(Txtc_Month), "00"), True)
    If Not ChkPeriodOpen(vPayPeriod, "W") Then
       Txtc_Month.SetFocus
       Exit Function
    End If
    
    Gen_Check = True

End Function

Private Sub WeeklyHours_Payroll_Process()
On Error GoTo Err_Upd
  Dim rsChk As New ADODB.Recordset
  Dim rsAttend As New ADODB.Recordset
  Dim rsWrkDtl As New ADODB.Recordset
 
  Dim tmpEmpNo As String, vSalaryType As String, vEmpType As String
  Dim i, j, k As Long
  Dim vWrkHrs, vWrkHrs_Dec, vOThrs, vLateHrs, vSunHrs, vPHHrs As Double
  Dim vSatNight, vSunNight, vMonMorning As Double
  Dim vTotBasicWrkHrs, vOt1, vOt2, vOt3, vSP1, vSP2, vTotLateHrs As Double
  Dim vLeaveHrs, vAbsentHrs, vSunPhHrs, vPresDays, vLopDays, vLopHrs As Double
  Dim vSunPh01, vSunPh02 As Double
  Dim vPublicHoliday, vLocalLeave, vSickLeave, vInjuryLeave, vProlongLeave, vWeddingLeave, vMatLeave, vPatLeave, vCompLeave, vOthLeave As Double
  Dim vWorkOffPay As Double, vPhOneYearEligible As Double
  Dim vNoWeek, vSlFullDay, vNoTravAllow, vNoMealAllow, vNoNightAllow, vNoMealAllowPrev As Double
  Dim vProcSql As String, vPresAbs As String
           
      If MsgBox("Confim to Process", vbYesNo, "Information") = vbNo Then
         Exit Sub
      End If
           
      vProcSql = "a.n_period = " & Trim(vPayPeriod)
      If Trim(Cmb_Company) <> "" Then
         vProcSql = Trim(vProcSql) & " and b.c_company = '" & Trim(Right(Trim(Cmb_Company), 7)) & "' "
      End If
      If Trim(Cmb_Branch) <> "" Then
         vProcSql = Trim(vProcSql) & " and b.c_branch = '" & Trim(Cmb_Branch) & "' "
      End If
      If Trim(Cmb_Dept) <> "" Then
         vProcSql = Trim(vProcSql) & " and b.c_dept = '" & Trim(Cmb_Dept) & "' "
      End If
      If Trim(Cmb_Desig) <> "" Then
         g_Sql = g_Sql & " and c.c_desig = '" & Trim(Cmb_Desig) & "' "
      End If
      If Trim(Cmb_EmpType) <> "" Then
         g_Sql = g_Sql & " and b.c_emptype = '" & Trim(Cmb_EmpType) & "' "
      End If
      If Trim(Txtc_EmployeeName) <> "" Then
         vProcSql = Trim(vProcSql) & " and b.c_empno = '" & Trim(Right(Trim(Txtc_EmployeeName), 7)) & "' "
      End If

      ProBar.Value = 5
      DoEvents
      
      ' Calculate Public holiday Work hrs
      Call SunPH_Work_Hours_Process(vProcSql)
      ProBar.Value = 10
      DoEvents
       
      
      ' Working hrs - Weekly calculated. Detail File.
      CON.Execute "SET DATEFIRST 1"
      
      Set rsAttend = Nothing
      g_Sql = "select a.c_empno, a.d_date, datepart(dw, a.d_date) n_wkday, a.c_presabs, " & _
              "a.n_deptime_dec, a.n_latehrs_dec, a.n_earlhrs_dec, a.n_permhrs_dec, a.n_workhrs_dec, a.n_sunphhrs_dec,  " & _
              "a.c_shift, a.n_present, c.n_shifthrs, c.n_maafter, c.n_naafter, " & _
              "b.c_emptype, b.c_stafftype, b.c_salarytype, b.c_daywork, b.d_doj, b.d_dol " & _
              "from pr_clock_emp a left outer join pr_clock_shift c on a.c_shift = c.c_shiftcode, pr_emp_mst b " & _
              "where a.c_empno = b.c_empno and b.c_rec_sta = 'A' and IsNull(b.c_clockcard,'1')<>'0' and " & Trim(vProcSql) & " " & _
              "order by a.c_empno, a.d_date "
      rsAttend.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
      If rsAttend.RecordCount > 0 Then
         rsAttend.MoveFirst
         
         Do While Not rsAttend.EOF()
            CON.BeginTrans
            
            Set rsWrkDtl = Nothing
            g_Sql = "select * from pr_workhrs_dtl " & _
                    "where n_period = " & vPayPeriod & " and c_empno = '" & rsAttend("c_empno").Value & "'"
            rsWrkDtl.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
            If rsWrkDtl.RecordCount = 0 Then
               rsWrkDtl.AddNew
            End If
               
            rsWrkDtl("n_period").Value = Is_Null(vPayPeriod, False)
            rsWrkDtl("c_empno").Value = Is_Null(rsAttend("c_empno").Value, False)
            vEmpType = Is_Null(rsAttend("c_emptype").Value, False)
            vSalaryType = Is_Null(rsAttend("c_salarytype").Value, False)
            
            For k = 1 To 6
                rsWrkDtl("n_wkwrkhrs" + Trim(Str(k))).Value = 0
                rsWrkDtl("n_wklatehrs" + Trim(Str(k))).Value = 0
                rsWrkDtl("n_wkothrs" + Trim(Str(k))).Value = 0
                rsWrkDtl("n_wksunhrs" + Trim(Str(k))).Value = 0
                rsWrkDtl("n_wkphhrs" + Trim(Str(k))).Value = 0
            Next k
           
            tmpEmpNo = Trim(rsAttend("c_empno").Value)
            i = 0: j = 0
            vWrkHrs = 0: vWrkHrs_Dec = 0: vOThrs = 0: vLateHrs = 0:  vSunHrs = 0: vPHHrs = 0
            vSatNight = 0: vSunNight = 0: vMonMorning = 0
            vNetNorHrs = 0: vNetOT1 = 0: vNetOT2 = 0: vNetOT3 = 0
            vSunPh01 = 0: vSunPh02 = 0: vSunPhHrs = 0: vPresDays = 0: vLopDays = 0: vLopHrs = 0
            vPublicHoliday = 0: vLocalLeave = 0: vSickLeave = 0: vInjuryLeave = 0: vProlongLeave = 0: vWeddingLeave = 0: vMatLeave = 0: vPatLeave = 0
            vCompLeave = 0: vOthLeave = 0:  vPhOneYearEligible = 0: vWorkOffPay = 0
            vAbsentHrs = 0: vNoWeek = 0: vSlFullDay = 0: vNoTravAllow = 0: vNoMealAllow = 0: vNoNightAllow = 0
            vNoMealAllowPrev = 0
            vPresAbs = ""
            
            Do While tmpEmpNo = Trim(rsAttend("c_empno").Value)
               If i < rsAttend("n_wkday").Value Then
                  i = rsAttend("n_wkday").Value
                  
                  vWrkHrs_Dec = Is_Null(rsAttend("n_workhrs_dec").Value, True)
                  vSunPhHrs = Is_Null(rsAttend("n_sunphhrs_dec").Value, True)
                  vPresAbs = Is_Null(rsAttend("c_presabs").Value, False)
                  vPresDays = Is_Null(rsAttend("n_present").Value, True)
                
                  ' Except Public holiday and Sunday
                  If vPresAbs <> "PH" And rsAttend("n_wkday").Value <> 7 Then
                     If rsAttend("c_salarytype").Value = "ML" And rsAttend("c_stafftype").Value = "O" Then
                        vLateHrs = vLateHrs + Is_Null(rsAttend("n_latehrs_dec").Value, True) + Is_Null(rsAttend("n_earlhrs_dec").Value, True) + Is_Null(rsAttend("n_permhrs_dec").Value, True)
                     End If
                     If rsAttend("n_wkday").Value = 6 Then 'sat
                        If Is_Null(rsAttend("n_deptime_dec").Value, True) > 24 Then
                           vSatNight = vSunPhHrs
                           vSunHrs = vSunHrs + vSatNight
                           vWrkHrs = vWrkHrs + (vWrkHrs_Dec - vSunPhHrs)
                        Else
                           vWrkHrs = vWrkHrs + vMonMorning + vWrkHrs_Dec
                           vMonMorning = 0
                        End If
                     Else
                        vWrkHrs = vWrkHrs + vMonMorning + vWrkHrs_Dec
                        vMonMorning = 0
                     End If
                  End If
                  
                  
                  ' Sunday
                  If rsAttend("n_wkday").Value = 7 Then
                     If Is_Null(rsAttend("n_deptime_dec").Value, True) > 24 Then
                        vSunNight = vSunPhHrs
                        vSunHrs = vSunHrs + vSunNight
                     Else
                        vSunHrs = vSunHrs + vWrkHrs_Dec
                     End If
                     If vSunHrs <= 8 Then
                        vSunPh01 = vSunPh01 + vSunHrs
                     Else
                        vSunPh01 = vSunPh01 + 8
                        vSunPh02 = vSunPh02 + (vSunHrs - 8)
                     End If
                  End If
                  
                
                  ' For Monday
                  If rsAttend("n_wkday").Value = 7 Then 'sun night
                     If Is_Null(rsAttend("n_deptime_dec").Value, True) > 24 Then
                        vMonMorning = (vWrkHrs_Dec - vSunPhHrs)
                     End If
                  End If
                  
                  
                  ' Public Holiday
                  If vPresAbs <> "PH" And rsAttend("n_wkday").Value <> 7 And rsAttend("n_wkday").Value <> 6 Then
                     If vSunPhHrs > 0 Then
                        vPHHrs = vPHHrs + vSunPhHrs
                        vWrkHrs = vWrkHrs - vSunPhHrs
                        If vSunPhHrs <= 8 Then
                           vSunPh01 = vSunPh01 + vSunPhHrs
                        Else
                           vSunPh01 = vSunPh01 + 8
                           vSunPh02 = vSunPh02 + (vSunPhHrs - 8)
                        End If
                     End If
                  End If
                  If vPresAbs = "PH" Then
                     ' Find out number of hrs worked on public holiday
                     vPHHrs = vPHHrs + vSunPhHrs
                     If Is_Null(rsAttend("n_deptime_dec").Value, True) > 24 Then
                        vWrkHrs = vWrkHrs + (vWrkHrs_Dec - vSunPhHrs)
                     End If
                       
                     ' To pay public holiday hrs
                     If Left(Is_Null(rsAttend("c_daywork").Value, False), 1) = "5" Then
                        If rsAttend("n_wkday").Value <> 6 Then
                           vWrkHrs = vWrkHrs + 9
                           If rsAttend("d_doj").Value + 365 <= rsAttend("d_date").Value Then
                              vPublicHoliday = vPublicHoliday + 9
                           Else
                              vPhOneYearEligible = vPhOneYearEligible + 9
                           End If
                        End If
                     Else
                        If rsAttend("n_wkday").Value <> 6 Then
                           vWrkHrs = vWrkHrs + 8.08
                           If rsAttend("d_doj").Value + 365 <= rsAttend("d_date").Value Then
                              vPublicHoliday = vPublicHoliday + 8.08
                           Else
                              vPhOneYearEligible = vPhOneYearEligible + 8.08
                           End If
                        Else
                           vWrkHrs = vWrkHrs + 4.58
                           If rsAttend("d_doj").Value + 365 <= rsAttend("d_date").Value Then
                              vPublicHoliday = vPublicHoliday + 4.58
                           Else
                              vPhOneYearEligible = vPhOneYearEligible + 4.58
                           End If
                        End If
                     End If
                       
                     If vSunPhHrs <= 8 Then
                        vSunPh01 = vSunPh01 + vSunPhHrs
                     Else
                        vSunPh01 = vSunPh01 + 8
                        vSunPh02 = vSunPh02 + (vSunPhHrs - 8)
                     End If
                  End If
                  
                  
                  ' Travel Allowance Refund
                  If vWrkHrs_Dec > 0 And Left(Is_Null(rsAttend("c_daywork").Value, False), 1) = "5" Then
                     vNoTravAllow = vNoTravAllow + 1
                     If Is_Null(rsAttend("n_deptime_dec").Value, True) <= 18 Then
                        vNoTravAllow = vNoTravAllow + 1
                     End If
                  End If
                  
                  
                  ' Meal Allowance & Night Allow
                  If Is_Null(rsAttend("n_deptime_dec").Value, True) > Is_Null(rsAttend("n_maafter").Value, True) Then
                     vNoMealAllow = vNoMealAllow + 1
                  End If

                  'Night Allowance
                  If Is_Null(rsAttend("n_deptime_dec").Value, True) > Is_Null(rsAttend("n_naafter").Value, True) Then
                     vNoNightAllow = vNoNightAllow + 1
                  End If
                  
                  ' Leave taken
                  ' WP - work off pay, incase of work 24 hrs
                  If vPresAbs = "CL" Or vPresAbs = "LC" Or vPresAbs = "SL" Or vPresAbs = "IL" Or vPresAbs = "PI" Or vPresAbs = "WL" Or vPresAbs = "ML" Or vPresAbs = "PL" Or vPresAbs = "CO" Or vPresAbs = "WP" Then
                     vLeaveHrs = 0
                     If Left(Is_Null(rsAttend("c_daywork").Value, False), 1) = "5" Then
                        If rsAttend("n_wkday").Value <> 6 Then
                           If vPresDays = 1 Or vPresDays = 0 Then
                              vLeaveHrs = 9
                           Else
                              vLeaveHrs = 4.5
                           End If
                        End If
                     Else
                        If rsAttend("n_wkday").Value <> 6 Then
                           If vPresDays = 1 Or vPresDays = 0 Then
                              vLeaveHrs = 8.08
                           Else
                              vLeaveHrs = 4.08
                           End If
                        Else
                           vLeaveHrs = 4.58
                        End If
                     End If
                     vWrkHrs = vWrkHrs + vLeaveHrs
                  End If
                  
                  If vPresAbs = "CL" Or vPresAbs = "LC" Then
                     vLocalLeave = vLocalLeave + vLeaveHrs
                  ElseIf vPresAbs = "SL" Then
                     vSickLeave = vSickLeave + vLeaveHrs
                     If vLeaveHrs > 0 And vWrkHrs_Dec = 0 Then
                        vSlFullDay = vSlFullDay + 1
                     End If
                  ElseIf vPresAbs = "IL" Then
                     vInjuryLeave = vInjuryLeave + vLeaveHrs
                  ElseIf vPresAbs = "PI" Then
                     vProlongLeave = vProlongLeave + vLeaveHrs
                  ElseIf vPresAbs = "WL" Then
                     vWeddingLeave = vWeddingLeave + vLeaveHrs
                  ElseIf vPresAbs = "ML" Then
                     vMatLeave = vMatLeave + vLeaveHrs
                  ElseIf vPresAbs = "PL" Then
                     vPatLeave = vPatLeave + vLeaveHrs
                  ElseIf vPresAbs = "CO" Then
                     vCompLeave = vCompLeave + vLeaveHrs
                  ElseIf vPresAbs = "WP" Then
                     vWorkOffPay = vWorkOffPay + vLeaveHrs
                  ElseIf vPresAbs = "XX" Then
                     vOthLeave = vOthLeave + vLeaveHrs
                  End If
                                  
                  
                  ' Absent. no present bonus
                  If rsAttend("d_doj").Value > CDate(vPayPeriodFrom) Then
                     vSlFullDay = 1
                  End If
                  If IsDate(rsAttend("d_dol").Value) And vSlFullDay < 1 Then
                     If Left(Trim(rsAttend("c_daywork").Value), 1) = 5 Then
                        If rsAttend("d_dol").Value < (CDate(vPayPeriodTo) - 2) Then
                           vSlFullDay = 1
                        End If
                     Else
                        If rsAttend("d_dol").Value < (CDate(vPayPeriodTo) - 1) Then
                           vSlFullDay = 1
                        End If
                     End If
                  End If
                  
                  ' Calculate this to find out LOP and Late hrs. Only for Staff
                  If Is_Null(rsAttend("c_salarytype").Value, False) = "ML" Then
                     If Is_Null(rsAttend("c_presabs").Value, False) = "A" Or Is_Null(rsAttend("c_presabs").Value, False) = "LP" Then
                        vLopDays = vLopDays + vPresDays
                        If Left(Is_Null(rsAttend("c_daywork").Value, False), 1) = "5" Then
                           If rsAttend("n_wkday").Value <> 6 Then
                              If vPresDays = 1 Or vPresDays = 0 Then
                                 vAbsentHrs = vAbsentHrs + 9
                              Else
                                 vAbsentHrs = vAbsentHrs + 4.5
                              End If
                           End If
                        Else
                           If rsAttend("n_wkday").Value <> 6 Then
                              If vPresDays = 1 Or vPresDays = 0 Then
                                 vAbsentHrs = vAbsentHrs + 8.08
                              Else
                                 vAbsentHrs = vAbsentHrs + 4.08
                              End If
                           Else
                              vAbsentHrs = vAbsentHrs + 4.58
                           End If
                        End If
                       ' lop for monthly paid
                       vLopHrs = vLopHrs + vAbsentHrs
                     End If
                  End If
               Else
                  j = j + 1

                  'Over Time
                  If vWrkHrs - 45 > 0 Then
                     rsWrkDtl("n_wkwrkhrs" + Trim(Str(j))).Value = 45 - (vWorkOffPay + vPhOneYearEligible)
                     rsWrkDtl("n_wkothrs" + Trim(Str(j))).Value = vWrkHrs - 45
                  Else
                     rsWrkDtl("n_wkwrkhrs" + Trim(Str(j))).Value = vWrkHrs - vPhOneYearEligible
                     rsWrkDtl("n_wkothrs" + Trim(Str(j))).Value = 0
                  End If
                  
                  ' Late Hrs
                  If vLateHrs > 0 Then
                     If vWrkHrs + vAbsentHrs >= 45 Then
                        rsWrkDtl("n_wklatehrs" + Trim(Str(j))).Value = 0
                     Else
                        If (45 - (vWrkHrs + vAbsentHrs)) < vLateHrs Then
                           rsWrkDtl("n_wklatehrs" + Trim(Str(j))).Value = 45 - (vWrkHrs + vAbsentHrs)
                        Else
                           rsWrkDtl("n_wklatehrs" + Trim(Str(j))).Value = vLateHrs
                        End If
                     End If
                  Else
                     rsWrkDtl("n_wklatehrs" + Trim(Str(j))).Value = 0
                  End If
                  
                  rsWrkDtl("n_wksunhrs" + Trim(Str(j))).Value = vSunHrs + vPHHrs
                  vWrkHrs = 0: vOThrs = 0: vLateHrs = 0: vSunHrs = 0: vPHHrs = 0
                  vSatNight = 0: vSunNight = 0: vAbsentHrs = 0: vWorkOffPay = 0: vPhOneYearEligible = 0
                  i = 0
                  rsAttend.MovePrevious
               End If
               rsAttend.MoveNext
               If rsAttend.EOF Then
                  Exit Do
               End If
            Loop
                  j = j + 1
                  vWrkHrs = vWrkHrs + vMonMorning

                  If vWrkHrs = 0 Then
                     rsWrkDtl("n_wkwrkhrs" + Trim(Str(j))).Value = 0.01
                  Else
                     rsWrkDtl("n_wkwrkhrs" + Trim(Str(j))).Value = vWrkHrs
                  End If
                  
                  'Over Time
                  If vWrkHrs - 45 > 0 Then
                     rsWrkDtl("n_wkwrkhrs" + Trim(Str(j))).Value = 45 - (vWorkOffPay + vPhOneYearEligible)
                     rsWrkDtl("n_wkothrs" + Trim(Str(j))).Value = vWrkHrs - 45
                  Else
                     If vWrkHrs = 0 Then
                        rsWrkDtl("n_wkwrkhrs" + Trim(Str(j))).Value = 0.01
                     Else
                        rsWrkDtl("n_wkwrkhrs" + Trim(Str(j))).Value = vWrkHrs - vPhOneYearEligible
                     End If
                     rsWrkDtl("n_wkothrs" + Trim(Str(j))).Value = 0
                  End If
                  
                  ' Late Hrs
                  If vLateHrs > 0 Then
                     If vWrkHrs + vAbsentHrs >= 45 Then
                        rsWrkDtl("n_wklatehrs" + Trim(Str(j))).Value = 0
                     Else
                        If (45 - (vWrkHrs + vAbsentHrs)) < vLateHrs Then
                           rsWrkDtl("n_wklatehrs" + Trim(Str(j))).Value = 45 - (vWrkHrs + vAbsentHrs)
                        Else
                           rsWrkDtl("n_wklatehrs" + Trim(Str(j))).Value = vLateHrs
                        End If
                     End If
                  Else
                     rsWrkDtl("n_wklatehrs" + Trim(Str(j))).Value = 0
                  End If
                  
                  rsWrkDtl("n_wksunhrs" + Trim(Str(j))).Value = vSunHrs + vPHHrs
                  vWrkHrs = 0: vOThrs = 0: vLateHrs = 0: vSunHrs = 0: vPHHrs = 0
                  vSatNight = 0: vAbsentHrs = 0: vWorkOffPay = 0
                  i = 0
            
            
            vTotBasicWrkHrs = 0: vOt1 = 0: vOt2 = 0: vOt3 = 0: vSP1 = 0: vSP2 = 0: vTotLateHrs = 0
            For k = 1 To 6
                If k = 1 And Weekday(CDate(vPayPeriodFrom)) <> vbMonday Then
                   Call CalLastMonthOtPay(Is_Null(rsWrkDtl("c_empno").Value, False), Is_Null(rsWrkDtl("n_wkwrkhrs1").Value, True))
                   rsWrkDtl("n_wkwrkhrs1").Value = vNetNorHrs
                   rsWrkDtl("n_wkothrs1").Value = vNetOT1 + vNetOT2 + vNetOT3
                   vOt1 = vNetOT1
                   vOt2 = vNetOT2
                   vOt3 = vNetOT3
                Else
                   If Is_Null(rsWrkDtl("n_wkothrs" + Trim(Str(k))).Value, True) > 0 And Is_Null(rsWrkDtl("n_wkothrs" + Trim(Str(k))).Value, True) <= 10 Then
                      vOt1 = vOt1 + Is_Null(rsWrkDtl("n_wkothrs" + Trim(Str(k))).Value, True)
                   ElseIf Is_Null(rsWrkDtl("n_wkothrs" + Trim(Str(k))).Value, True) > 10 And Is_Null(rsWrkDtl("n_wkothrs" + Trim(Str(k))).Value, True) <= 15 Then
                      vOt1 = vOt1 + 10
                      vOt2 = vOt2 + (Is_Null(rsWrkDtl("n_wkothrs" + Trim(Str(k))).Value, True) - 10)
                   ElseIf Is_Null(rsWrkDtl("n_wkothrs" + Trim(Str(k))).Value, True) > 15 Then
                      vOt1 = vOt1 + 10
                      vOt2 = vOt2 + 5
                      vOt3 = vOt3 + (Is_Null(rsWrkDtl("n_wkothrs" + Trim(Str(k))).Value, True) - 15)
                   End If
                End If
                
                If Is_Null(rsWrkDtl("n_wksunhrs" + Trim(Str(k))).Value, True) <= 8 Then
                   vSP1 = vSP1 + Is_Null(rsWrkDtl("n_wksunhrs" + Trim(Str(k))).Value, True)
                Else
                   vSP1 = vSP1 + 8
                   vSP2 = vSP2 + (Is_Null(rsWrkDtl("n_wksunhrs" + Trim(Str(k))).Value, True) - 8)
                End If
                
                vTotLateHrs = vTotLateHrs + Is_Null(rsWrkDtl("n_wklatehrs" + Trim(Str(k))).Value, True)
                vTotBasicWrkHrs = vTotBasicWrkHrs + Is_Null(rsWrkDtl("n_wkwrkhrs" + Trim(Str(k))).Value, True)
            Next k
            
            rsWrkDtl("n_othrs15").Value = vOt1
            rsWrkDtl("n_othrs20").Value = vOt2
            rsWrkDtl("n_othrs30").Value = vOt3
            rsWrkDtl("n_sunphhrs20").Value = vSunPh01
            rsWrkDtl("n_sunphhrs30").Value = vSunPh02
            
            If vSalaryType = "HR" Then
               rsWrkDtl("n_wrkhrs").Value = vTotBasicWrkHrs - (vPublicHoliday + vLocalLeave + vSickLeave + vInjuryLeave + vProlongLeave + vWeddingLeave + vMatLeave + vPatLeave + vCompLeave + vOthLeave)
            Else
               rsWrkDtl("n_wrkhrs").Value = vTotBasicWrkHrs
            End If
              
            ' Hrs
            rsWrkDtl("n_lophrs").Value = vLopHrs
            rsWrkDtl("n_latehrs").Value = vTotLateHrs
            rsWrkDtl("n_publicholiday").Value = vPublicHoliday
            rsWrkDtl("n_localleave").Value = vLocalLeave
            rsWrkDtl("n_sickleave").Value = vSickLeave
            rsWrkDtl("n_injuryleave").Value = vInjuryLeave
            rsWrkDtl("n_prolongleave").Value = vProlongLeave
            rsWrkDtl("n_weddingleave").Value = vWeddingLeave
            rsWrkDtl("n_matleave").Value = vMatLeave
            rsWrkDtl("n_patleave").Value = vPatLeave
            rsWrkDtl("n_compleave").Value = vCompLeave
            rsWrkDtl("n_othleave").Value = vOthLeave
            
            'Days
            rsWrkDtl("n_lopdays").Value = vLopDays
            rsWrkDtl("n_sl_fullday").Value = vSlFullDay
            rsWrkDtl("n_no_travelallow").Value = vNoTravAllow
            rsWrkDtl("n_no_mealallow").Value = vNoMealAllow
            rsWrkDtl("n_no_nightallow").Value = vNoNightAllow
            
            rsWrkDtl("c_mode").Value = "S"  ' S - System, M - Manual Entry
            rsWrkDtl("c_usr_id").Value = g_UserName
            rsWrkDtl("d_created").Value = GetDateTime
            
            rsWrkDtl.Update
            CON.CommitTrans
         Loop
         
         ProBar.Value = 25
         DoEvents
         
         
         Call Update_No_Of_Weeks(vProcSql)
         ProBar.Value = 30
         DoEvents
         
      End If
    
      Call Salary_Process
      ProBar.Value = 100
      DoEvents
    
      MsgBox "Process completed Successfully", vbInformation, "Information"
      
      
   Exit Sub

Err_Upd:
   CON.RollbackTrans
   Screen.MousePointer = vbDefault
   MsgBox "Error while creating weekly work hrs - " & Err.Description & vbCrLf & rsWrkDtl("c_empno").Value
End Sub

Private Sub SunPH_Work_Hours_Process(ByVal vSqlFilter As String)
  Dim rsChk As New ADODB.Recordset
  Dim tmpEmpNo As String
  Dim tmpPhDate As Date
  Dim vWrkHrs, vDepTime, vActDepTime, vShEndTime, vShLateMins, vSunPhBreak, vPHHrs As Double
  Dim i As Integer
  
    CON.Execute "SET DATEFIRST 1"
    
    Set rsChk = Nothing
    g_Sql = "Select a.c_empno, a.d_date, a.n_wkday, a.n_workhrs, a.n_deptime, a.c_presabs, " & _
            "c.starthrs, c.endhrs, c.shifthrs, c.breakhrs, c.latehrs, " & _
            "c.break1+c.mins1 break1, c.break2+c.mins2 break2, c.break3+c.mins3 break3, " & _
            "c.mins1, c.mins2, c.mins3 " & _
            "from pr_clock_emp a, pr_emp_mst b, pr_clock_shift c " & _
            "where a.c_empno = b.c_empno and b.c_rec_sta = 'A' and a.c_shift = c.c_shiftcode and " & Trim(vSqlFilter)
  
    g_Sql = g_Sql & " order by a.c_empno, a.d_date "
    
    rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If rsChk.RecordCount > 0 Then
       Do While Not rsChk.EOF()
          tmpEmpNo = Is_Null(rsChk("c_empno").Value, False)
          
          Do While tmpEmpNo = Is_Null(rsChk("c_empno").Value, False)
             vWrkHrs = 0: vDepTime = 0: vActDepTime = 0
             vShEndTime = 0: vShLateMins = 0: vSunPhBreak = 0: vPHHrs = 0
             
             If UCase(Is_Null(rsChk("c_presabs").Value, False)) = "PH" Or Is_Null(rsChk("n_wkday").Value, True) = 1 Then
                tmpPhDate = CDate(Format(rsChk("d_date").Value, "dd/mm/yyyy"))
                
                rsChk.MovePrevious
                If Not (rsChk.BOF) Then
                    If tmpEmpNo = Is_Null(rsChk("c_empno").Value, False) And (tmpPhDate - 1) = CDate(Format(rsChk("d_date").Value, "dd/mm/yyyy")) Then
                       If Is_Null(rsChk("n_deptime").Value, True) > 24 Then
                          vWrkHrs = TimeToMins(Is_Null(rsChk("n_workhrs").Value, True))
                          vActDepTime = TimeToMins(Is_Null(rsChk("n_deptime").Value, True))
                          vShEndTime = Is_Null(rsChk("endhrs").Value, True)
                          vShLateMins = Is_Null(rsChk("latehrs").Value, True)
                          
                          If vActDepTime >= (vShEndTime - vShLateMins) And vActDepTime <= (vShEndTime + vShLateMins) Then
                             vDepTime = vShEndTime
                          Else
                             vDepTime = vActDepTime
                          End If
                          
                          vSunPhBreak = 0
                          If Is_Null(rsChk("break1").Value, True) >= 1440 Then
                             vSunPhBreak = vSunPhBreak + Is_Null(rsChk("mins1").Value, True)
                          End If
                          If Is_Null(rsChk("break2").Value, True) >= 1440 Then
                             vSunPhBreak = vSunPhBreak + Is_Null(rsChk("mins2").Value, True)
                          End If
                          If Is_Null(rsChk("break3").Value, True) >= 1440 Then
                             vSunPhBreak = vSunPhBreak + Is_Null(rsChk("mins3").Value, True)
                          End If
                                                
                          vPHHrs = vDepTime - (1440 + vSunPhBreak)
                          If vWrkHrs < vPHHrs Then
                             vPHHrs = 0
                          End If
                          vPHHrs = MinsToTime(vPHHrs)
                           
                          g_Sql = "update pr_clock_emp set n_sunphhrs = " & vPHHrs & ", n_sunphhrs_dec = " & MinsToDecimal(vPHHrs) & " " & _
                                  "where c_empno = '" & tmpEmpNo & "' and d_date = '" & Is_Date(rsChk("d_date").Value, "S") & "'"
                          CON.Execute g_Sql
                       End If
                    End If
                End If
                rsChk.MoveNext
                If Is_Null(rsChk("n_deptime").Value, True) > 24 Then
                   vWrkHrs = TimeToMins(Is_Null(rsChk("n_workhrs").Value, True))
                   vActDepTime = TimeToMins(Is_Null(rsChk("n_deptime").Value, True))
                   vShEndTime = Is_Null(rsChk("endhrs").Value, True)
                   vShLateMins = Is_Null(rsChk("latehrs").Value, True)
                      
                   If vActDepTime >= (vShEndTime - vShLateMins) And vActDepTime <= (vShEndTime + vShLateMins) Then
                      vDepTime = vShEndTime
                   Else
                      vDepTime = vActDepTime
                   End If
                      
                   vSunPhBreak = 0
                   If Is_Null(rsChk("break1").Value, True) >= 1440 Then
                      vSunPhBreak = vSunPhBreak + Is_Null(rsChk("mins1").Value, True)
                   End If
                   If Is_Null(rsChk("break2").Value, True) >= 1440 Then
                      vSunPhBreak = vSunPhBreak + Is_Null(rsChk("mins2").Value, True)
                   End If
                   If Is_Null(rsChk("break3").Value, True) >= 1440 Then
                      vSunPhBreak = vSunPhBreak + Is_Null(rsChk("mins3").Value, True)
                   End If
                                            
                   vPHHrs = vDepTime - (1440 + vSunPhBreak)
                   vPHHrs = vWrkHrs - vPHHrs
                      
                   If vPHHrs > 0 Then
                      vPHHrs = MinsToTime(vPHHrs)
                   Else
                      vPHHrs = 0
                   End If
                Else
                   vPHHrs = Is_Null(rsChk("n_workhrs").Value, True)
                End If
             End If
             
             g_Sql = "update pr_clock_emp set n_sunphhrs = " & vPHHrs & ", n_sunphhrs_dec = " & MinsToDecimal(vPHHrs) & " " & _
                     "where c_empno = '" & tmpEmpNo & "' and d_date = '" & Is_Date(rsChk("d_date").Value, "S") & "'"
             CON.Execute g_Sql
             
             rsChk.MoveNext
             If rsChk.EOF Then
                Exit Do
             End If
          Loop
       Loop
           
       rsChk.MoveFirst
       Do While Not rsChk.EOF()
          tmpEmpNo = Is_Null(rsChk("c_empno").Value, False)
          
          Do While tmpEmpNo = Is_Null(rsChk("c_empno").Value, False)
             vPHHrs = 0
             If UCase(Is_Null(rsChk("c_presabs").Value, False)) = "PH" Or Is_Null(rsChk("n_wkday").Value, True) = 1 Then
                If Is_Null(rsChk("n_deptime").Value, True) > 24 Then
                   vPHHrs = Is_Null(rsChk("n_workhrs").Value, True)
                   tmpPhDate = CDate(Format(rsChk("d_date").Value, "dd/mm/yyyy"))
                   
                   rsChk.MoveNext
                   If rsChk.EOF Then
                      Exit Do
                   End If
                   If UCase(Is_Null(rsChk("c_presabs").Value, False)) = "PH" Or Is_Null(rsChk("n_wkday").Value, True) = 1 Then
                      g_Sql = "update pr_clock_emp set n_sunphhrs = " & vPHHrs & ", n_sunphhrs_dec = " & MinsToDecimal(vPHHrs) & " " & _
                              "where c_empno = '" & tmpEmpNo & "' and d_date = '" & Is_Date(tmpPhDate, "S") & "'"
                      CON.Execute g_Sql
                   End If
                   rsChk.MovePrevious
                End If
             End If
             
             rsChk.MoveNext
             If rsChk.EOF Then
                Exit Do
             End If
          Loop
       Loop
    End If
End Sub

Private Sub Update_No_Of_Weeks(ByVal vSqlFilter As String)
  Dim rsChk As New ADODB.Recordset
  Dim vNoWeek As Integer, vProcSql As String, tmpEmpNo As String, tmpStr As String
  Dim i As Integer
  Dim tmpDate As Date, tmpFirstDate As Date
    
    CON.Execute "SET DATEFIRST 1"
    vNoWeek = 0
    
    
    ' Update to all employee
    Set rsChk = Nothing
    g_Sql = "select * from pr_payperiod_dtl where n_period = " & Trim(vPayPeriod) & " and c_type = 'W'"
    rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If rsChk.RecordCount > 0 Then
       tmpDate = rsChk("d_fromdate").Value
       tmpFirstDate = rsChk("d_fromdate").Value
       For tmpDate = rsChk("d_fromdate").Value To rsChk("d_todate").Value
           If Weekday(tmpDate) = 2 Then
              vNoWeek = vNoWeek + 1
           End If
       Next tmpDate
       If Weekday(tmpDate) > 2 Then
          vNoWeek = vNoWeek + 1
       End If
    Else
       MsgBox "Pay Period not found ", vbCritical, "Critical"
    End If
    
    g_Sql = "update pr_workhrs_dtl set n_noweek = " & vNoWeek & " " & _
            "from pr_workhrs_dtl a, pr_emp_mst b " & _
            "where a.c_empno = b.c_empno and b.c_rec_sta = 'A' and " & Trim(vSqlFilter)
    CON.Execute g_Sql
    
    
    ' Update for who join in between month
       
    Set rsChk = Nothing
    g_Sql = "select a.c_empno, a.d_date " & _
            "from pr_clock_emp a, pr_emp_mst b " & _
            "where a.c_empno = b.c_empno and b.d_doj > '" & Format(tmpFirstDate, "yyyy-mm-dd") & "' and " & Trim(vSqlFilter) & " " & _
            "order by a.c_empno, a.d_date "
    rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    Do While Not rsChk.EOF()
       vNoWeek = 0
       tmpEmpNo = Is_Null(rsChk("c_empno").Value, False)
       tmpDate = rsChk("d_date").Value
       Do While tmpEmpNo = Is_Null(rsChk("c_empno").Value, False)
          If Weekday(rsChk("d_date").Value) = 2 Then
             vNoWeek = vNoWeek + 1
          End If
          rsChk.MoveNext
          If rsChk.EOF Then
             Exit Do
          End If
       Loop
       If Weekday(tmpDate) > 2 Then
          vNoWeek = vNoWeek + 1
       End If
       tmpStr = "update pr_workhrs_dtl set n_noweek = " & vNoWeek & " " & _
                "where n_period = " & Trim(vPayPeriod) & " and c_empno = '" & Trim(tmpEmpNo) & "'"
       CON.Execute tmpStr
    Loop
    
End Sub

Private Sub CalLastMonthOtPay(ByVal vEmpNo As String, ByVal WkOneWrkHrs As Double)
  Dim rsPreAttend As New ADODB.Recordset
  Dim vLWkWrkHrs, vLWkNorHrs, vLWkOT1, vLWkOT2, vLWkOT3, vCWkWrkHrs, vCWkNorHrs, vCWkOT1, vCWkOT2, vCWkOT3 As Double
      
      vLWkWrkHrs = 0: vLWkNorHrs = 0: vLWkOT1 = 0: vLWkOT2 = 0: vLWkOT3 = 0
      vCWkWrkHrs = 0: vCWkNorHrs = 0: vCWkOT1 = 0: vCWkOT2 = 0: vCWkOT3 = 0
      
      Set rsPreAttend = Nothing
      g_Sql = "select n_wkwrkhrs4, n_wkothrs4, n_wkwrkhrs5, n_wkothrs5, n_wkwrkhrs6, n_wkothrs6, n_noweek " & _
              "from pr_workhrs_dtl where c_empno = '" & Trim(vEmpNo) & "' and n_period = " & vPrevPayPeriod
      rsPreAttend.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
      If rsPreAttend.RecordCount > 0 Then
         If Is_Null(rsPreAttend("n_noweek").Value, True) = 6 Then
            vLWkWrkHrs = Is_Null(rsPreAttend("n_wkwrkhrs6").Value, True) + Is_Null(rsPreAttend("n_wkothrs6").Value, True)
         ElseIf Is_Null(rsPreAttend("n_noweek").Value, True) = 5 Then
            vLWkWrkHrs = Is_Null(rsPreAttend("n_wkwrkhrs5").Value, True) + Is_Null(rsPreAttend("n_wkothrs5").Value, True)
         ElseIf Is_Null(rsPreAttend("n_noweek").Value, True) = 4 Then
            vLWkWrkHrs = Is_Null(rsPreAttend("n_wkwrkhrs4").Value, True) + Is_Null(rsPreAttend("n_wkothrs4").Value, True)
         Else
            vLWkWrkHrs = 0
         End If
      End If
       
      If vLWkWrkHrs > 45 Then
         vLWkNorHrs = 45
         If vLWkWrkHrs <= 55 Then
            vLWkOT1 = vLWkWrkHrs - 45
         ElseIf vLWkWrkHrs > 55 And vLWkWrkHrs <= 60 Then
            vLWkOT1 = 10
            vLWkOT2 = vLWkWrkHrs - 55
         ElseIf vLWkWrkHrs > 60 Then
            vLWkOT1 = 10
            vLWkOT2 = 5
            vLWkOT3 = vLWkWrkHrs - 60
         End If
      Else
         vLWkNorHrs = vLWkWrkHrs
      End If
       
      vCWkWrkHrs = vLWkWrkHrs + WkOneWrkHrs
      If vCWkWrkHrs > 45 Then
         vCWkNorHrs = 45
         If vCWkWrkHrs <= 55 Then
            vCWkOT1 = vCWkWrkHrs - 45
         ElseIf vCWkWrkHrs > 55 And vCWkWrkHrs <= 60 Then
            vCWkOT1 = 10
            vCWkOT2 = vCWkWrkHrs - 55
         ElseIf vCWkWrkHrs > 60 Then
            vCWkOT1 = 10
            vCWkOT2 = 5
            vCWkOT3 = vCWkWrkHrs - 60
         End If
      Else
         vCWkNorHrs = vCWkWrkHrs
      End If
      
      vNetNorHrs = vCWkNorHrs - vLWkNorHrs
      vNetOT1 = vCWkOT1 - vLWkOT1
      vNetOT2 = vCWkOT2 - vLWkOT2
      vNetOT3 = vCWkOT3 - vLWkOT3
End Sub

Private Sub Salary_Process()
On Error GoTo Err_Proc
   Dim rsProc As New ADODB.Recordset
   Dim vCompany As String, vBranch As String, vDept As String, vDesig As String, vEmpType As String, vEmpNo As String
   Dim tmpStr As String
   
      If Trim(Cmb_Company) = "" Then
         vCompany = "A"
      Else
         vCompany = Trim(Right(Trim(Cmb_Company), 7))
      End If
      
      If Trim(Cmb_Branch) = "" Then
         vBranch = "A"
      Else
         vBranch = Trim(Cmb_Branch)
      End If
       
      If Trim(Cmb_Dept) = "" Then
         vDept = "A"
      Else
         vDept = Trim(Cmb_Dept)
      End If
      
      If Trim(Cmb_Desig) = "" Then
         vDesig = "A"
      Else
         vDesig = Trim(Cmb_Desig)
      End If
      
      If Trim(Cmb_EmpType) = "" Then
         vEmpType = "A"
      Else
         vEmpType = Trim(Cmb_EmpType)
      End If
       
      If Trim(Txtc_EmployeeName) = "" Then
         vEmpNo = "A"
      Else
         vEmpNo = Trim(Right(Trim(Txtc_EmployeeName), 7))
      End If
      
      tmpStr = Trim(vPayPeriod) & ", '" & _
               Trim(vCompany) & "', '" & _
               Trim(vBranch) & "', '" & _
               Trim(vDept) & "', '" & _
               Trim(vDesig) & "', '" & _
               Trim(vEmpType) & "', '" & _
               Trim(vEmpNo) & "', '" & _
               Trim(g_UserName)
      
      If Val(Txtc_Month) = 13 Then
         g_Sql = "HR_EOY_BONUS_PROCESS " & tmpStr & "'"
         CON.Execute g_Sql
      Else
         g_Sql = "HR_PAYROLL_PROCESS " & tmpStr & "'"
         CON.Execute g_Sql
      End If
      
   Exit Sub

Err_Proc:
     Screen.MousePointer = vbDefault
     MsgBox Err.Description, vbCritical, "Error"
End Sub

