VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_Report_Generate 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12600
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3990
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
      Height          =   1545
      Left            =   375
      TabIndex        =   9
      Top             =   645
      Width           =   11760
      Begin VB.ComboBox Cmb_EmpType 
         Height          =   315
         Left            =   8850
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   675
         Width           =   2460
      End
      Begin VB.ComboBox Cmb_Company 
         Height          =   315
         Left            =   4620
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   675
         Width           =   2505
      End
      Begin VB.TextBox Txtc_Month 
         Height          =   300
         Left            =   1440
         TabIndex        =   0
         Top             =   675
         Width           =   555
      End
      Begin VB.TextBox Txtc_Year 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   2055
         TabIndex        =   1
         Top             =   675
         Width           =   765
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
         TabIndex        =   12
         Top             =   705
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
         TabIndex        =   11
         Top             =   705
         Width           =   1035
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
         TabIndex        =   10
         Top             =   705
         Width           =   570
      End
   End
   Begin VB.Frame Fme_Generate 
      Height          =   1350
      Left            =   375
      TabIndex        =   8
      Top             =   2190
      Width           =   11760
      Begin VB.CommandButton Btn_Preview 
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
         Left            =   5145
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   540
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
         Left            =   6705
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   540
         Width           =   1155
      End
      Begin VB.CommandButton Btn_Generate 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Generate"
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
         Left            =   3570
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   4
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
      TabIndex        =   7
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
Attribute VB_Name = "frm_Report_Generate"
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
Private vMonthName As String, vCompShortName As String

Private Sub Form_Load()
   
   ' PAYE
   If RepName = "salpayecsv" Then
      lbl_scr_name.Caption = "PAYE CSV File Generation"
      Call GenViewExitButtonStatus(False)
      Cmb_EmpType.Enabled = False
   
   ' NPF
   ElseIf RepName = "salnpfcsv" Then
      lbl_scr_name.Caption = "NPF CSV File Generation"
      Call GenViewExitButtonStatus(False)
      Cmb_EmpType.Enabled = False
   
   ' EPZ
   ElseIf RepName = "salepzcsv" Then
      lbl_scr_name.Caption = "EPZ CSV File Generation"
      Call GenViewExitButtonStatus(False)
      Cmb_EmpType.Enabled = False
   
   ' EPZ Loan
   ElseIf RepName = "salepzloancsv" Then
      lbl_scr_name.Caption = "EPZ Loan CSV File Generation"
      Call GenViewExitButtonStatus(False)
      Cmb_EmpType.Enabled = False
   
   ' MCB
   ElseIf RepName = "salmcbpaytr" Then
      lbl_scr_name.Caption = "MCB Bank Transfer"
      Call GenViewExitButtonStatus(False)
   
   ' SBM
   ElseIf RepName = "salsbmpaytr" Then
      lbl_scr_name.Caption = "SBM Bank Transfer"
      Call GenViewExitButtonStatus(False)
   
   ' Barclays
   ElseIf RepName = "salbarpaytr" Then
      lbl_scr_name.Caption = "Barclays Bank Transfer"
      Call GenViewExitButtonStatus(False)
   
   ' Emoulment
   ElseIf RepName = "salemotran" Then
      lbl_scr_name.Caption = "Barclays Bank Transfer"
      Call GenViewExitButtonStatus(True)
      Txtc_Month.Visible = False
      Lbl_Period.Left = Lbl_Period.Left + 600
      Txtc_Month.Text = "07"
   
   End If
   
   Call Combo_Load
   Call TGControlProperty(Me)
   Cmb_Company.ListIndex = 1
   
   vMonthName = ""
End Sub

Private Sub Btn_Generate_Click()
  Dim tmpStr As String, RepOpt As String
  
   If Not ChkPayPeriod_Entered Then
      Exit Sub
   End If
   
   ' PAYE
   If RepName = "salpayecsv" Then
      Call PayeDataTransfer_CSV
      Exit Sub
   
   ' NPF
   ElseIf RepName = "salnpfcsv" Then
      Call NpfDataTransfer_CSV
      Exit Sub
   
   ' EPZ
   ElseIf RepName = "salepzcsv" Then
      Call EPZ_Contrib_DataTran_CSV
      Exit Sub
      
   ' EPZ Loan
   ElseIf RepName = "salepzloancsv" Then
      Call EPZ_Loan_DataTran_CSV
      Exit Sub
   
   ' Emoulment
   ElseIf RepName = "salemotran" Then
      Call Emoulment_Statement_Gen
      Exit Sub
   
   ' MCB Bank Transfer
   ElseIf RepName = "salmcbpaytr" Then
   
      If Trim(Cmb_Company) = "" Then
         MsgBox "Please select company", vbInformation, "Information"
         Exit Sub
      End If
      
      If Not MCBDataTransfer Then
         Exit Sub
      End If
   
      tmpStr = "1. MCB Bank Transfer List" & vbCrLf & _
               "2. Discrepancy List "
      RepOpt = InputBox(tmpStr, "Select Your Option", "1")
      
      If Val(RepOpt) = 0 Then
         Exit Sub
      End If
      
      RepDate = Trim(UCase(Left(Cmb_Company, 50))) & "  -  " & "Period for " & vMonthName & " " & Trim(Txtc_Year)
      SelFor = "{MCBPAYTR.RECTYPE} <> '9'"
      SelFor = SelFor & " AND {PR_SALARY_MST.N_PERIOD} = " & vPayPeriod
      If Trim(Cmb_EmpType) <> "" Then
         SelFor = SelFor & " AND {PR_SALARY_MST.C_EMPTYPE} = '" & Trim(Cmb_EmpType) & "'"
      End If
      
      If Val(RepOpt) = 1 Then
         Call Print_Rpt(SelFor, "Pr_MCBPayTr.rpt")
      ElseIf Val(RepOpt) = 2 Then
         SelFor = SelFor & " AND ({MCBPAYTR.ACNO}='' OR Val({MCBPAYTR.AMOUNT})=0) "
         Call Print_Rpt(SelFor, "Pr_MCBPayTr.rpt")
      End If
  
   ElseIf RepName = "salscbpaytr" Then
      If Trim(Cmb_Company) = "" Then
         MsgBox "Please select company", vbInformation, "Information"
         Exit Sub
      End If
   
      If Not SBMDataTransfer Then
         Exit Sub
      End If
      
      tmpStr = "1. SBM Bank Transfer List" & vbCrLf & _
               "2. Discrepancy List "
      RepOpt = InputBox(tmpStr, "Select Your Option", "1")
      
      RepDate = Trim(UCase(Left(Cmb_Company, 50))) & "  -  " & "Period for " & vMonthName & " " & Trim(Txtc_Year)
      
      If Val(RepOpt) = 1 Then
         Call Print_Rpt(SelFor, "Pr_SCBPayTr.rpt")
      ElseIf Val(RepOpt) = 2 Then
         Call Save_SCBPaytr_Discp
         Call Print_Rpt(SelFor, "Pr_SCBPayTr_Discp.rpt")
      Else
         Exit Sub
      End If
   
   ElseIf RepName = "salbarpaytr" Then
      RepDate = Trim(UCase(Left(Cmb_Company, 50))) & "  -  " & "Period for " & vMonthName & " " & Trim(Txtc_Year)
      SelFor = "{V_PR_SALARY_MST.N_PERIOD} = " & vPayPeriod & " AND {V_PR_SALARY_MST.C_BANK} = 'B03' "
      If Trim(Cmb_Company) <> "" Then
         SelFor = SelFor & " AND {V_PR_SALARY_MST.C_COMPANY} = '" & Trim(Right(Trim(Cmb_Company), 7)) & "'"
      End If
      If Trim(Cmb_EmpType) <> "" Then
         SelFor = SelFor & " AND {V_PR_SALARY_MST.C_EMPTYPE} = '" & Trim(Cmb_EmpType) & "'"
      End If
      
      Call Print_Rpt(SelFor, "Pr_BarPayTr.rpt")
   
   Else
      Exit Sub
   End If
   
   If Trim(RepTitle) <> "" Then
      Mdi_Ta_HrPay.CRY1.Formulas(1) = "ReportHead='" & UCase(Trim(RepTitle)) & "'"
   End If

   If Trim(RepDate) <> "" Then
      Mdi_Ta_HrPay.CRY1.Formulas(2) = "RepHeadDate='" & Trim(RepDate) & "'"
   End If
   Mdi_Ta_HrPay.CRY1.Action = 1
   
End Sub

Private Sub Btn_Preview_Click()
  Dim tmpStr As String, RepOpt As String
  
   If Not ChkPayPeriod_Entered Then
      Exit Sub
   End If

   ' Emoulment Statements
   If RepName = "salemotran" Then
   
      If Trim(Cmb_Company) = "" Then
         MsgBox "Please select company", vbInformation, "Information"
         Exit Sub
      End If
   
      tmpStr = "1. Emolument Working List " & vbCrLf & _
               "2. Emolument Working List - PAYE only " & vbCrLf & _
               "3. Emolument Statement " & vbCrLf & vbCrLf & _
               "4. Emolument Statement - PAYE only" & vbCrLf & vbCrLf & _
               "5. Yearly Return CSV File "
   
      RepOpt = InputBox(tmpStr, "Select Your Option", "1")
      If Val(RepOpt) = 0 Then
         Exit Sub
      End If
      
      Call PAYE_Yr_Retn_RepFilter
   
      If Val(RepOpt) = 1 Then
         Call Print_Rpt(SelFor, "Pr_Paye_Emoul_List.rpt")
      ElseIf Val(RepOpt) = 2 Then
         SelFor = SelFor & " AND {PR_PAYE_RETN_MRA_DTL.N_PAYE} >0 "
         Call Print_Rpt(SelFor, "Pr_Paye_Emoul_List.rpt")
      
      ElseIf Val(RepOpt) = 3 Then
         Call Print_Rpt(SelFor, "Pr_Paye_EmoulState.rpt")
      ElseIf Val(RepOpt) = 4 Then
         SelFor = SelFor & " AND {PR_PAYE_RETN_MRA_DTL.N_PAYE} >0 "
         Call Print_Rpt(SelFor, "Pr_Paye_EmoulState.rpt")
         
      ElseIf Val(RepOpt) = 5 Then
         Call PAYE_YR_DataTran_CSV
         Exit Sub
      Else
         Exit Sub
      End If
   
   Else
      Exit Sub
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
    Call LoadComboCompany(Me)
    Call LoadComboEmpType(Me)
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

Private Function ChkPayPeriod_Entered() As Boolean
    If vPayPeriod = 0 Then
       MsgBox "Period should not be blank", vbInformation, "Information"
       If Txtc_Month.Visible = True Then
          Txtc_Month.SetFocus
       Else
          Txtc_Year.SetFocus
       End If
       Exit Function
    End If
    ChkPayPeriod_Entered = True
End Function


Private Function MCBDataTransfer() As Boolean
On Error GoTo Err_Flag
  Dim rsChk As New ADODB.Recordset, rsCsv As New ADODB.Recordset
  Dim vFilePath As String
  Dim nCtr As Integer
  Dim nTotalAmount As Double

    Screen.MousePointer = vbHourglass
    g_SaveFlagNull = False

    CON.BeginTrans
    CON.Execute "truncate table mcbpaytr"
    CON.CommitTrans

    Set rsCsv = Nothing
    g_Sql = "select * from mcbpaytr "
    rsCsv.Open g_Sql, CON, adOpenDynamic, adLockOptimistic

    Set rsChk = Nothing
    g_Sql = "select a.c_empno, b.c_name, b.c_othername, a.c_bank, b.c_bankcode, b.c_acctno, " & _
            "a.n_net, floor(a.n_net) n_netfloor " & _
            "from pr_salary_mst a, pr_emp_mst b " & _
            "where a.c_empno = b.c_empno and a.c_paytype = 'BA' and a.c_bank = 'B02' and a.n_net > 0 and b.c_rec_sta = 'A' and " & _
            "a.n_period = " & vPayPeriod
            
    If Trim(Cmb_Company) <> "" Then
       g_Sql = g_Sql & " and a.c_company = '" & Right(Trim(Cmb_Company), 7) & "' "
    End If
    If Trim(Cmb_EmpType) <> "" Then
       g_Sql = g_Sql & " and a.c_emptype = '" & Trim(Cmb_EmpType) & "' "
    End If
    g_Sql = g_Sql & " order by a.c_empno "

    rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If rsChk.RecordCount = 0 Then
       Screen.MousePointer = vbDefault
       MsgBox "Details are not available to generate", vbInformation, "Information"
       Exit Function
    End If
    rsChk.MoveFirst
    nCtr = 0: nTotalAmount = 0

    CON.BeginTrans

    Do While Not rsChk.EOF
       nCtr = nCtr + 1
       rsCsv.AddNew
       rsCsv("slno").Value = Format(nCtr, "00000")
       rsCsv("rectype").Value = "1"
       If Is_Null(rsChk("c_bankcode").Value, False) = "" Then
          rsCsv("acno").Value = Is_Null(rsChk("c_acctno").Value, False)
       Else
          rsCsv("acno").Value = Is_Null(rsChk("c_bankcode").Value, False)
          rsCsv("bacno").Value = Is_Null(rsChk("c_acctno").Value, False)
          rsCsv("bname").Value = Left(Is_Null(rsChk("c_name").Value, False) + " " + Is_Null(rsChk("c_othername").Value, False), 22)
       End If
       rsCsv("amount").Value = AmountToBankChar(Is_Null(rsChk("n_net").Value, True), Is_Null(rsChk("n_netfloor").Value, True))
       nTotalAmount = nTotalAmount + Is_Null(rsChk("n_net").Value, True)
       rsCsv("remid").Value = "A01"
       If Val(Txtc_Month) = 13 Then
          rsCsv("paymonth").Value = "12"
       Else
          rsCsv("paymonth").Value = Format(Trim(Txtc_Month), "00")
       End If
       rsCsv("c_empno").Value = Is_Null(rsChk("c_empno").Value, False)

       rsCsv.Update
       rsChk.MoveNext
    Loop
    
    g_Sql = "update mcbpaytr set mcbpaytr = IsNull(rectype,'') + IsNull(acno,'') + " & _
            "IsNull(amount,'') + IsNull(remid,'') + IsNull(bacno,'') +IsNull(bname,'')+ IsNull(paymonth,'') "
    CON.Execute g_Sql

    rsCsv.AddNew
    rsCsv("slno").Value = Format(nCtr + 1, "00000")
    rsCsv("rectype").Value = "9"
    rsCsv("acno").Value = Format(nCtr, "00000")
    rsCsv("amount").Value = AmountToBankChar(nTotalAmount, Floor(nTotalAmount))
    rsCsv("mcbpaytr").Value = "9" & Format(nCtr, "00000") & AmountToBankChar(nTotalAmount, Floor(nTotalAmount))
    rsCsv.Update
    
    
    CON.Execute "truncate table pr_export_csv"
    
    g_Sql = "Insert into pr_export_csv (n_seq, c_csv) " & _
            "Select Convert(int, slno), mcbpaytr From mcbpaytr "
    CON.Execute g_Sql
    
    CON.CommitTrans

    Screen.MousePointer = vbDefault
    g_SaveFlagNull = False

    comDialog.FileName = "MCBPAYTR"
    comDialog.ShowSave
    vFilePath = comDialog.FileName

    Screen.MousePointer = vbHourglass
    Call CsvFileExport_Process(vFilePath)
    Screen.MousePointer = vbDefault
    
    MsgBox "Transfered MCB format Completed Successfully ", vbInformation, "Information"
    MCBDataTransfer = True

  Exit Function

Err_Flag:
   Screen.MousePointer = vbDefault
   g_SaveFlagNull = False
   CON.RollbackTrans
   MsgBox "Error while Generating ACII File Process " & Err.Description & vbCrLf & rsChk("c_empno").Value
End Function

Private Function SBMDataTransfer() As Boolean
On Error GoTo Err_Flag
  Dim rsChk As New ADODB.Recordset, rsCsv As New ADODB.Recordset
  Dim vFilePath As String, AsciiFileName As String
  Dim tmpAmount As Double, tmpAmountStr As String
  Dim vSeq As Long

    Screen.MousePointer = vbHourglass
    g_SaveFlagNull = False

    CON.BeginTrans
    CON.Execute "truncate table scbpaytr"
    CON.CommitTrans

    Set rsChk = Nothing
    g_Sql = "select c_companyname from pr_company_mst where c_company = '" & Trim(Right(Trim(Cmb_Company), 7)) & "'"
    rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If rsChk.RecordCount > 0 Then
       AsciiFileName = UCase(rsChk("c_companyname").Value)
    End If
    AsciiFileName = Trim(AsciiFileName) & Trim(Format(Trim(Txtc_Month), "00"))
    AsciiFileName = Trim(AsciiFileName) & "01"


    Set rsCsv = Nothing
    g_Sql = "select * from scbpaytr "
    rsCsv.Open g_Sql, CON, adOpenDynamic, adLockOptimistic

    Set rsChk = Nothing
    g_Sql = "select a.c_empno, b.c_name, b.c_othername, a.c_bank, b.c_bankcode, b.c_acctno, a.n_net " & _
            "from pr_salary_mst a, pr_emp_mst b " & _
            "where a.c_empno = b.c_empno and a.c_paytype = 'BA' and a.c_bank = 'B01' and a.n_net > 0 and b.c_rec_sta = 'A' and " & _
            "a.n_period = " & vPayPeriod
    If Trim(Cmb_EmpType) <> "" Then
       g_Sql = g_Sql & " and a.c_emptype = '" & Trim(Cmb_EmpType) & "' "
    End If
    g_Sql = g_Sql & " order by b.c_bankcode, a.c_empno"

    rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If rsChk.RecordCount = 0 Then
       Screen.MousePointer = vbDefault
       MsgBox "Details are not available to generate", vbInformation, "Information"
       Exit Function
    End If
    rsChk.MoveFirst

    CON.BeginTrans
    
    Do While Not rsChk.EOF
       rsCsv.AddNew
       rsCsv("acno").Value = Is_Null(rsChk("c_acctno").Value, False)

       tmpAmountStr = Replace(Trim(Format(Is_Null(rsChk("n_net").Value, True), "00000000000.00")), ".", "")
       tmpAmount = Val(tmpAmountStr)

       If tmpAmount < 9 Then
          tmpAmountStr = Space(15) & Format(tmpAmount, "0")
       ElseIf tmpAmount < 99 Then
          tmpAmountStr = Space(14) & Format(tmpAmount, "00")
       ElseIf tmpAmount < 999 Then
          tmpAmountStr = Space(13) & Format(tmpAmount, "000")
       ElseIf tmpAmount < 9999 Then
          tmpAmountStr = Space(12) & Format(tmpAmount, "0000")
       ElseIf tmpAmount < 99999 Then
          tmpAmountStr = Space(11) & Format(tmpAmount, "00000")
       ElseIf tmpAmount < 999999 Then
          tmpAmountStr = Space(10) & Format(tmpAmount, "000000")
       ElseIf tmpAmount < 9999999 Then
          tmpAmountStr = Space(9) & Format(tmpAmount, "0000000")
       ElseIf tmpAmount < 99999999 Then
          tmpAmountStr = Space(8) & Format(tmpAmount, "00000000")
       ElseIf tmpAmount < 999999999 Then
          tmpAmountStr = Space(7) & Format(tmpAmount, "000000000")
       Else
          tmpAmountStr = Space(6) & Format(tmpAmount, "0000000000")
       End If

       If Len(tmpAmountStr) <> 16 Then
          CON.RollbackTrans
          MsgBox "Amount format is not valid " & rsChk("c_empno").Value & " ", vbInformation, "Information"
          Exit Function
       End If

       rsCsv("amount").Value = tmpAmountStr  'don't put trim
       rsCsv("empname").Value = Left(Is_Null(rsChk("c_name").Value, False) & " " & Replace(Is_Null(rsChk("c_othername").Value, False), ".", " "), 40)
       If Is_Null(rsChk("c_bankcode").Value, False) = "" Then
          rsCsv("bankcode").Value = "11"
       Else
          rsCsv("bankcode").Value = Right(Is_Null(rsChk("c_bankcode").Value, False), 2)
       End If
       rsCsv("comname").Value = UCase(Trim(Left(Trim(Cmb_Company), 50)))
       rsCsv("ref").Value = "Wages"
       rsCsv("flag").Value = "N"

       rsCsv.Update
       rsChk.MoveNext
    Loop

    g_Sql = "update scbpaytr set scbpaytr = IsNull(acno,'') + IsNull(amount,'') + " & _
            "IsNull(empname,'') + IsNull(bankcode,'') + IsNull(comname,'') + IsNull(ref,'') + IsNull(flag,'') "
    CON.Execute g_Sql
    
    CON.CommitTrans
    
    
    
    CON.BeginTrans
    
    CON.Execute "truncate table pr_export_csv"
    
    vSeq = 0
    Set rsCsv = Nothing
    g_Sql = "select * from pr_export_csv"
    rsCsv.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
    
    Set rsChk = Nothing
    g_Sql = "select scbpaytr from scbpaytr "
    rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    
    rsChk.MoveFirst
    Do While Not rsChk.EOF
       rsCsv.AddNew
       
       vSeq = vSeq + 1
       rsCsv("n_seq").Value = vSeq
       rsCsv("c_csv").Value = rsChk("scbpaytr").Value
       
       rsCsv.Update
       rsChk.MoveNext
    Loop
    
    CON.CommitTrans

    Screen.MousePointer = vbDefault
    g_SaveFlagNull = False


    comDialog.FileName = AsciiFileName
    comDialog.ShowSave
    vFilePath = comDialog.FileName

    Screen.MousePointer = vbHourglass
    Call CsvFileExport_Process(vFilePath)
    Screen.MousePointer = vbDefault
    
    MsgBox "Transfered SBM format Completed Successfully ", vbInformation, "Information"
    SBMDataTransfer = True


  Exit Function

Err_Flag:
   Screen.MousePointer = vbDefault
   g_SaveFlagNull = False
   CON.RollbackTrans
   MsgBox "Error while Generating ACII File Process " & Err.Description & vbCrLf & rsChk("c_empno").Value
End Function

Private Sub NpfDataTransfer_CSV()
On Error GoTo Err_Flag
  Dim rsChk As New ADODB.Recordset, rsCsv As New ADODB.Recordset
  Dim vFilePath As String, tmpStr As String
  Dim tmpAmount As Double, tmpBasic As Double, vLevy As Double, vLevyBasic As Double
  Dim vEmpNps As Long, vEmpNpsMax As Long, vEmpEwfMax As Long, vComNps As Long, vComNpsMax As Long, vComEwfMax As Long
  Dim vSeq As Long
  
    vSeq = 0
    Screen.MousePointer = vbHourglass
    g_SaveFlagNull = False

    CON.BeginTrans
    CON.Execute "truncate table pr_export_csv"
    CON.CommitTrans

    ' to get master input values
    Set rsChk = Nothing
    g_Sql = "select c_companyname, n_empnpf, n_empnpfmax, n_empmedmax, n_comnpf, n_comnpfmax, n_commedmax " & _
            "from pr_company_mst where c_company = '" & Right(Trim(Cmb_Company), 7) & "' "
    rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If rsChk.RecordCount > 0 Then
       vCompShortName = Is_Null(rsChk("c_companyname").Value, False)
       vEmpNps = Is_Null(rsChk("n_empnpf").Value, True)
       vEmpNpsMax = Is_Null(rsChk("n_empnpfmax").Value, True)
       vEmpEwfMax = Is_Null(rsChk("n_empmedmax").Value, True)
       vComNps = Is_Null(rsChk("n_comnpf").Value, True)
       vComNpsMax = Is_Null(rsChk("n_comnpfmax").Value, True)
       vComEwfMax = Is_Null(rsChk("n_commedmax").Value, True)
    End If

    Set rsCsv = Nothing
    g_Sql = "select * from pr_export_csv "
    rsCsv.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
    
    'details
    vLevy = 0: vLevyBasic = 0
    Set rsChk = Nothing
    g_Sql = "select a.c_company, a.c_empno, b.c_nicno, b.c_itno, b.c_nationality, b.c_dept, " & _
            "b.c_name, b.c_othername, a.c_emptype, " & _
            "a.n_paye, a.n_comnps, a.n_empnps, a.n_comewf+a.n_empewf n_comewf, a.n_comlevy, round(round(a.n_earnbasic,2),0) n_basic " & _
            "from pr_salary_mst a, pr_emp_mst b " & _
            "where a.c_empno = b.c_empno and a.n_period = " & vPayPeriod & " and " & _
            "b.c_rec_sta = 'A' and (a.n_empnps+a.n_comnps+a.n_empewf+a.n_comewf)>0 "

    If Trim(Cmb_Company) <> "" Then
       g_Sql = g_Sql & " and a.c_company = '" & Right(Trim(Cmb_Company), 7) & "' "
    End If
    g_Sql = g_Sql & " order by a.c_empno "

    rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If rsChk.RecordCount = 0 Then
       Screen.MousePointer = vbDefault
       MsgBox "Details are not available to generate", vbInformation, "Information"
       Exit Sub
    End If
    rsChk.MoveFirst

    CON.BeginTrans

    Do While Not rsChk.EOF
       tmpStr = Is_Null(rsChk("c_nicno").Value, False)
       tmpStr = tmpStr & "," & Left(Is_Null(rsChk("c_othername").Value, False), 35)
       tmpStr = tmpStr & "," & Left(Is_Null(rsChk("c_name").Value, False), 35)

       If Is_Null(rsChk("n_comnps").Value, True) >= vComNpsMax Then
          tmpAmount = vComNpsMax
          tmpBasic = ((vComNpsMax * 100) / vComNps)
       Else
          tmpAmount = Is_Null(rsChk("n_comnps").Value, True)
          tmpBasic = Is_Null(rsChk("n_basic").Value, True)
       End If

       If Is_Null(rsChk("n_empnps").Value, True) >= vEmpNpsMax Then
          tmpAmount = vEmpNpsMax
       Else
          tmpAmount = Is_Null(rsChk("n_empnps").Value, True)
       End If

       If Is_Null(rsChk("n_comewf").Value, True) >= (vEmpEwfMax + vComEwfMax) Then
          tmpAmount = (vEmpEwfMax + vComEwfMax)
       Else
          tmpAmount = Is_Null(rsChk("n_comewf").Value, True)
       End If
       
       
       tmpStr = tmpStr & "," & Trim(Str(Round(tmpBasic, 0)))
       If Is_Null(rsChk("n_comewf").Value, True) > 0 Then
          tmpStr = tmpStr & "," & Trim(Str(Round(tmpBasic, 0)))
       Else
          tmpStr = tmpStr & "," & Trim(Str(Round(0, 2)))
       End If
        
       vLevy = vLevy + Is_Null(rsChk("n_comlevy").Value, True)
       vLevyBasic = Round(((vLevy * 100) / 1.5), 0)

       tmpStr = tmpStr & "," & "S2"
       tmpStr = tmpStr & "," & "M"
       tmpStr = tmpStr & "," & Trim(Str(Val(Txtc_Month)))

       rsCsv.AddNew
       
       vSeq = vSeq + 1
       rsCsv("n_seq").Value = vSeq
       rsCsv("c_refno") = "1"
       rsCsv("c_csv") = tmpStr

       rsCsv.Update
       rsChk.MoveNext
    Loop

    'master
    Set rsChk = Nothing
    g_Sql = "select c_displayname, c_tan from pr_company_mst where c_company = '" & Right(Trim(Cmb_Company), 7) & "' "
    rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If rsChk.RecordCount > 0 Then
       tmpStr = "NPFINFO,"
       tmpStr = tmpStr & Is_Null(rsChk("c_tan").Value, False) & ","
       tmpStr = tmpStr & Is_Null(rsChk("c_displayname").Value, False) & ","
       tmpStr = tmpStr & Format(Val(Txtc_Year), "0000") & Format(Val(Txtc_Month), "00") & ","
       tmpStr = tmpStr & Trim(Str(vLevyBasic)) & ","
       tmpStr = tmpStr & Trim(Str(vLevy))
       
       rsCsv.AddNew
       
       rsCsv("n_seq") = 0
       rsCsv("c_refno").Value = "0"
       rsCsv("c_csv") = tmpStr
       
       rsCsv.Update
    End If

    CON.CommitTrans
    Screen.MousePointer = vbDefault
    g_SaveFlagNull = False

    MsgBox "NPF CSV File Format Processed Successfully", vbInformation, "Information"

    If vCompShortName <> "" Then
       comDialog.FileName = "NPF_" & Trim(vCompShortName) & ".CSV"
    Else
       comDialog.FileName = "NPF.CSV"
    End If
    comDialog.ShowSave
    vFilePath = comDialog.FileName
    
    Screen.MousePointer = vbHourglass
    Call CsvFileExport_Process(vFilePath)
    Screen.MousePointer = vbDefault

    MsgBox "Transfered to CSV File format Completed Successfully ", vbInformation, "Information"
  
  Exit Sub

Err_Flag:
   Screen.MousePointer = vbDefault
   g_SaveFlagNull = False
   MsgBox "Error while Generating CSV File Process - " & Err.Number & Space(5) & Err.Description
   CON.RollbackTrans
End Sub

Private Sub PayeDataTransfer_CSV()
On Error GoTo Err_Flag
  Dim rsChk As New ADODB.Recordset, rsCsv As New ADODB.Recordset
  Dim vFilePath As String, tmpStr As String
  Dim tmpAmount As Double, tmpBasic As Double, vLevy As Double, vLevyBasic As Double
  Dim IsCombined As Boolean
  Dim vSeq As Long
  
    Screen.MousePointer = vbHourglass
    g_SaveFlagNull = False

    CON.BeginTrans
    CON.Execute "truncate table pr_export_csv"
    CON.CommitTrans

    CON.BeginTrans
    Set rsCsv = Nothing
    g_Sql = "select * from pr_export_csv "
    rsCsv.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
    
    'Line 1
    rsCsv.AddNew
    vSeq = vSeq + 1
    rsCsv("n_seq") = vSeq
    rsCsv("c_refno") = "1"
    rsCsv("c_csv") = "MNS,PAYE,V1.0"
    rsCsv.Update
    
    'Line 2
    rsCsv.AddNew
    vSeq = vSeq + 1
    rsCsv("n_seq") = vSeq
    rsCsv("c_refno") = "2"
    rsCsv("c_csv") = "Employer Registration Number,Employer Business Registration Number,Employer Name,Tax Period,Telephone Number,Mobile Number,Name of Declarant,Email Address"
    rsCsv.Update
    
    'Line 3
    Set rsChk = Nothing
    g_Sql = "select c_companyname, c_displayname, c_tan, c_brn, c_contact, c_tel, c_email " & _
            "from pr_company_mst where c_company = '" & Right(Trim(Cmb_Company), 7) & "' "
    rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If rsChk.RecordCount > 0 Then
       vCompShortName = Is_Null(rsChk("c_companyname").Value, False)
       tmpStr = Is_Null(rsChk("c_tan").Value, False) & ","
       tmpStr = tmpStr & Is_Null(rsChk("c_brn").Value, False) & ","
       tmpStr = tmpStr & Is_Null(rsChk("c_displayname").Value, False) & ","
       tmpStr = tmpStr & Right(Format(Val(Txtc_Year), "0000"), 2) & Format(Val(Txtc_Month), "00") & ","
       tmpStr = tmpStr & Is_Null(rsChk("c_tel").Value, False) & ","
       tmpStr = tmpStr & ","
       tmpStr = tmpStr & Is_Null(rsChk("c_contact").Value, False) & ","
       tmpStr = tmpStr & Is_Null(rsChk("c_email").Value, False)
       
       rsCsv.AddNew
       vSeq = vSeq + 1
       rsCsv("n_seq") = vSeq
       rsCsv("c_refno") = "3"
       rsCsv("c_csv") = tmpStr
       rsCsv.Update
    End If
    
    'Line 4
    rsCsv.AddNew
    vSeq = vSeq + 1
    rsCsv("n_seq") = vSeq
    rsCsv("c_refno") = "4"
    rsCsv("c_csv") = "Employee ID,Surname of Employee,Other Names of Employee,Salary Wages Overtime pay Leave pay and other allowances excluding travelling and end of year bonus,PAYE Amount"
    rsCsv.Update
    
    'Line 5 - details
    IsCombined = False
    Set rsChk = Nothing
    If Val(Txtc_Month) = "12" Then
       If MsgBox("Do you want to combined with EOY Bonus?", vbYesNo, "Confirmation") = vbYes Then
          g_Sql = "select a.c_empno, b.c_name, b.c_othername, b.c_nicno, b.c_itno, sum(distinct a.n_paye) n_paye, " & _
                  "sum(case when c.c_saltype='I' then c.n_amount else c.n_amount*-1 end) n_salary " & _
                  "from pr_salary_mst a, pr_emp_mst b, pr_salary_dtl c, pr_paystructure_dtl d " & _
                  "where a.c_empno = b.c_empno and a.c_empno = c.c_empno and a.n_period = c.n_period and c.c_salary = d.c_salary and " & _
                  "d.c_type <> 3 and d.c_salary not in ('SAL0035','SAL0036') and d.c_paye = 'Y' and b.c_rec_sta = 'A' and " & _
                  "a.n_period in (" & vPayPeriod & ", " & vPayPeriod + 1 & ") "
          If Trim(Cmb_Company) <> "" Then
             g_Sql = g_Sql & " and a.c_company = '" & Right(Trim(Cmb_Company), 7) & "' "
          End If
          g_Sql = g_Sql & "group by a.c_empno, b.c_name, b.c_othername, b.c_nicno, b.c_itno " & _
                          "order by a.c_empno "
          IsCombined = True
       End If
    End If

    If Not IsCombined Then
       g_Sql = "select a.c_empno, b.c_name, b.c_othername, b.c_nicno, b.c_itno, max(a.n_paye) n_paye, " & _
               "sum(case when c.c_saltype='I' then c.n_amount else c.n_amount*-1 end) n_salary " & _
               "from pr_salary_mst a, pr_emp_mst b, pr_salary_dtl c, pr_paystructure_dtl d " & _
               "where a.c_empno = b.c_empno and a.c_empno = c.c_empno and a.n_period = c.n_period and c.c_salary = d.c_salary and " & _
               "d.c_type <> 3 and d.c_salary not in ('SAL0035','SAL0036') and d.c_paye = 'Y' and b.c_rec_sta = 'A' and a.n_period = " & vPayPeriod

       If Trim(Cmb_Company) <> "" Then
          g_Sql = g_Sql & " and a.c_company = '" & Right(Trim(Cmb_Company), 7) & "' "
       End If
       g_Sql = g_Sql & " group by a.c_empno, b.c_name, b.c_othername, b.c_nicno, b.c_itno "
       g_Sql = g_Sql & " order by a.c_empno "
    End If

    rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If rsChk.RecordCount = 0 Then
       Screen.MousePointer = vbDefault
       MsgBox "Details are not available to generate", vbInformation, "Information"
       Exit Sub
    End If
    rsChk.MoveFirst

    Do While Not rsChk.EOF
       tmpStr = Is_Null(rsChk("c_nicno").Value, False) & ","
       tmpStr = tmpStr & Left(Is_Null(rsChk("c_name").Value, False), 80) & ","
       tmpStr = tmpStr & Left(Is_Null(rsChk("c_othername").Value, False), 80) & ","
       tmpStr = tmpStr & Trim(Str(Round(Is_Null(rsChk("n_salary").Value, True), 0))) & ","
       tmpStr = tmpStr & Trim(Str(Round(Is_Null(rsChk("n_paye").Value, True), 0)))
              
       rsCsv.AddNew
       vSeq = vSeq + 1
       rsCsv("n_seq") = vSeq
       rsCsv("c_refno") = "5"
       rsCsv("c_csv") = tmpStr
       rsCsv.Update
       
       rsChk.MoveNext
    Loop
    CON.CommitTrans
    
    Screen.MousePointer = vbDefault
    g_SaveFlagNull = False

    MsgBox "PAYE CSV File Format Processed Successfully", vbInformation, "Information"

    If vCompShortName <> "" Then
       comDialog.FileName = "PAYE_" & Trim(vCompShortName) & ".CSV"
    Else
       comDialog.FileName = "PAYE.CSV"
    End If
    comDialog.ShowSave
    vFilePath = comDialog.FileName
    
    Screen.MousePointer = vbHourglass
    Call CsvFileExport_Process(vFilePath)
    Screen.MousePointer = vbDefault

    MsgBox "Transfered to CSV File format Completed Successfully ", vbInformation, "Information"


  Exit Sub

Err_Flag:
   Screen.MousePointer = vbDefault
   CON.RollbackTrans
   g_SaveFlagNull = False
   MsgBox "Error while Generating CSV File Process  - " & Err.Number & Space(5) & Err.Description
End Sub

Private Sub EPZ_Contrib_DataTran_CSV()
On Error GoTo Err_Flag
  Dim rsChk As New ADODB.Recordset, rsCsv As New ADODB.Recordset
  Dim vFilePath As String, tmpStr As String
  Dim vSeq As Long
  
    vSeq = 0
    Screen.MousePointer = vbHourglass
    g_SaveFlagNull = False

    CON.BeginTrans
    CON.Execute "truncate table pr_export_csv"
    CON.CommitTrans
    
    Set rsChk = Nothing
    g_Sql = "select c_companyname from pr_company_mst where c_company = '" & Right(Trim(Cmb_Company), 7) & "' "
    rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If rsChk.RecordCount > 0 Then
       vCompShortName = Is_Null(rsChk("c_companyname").Value, False)
    End If

    Set rsCsv = Nothing
    g_Sql = "select * from pr_export_csv "
    rsCsv.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
    
    'details
    Set rsChk = Nothing
    g_Sql = "select a.c_company, a.c_empno, b.c_nicno, b.c_name, b.c_othername " & _
            "from pr_salary_mst a, pr_emp_mst b " & _
            "where a.c_empno = b.c_empno and a.n_period = " & vPayPeriod & " and b.c_rec_sta = 'A' and a.n_empepz > 0 and a.n_comepz > 0 "

    If Trim(Cmb_Company) <> "" Then
       g_Sql = g_Sql & " and a.c_company = '" & Right(Trim(Cmb_Company), 7) & "' "
    End If
    g_Sql = g_Sql & " order by b.c_name "

    rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If rsChk.RecordCount = 0 Then
       Screen.MousePointer = vbDefault
       MsgBox "Details are not available to generate", vbInformation, "Information"
       Exit Sub
    End If
    rsChk.MoveFirst

    CON.BeginTrans

    Do While Not rsChk.EOF
       tmpStr = Is_Null(rsChk("c_name").Value, False) & " "
       tmpStr = tmpStr & Is_Null(rsChk("c_othername").Value, False) & ","
       tmpStr = tmpStr & Is_Null(rsChk("c_nicno").Value, False)
       
       rsCsv.AddNew
       
       vSeq = vSeq + 1
       rsCsv("n_seq") = vSeq
       rsCsv("c_refno") = "1"
       rsCsv("c_csv") = tmpStr
       
       rsCsv.Update
       rsChk.MoveNext
    Loop
    
    Screen.MousePointer = vbDefault
    g_SaveFlagNull = False
    CON.CommitTrans
    MsgBox "EPZ Contribution Return CSV File Format Processed Successfully", vbInformation, "Information"

    If vCompShortName <> "" Then
       comDialog.FileName = "EPZ_" & Trim(vCompShortName) & ".CSV"
    Else
       comDialog.FileName = "EPZ.CSV"
    End If
    comDialog.ShowSave
    vFilePath = comDialog.FileName
    
    Screen.MousePointer = vbHourglass
    Call CsvFileExport_Process(vFilePath)
    Screen.MousePointer = vbDefault

    MsgBox "Transfered to CSV File format Completed Successfully ", vbInformation, "Information"

  Exit Sub

Err_Flag:
   Screen.MousePointer = vbDefault
   g_SaveFlagNull = False
   CON.RollbackTrans
   MsgBox "Error while Generating CSV File Process - " & Err.Number & Space(5) & Err.Description
End Sub

Private Sub EPZ_Loan_DataTran_CSV()
On Error GoTo Err_Flag
  Dim rsChk As New ADODB.Recordset, rsCsv As New ADODB.Recordset, rs As New ADODB.Recordset
  Dim vFilePath As String, tmpStr As String, tmpRemarks As String
  Dim vSeq As Long
  
    vSeq = 0
    Screen.MousePointer = vbHourglass
    g_SaveFlagNull = False

    CON.BeginTrans
    CON.Execute "truncate table pr_export_csv"
    CON.CommitTrans

    Set rsChk = Nothing
    g_Sql = "select c_companyname from pr_company_mst where c_company = '" & Right(Trim(Cmb_Company), 7) & "' "
    rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If rsChk.RecordCount > 0 Then
       vCompShortName = Is_Null(rsChk("c_companyname").Value, False)
    End If


    Set rsCsv = Nothing
    g_Sql = "select * from pr_export_csv "
    rsCsv.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
    
    'details
    Set rsChk = Nothing
    g_Sql = "Select a.c_empno, max(c.c_name) c_name, max(c.c_othername) c_othername, max(c.c_nicno) c_nicno, " & _
            "max(a.c_epzremarks) c_epzremarks, sum(b.n_paidamount) n_paid " & _
            "from pr_loan_mst a, pr_loan_dtl b, pr_emp_mst c, pr_salary_mst d " & _
            "where a.c_loancode = b.c_loancode and a.c_empno = c.c_empno and c.c_empno = d.c_empno and b.n_period = d.n_period and " & _
            "a.c_type = '0' and c.c_rec_sta = 'A' and b.n_period = " & vPayPeriod

    If Trim(Cmb_Company) <> "" Then
       g_Sql = g_Sql & " and c.c_company = '" & Right(Trim(Cmb_Company), 7) & "' "
    End If
    g_Sql = g_Sql & " group by a.c_empno  order by a.c_name "

    rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If rsChk.RecordCount = 0 Then
       Screen.MousePointer = vbDefault
       MsgBox "Details are not available to generate", vbInformation, "Information"
       Exit Sub
    End If
    rsChk.MoveFirst

    CON.BeginTrans

    Do While Not rsChk.EOF
       tmpRemarks = Is_Null(rsChk("c_epzremarks").Value, False)
       If tmpRemarks = "" Then
          tmpRemarks = "Multipurpose"
       End If
       
       tmpStr = Is_Null(rsChk("c_name").Value, False) & " "
       tmpStr = tmpStr & Is_Null(rsChk("c_othername").Value, False) & ","
       tmpStr = tmpStr & Is_Null(rsChk("c_nicno").Value, False) & ","
       tmpStr = tmpStr & tmpRemarks & ","
       tmpStr = tmpStr & Trim(Str(Is_Null_D(rsChk("n_paid").Value, True, True)))
       
       rsCsv.AddNew
       
       vSeq = vSeq + 1
       rsCsv("n_seq") = vSeq
       rsCsv("c_refno") = "1"
       rsCsv("c_csv") = tmpStr
       
       rsCsv.Update
       rsChk.MoveNext
    Loop
    
    Screen.MousePointer = vbDefault
    g_SaveFlagNull = False
    CON.CommitTrans
    
    MsgBox "EPZ Loan Repayment/Schedule CSV File Format Processed Successfully", vbInformation, "Information"

    If vCompShortName <> "" Then
       comDialog.FileName = "EPZ_Loan_" & Trim(vCompShortName) & ".CSV"
    Else
       comDialog.FileName = "EPZ_Loan.CSV"
    End If
    comDialog.ShowSave
    vFilePath = comDialog.FileName
    
    Screen.MousePointer = vbHourglass
    Call CsvFileExport_Process(vFilePath)
    Screen.MousePointer = vbDefault

    MsgBox "Transfered to CSV File format Completed Successfully ", vbInformation, "Information"

  Exit Sub

Err_Flag:
   Screen.MousePointer = vbDefault
   g_SaveFlagNull = False
   CON.RollbackTrans
   MsgBox "Error while Generating CSV File Process - " & Err.Number & Space(5) & Err.Description
End Sub

Private Sub Emoulment_Statement_Gen()
On Error GoTo Err_Flag
   Dim rsChk As New ADODB.Recordset
   Dim vEmpType As String, vCompany As String, vContact As String, vDesig As String
   Dim vYear As Integer
    
    Set rsChk = Nothing
    g_Sql = "select * from pr_company_mst where c_company = '" & Trim(Right(Trim(Cmb_Company), 7)) & "'"
    rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If rsChk.RecordCount > 0 Then
       vContact = Is_Null(rsChk("c_contact").Value, False)
       vDesig = Is_Null(rsChk("c_desig").Value, False)
    Else
       vContact = "Payroll Admin"
       vDesig = "Payroll Admin"
    End If
    
    
    vYear = Val(Txtc_Year)
    vCompany = Trim(Right(Trim(Cmb_Company), 7))
    If Trim(Cmb_EmpType) <> "" Then
       vEmpType = Trim(Cmb_EmpType)
    Else
       vEmpType = "A"
    End If
    
   
    Screen.MousePointer = vbHourglass

    CON.BeginTrans
    
    g_Sql = "HR_PAYE_YR_RETN_PROC " & vYear & ", '" & vCompany & "', '" & vEmpType & "', '" & vContact & "', '" & vDesig & "', '" & g_UserName & "'"
    CON.Execute g_Sql
    
    CON.CommitTrans
    
    Screen.MousePointer = vbDefault
    MsgBox "Processed Successfully ", vbInformation, "Information"

  Exit Sub

Err_Flag:
   Screen.MousePointer = vbDefault
   g_SaveFlagNull = False
   CON.RollbackTrans
   MsgBox "Error while Generating the Process " & Err.Description
End Sub


Private Sub PAYE_YR_DataTran_CSV()
On Error GoTo Err_Flag
  Dim rsChk As New ADODB.Recordset, rsCsv As New ADODB.Recordset, rsTmp As New ADODB.Recordset
  Dim vFilePath As String, tmpRecord As String, vCompShortName As String
  Dim vSeq As Long
  Dim vEdfAmt As Double
  
    If Trim(Cmb_Company) = "" Then
       MsgBox "Please select company", vbInformation, "Information"
       Exit Sub
    End If
    
    Set rsChk = Nothing
    g_Sql = "select max(n_edfamount) n_edfamount from pr_salary_mst where c_edfcat = 'A' and n_period = " & vPayPeriod
    rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If rsChk.RecordCount > 0 Then
       vEdfAmt = Is_Null(rsChk("n_edfamount").Value, True)
    End If
    
    If vEdfAmt = 0 Then
       Set rsChk = Nothing
       g_Sql = "select n_edfamt from pr_edfmast where c_category = 'A'"
       rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
       If rsChk.RecordCount > 0 Then
          vEdfAmt = Is_Null(rsChk("n_edfamt").Value, True)
       Else
          vEdfAmt = 300000
       End If
    End If
  
    Screen.MousePointer = vbHourglass
    g_SaveFlagNull = False

    CON.BeginTrans
    CON.Execute "truncate table pr_export_csv"
    CON.CommitTrans

    Set rsCsv = Nothing
    g_Sql = "select * from pr_export_csv "
    rsCsv.Open g_Sql, CON, adOpenDynamic, adLockOptimistic

    Set rsChk = Nothing
    g_Sql = "select a.c_empno, a.c_nicno, a.c_itno, a.c_name c_surname, a.c_othername, " & _
            "b.n_salary, b.n_bonus, b.n_rent, b.n_allow, b.n_travel, b.n_travelexp, " & _
            "b.n_othallow, b.n_perexp, b.n_air, b.n_car, b.n_notice, b.n_pension, " & _
            "b.n_except, b.n_eduamount, b.n_intamount, b.n_preamount, b.n_edf+b.n_eduamount+b.n_intamount+b.n_preamount n_edf, " & _
            "b.n_paye, c.c_companyname, c.c_displayname, c.c_tel, c.c_mobile, c.c_contact, c.c_email, c.c_tan, c.c_vat, c.c_brn " & _
            "from pr_emp_mst a, pr_paye_retn_mra_dtl b,  pr_company_mst c " & _
            "where a.c_empno = b.c_empno and b.c_company = c.c_company and a.c_rec_sta = 'A' and b.n_year = " & Val(Txtc_Year)

    If Trim(Cmb_Company) <> "" Then
       g_Sql = g_Sql & " and b.c_company = '" & Right(Trim(Cmb_Company), 7) & "' "
    End If
    g_Sql = g_Sql & " order by a.c_empno "

    rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If rsChk.RecordCount = 0 Then
       Screen.MousePointer = vbDefault
       MsgBox "Details are not available to generate", vbInformation, "Information"
       Exit Sub
    End If
    rsChk.MoveFirst
    
    vCompShortName = Trim(rsChk("c_companyname").Value)
    vSeq = 0
 
    
    CON.BeginTrans
    
    'source
    vSeq = vSeq + 1
    tmpRecord = "MNS,ROEA,V1.0"
    rsCsv.AddNew
    rsCsv("n_seq").Value = vSeq
    rsCsv("c_csv").Value = Trim(tmpRecord)
    rsCsv.Update
    
    
    'header
    vSeq = vSeq + 1
    tmpRecord = "Employer Registration Number,Employer Business Registration Number,Employer Name,Tax Period,Total PAYE Withheld,Telephone Number,Mobile Number,Name of Declarant,Email Address"
    rsCsv.AddNew
    rsCsv("n_seq").Value = vSeq
    rsCsv("c_csv").Value = Trim(tmpRecord)
    rsCsv.Update
    
    
    'header detail
    vSeq = vSeq + 1
    tmpRecord = Trim(rsChk("c_tan").Value)
    tmpRecord = tmpRecord & "," & Trim(rsChk("c_brn").Value)
    tmpRecord = tmpRecord & "," & Trim(rsChk("c_displayname").Value)
    tmpRecord = tmpRecord & "," & Format(Val(Txtc_Year), "0000")
    
    Set rsTmp = Nothing
    g_Sql = "select sum(n_paye) n_paye from pr_paye_retn_mra_dtl where n_year = " & Val(Txtc_Year) & " and c_company = '" & Trim(Right(Trim(Cmb_Company), 7)) & "'"
    rsTmp.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If rsTmp.RecordCount > 0 Then
       tmpRecord = tmpRecord & "," & Trim(Str(Round(rsTmp("n_paye").Value, 0)))
    Else
       tmpRecord = tmpRecord & ",0"
    End If
    tmpRecord = tmpRecord & "," & Trim(rsChk("c_tel").Value)
    tmpRecord = tmpRecord & "," & Trim(rsChk("c_mobile").Value)
    tmpRecord = tmpRecord & "," & Trim(rsChk("c_contact").Value)
    tmpRecord = tmpRecord & "," & Trim(rsChk("c_email").Value)
    rsCsv.AddNew
    rsCsv("n_seq").Value = vSeq
    rsCsv("c_csv").Value = Trim(tmpRecord)
    rsCsv.Update
    
    ' detail header
    vSeq = vSeq + 1
    tmpRecord = "Employee TAN,Employee NID,Surname of Employee,Other Names of Employee,Salary,Entertainment,Transport,Reimbursement,Car Benefits,House Benefits,Tax Benefits,Other Benefits,Lump Sum,Retirement,Emoluments,EDF Declaration,Additional Exemption for Children,Interest Relief on Secured Housing Loan,Relief for Medical insurance premium or contr. for approved provident fund,PAYE Withheld"
    rsCsv.AddNew
    rsCsv("n_seq").Value = vSeq
    rsCsv("c_csv").Value = Trim(tmpRecord)
    rsCsv.Update
    
    ' details
    rsChk.MoveFirst
    Do While Not rsChk.EOF
       tmpRecord = ""
       tmpRecord = Trim(Is_Null(rsChk("c_itno").Value, False))
       If Len(Trim(rsChk("c_nicno").Value)) <> 14 Then
          tmpRecord = tmpRecord & ","
       Else
          tmpRecord = tmpRecord & "," & Trim(rsChk("c_nicno").Value)
       End If
       tmpRecord = tmpRecord & "," & Trim(rsChk("c_othername").Value)
       tmpRecord = tmpRecord & "," & Trim(rsChk("c_surname").Value)
       tmpRecord = tmpRecord & "," & Trim(Str(Round(rsChk("n_salary").Value + rsChk("n_bonus").Value + rsChk("n_othallow").Value + rsChk("n_allow").Value, 0)))
       tmpRecord = tmpRecord & ",0"
       tmpRecord = tmpRecord & "," & Trim(Str(Round(rsChk("n_travel").Value, 0)))
       tmpRecord = tmpRecord & "," & Trim(Str(Round(rsChk("n_travelexp").Value + rsChk("n_air").Value + rsChk("n_perexp").Value, 0)))
       tmpRecord = tmpRecord & "," & Trim(Str(Round(rsChk("n_car").Value, 0)))
       tmpRecord = tmpRecord & "," & Trim(Str(Round(rsChk("n_rent").Value, 0)))
       tmpRecord = tmpRecord & ",0"
       tmpRecord = tmpRecord & ",0"
       tmpRecord = tmpRecord & "," & Trim(Str(Round(rsChk("n_notice").Value, 0)))
       tmpRecord = tmpRecord & "," & Trim(Str(Round(rsChk("n_pension").Value, 0)))
       tmpRecord = tmpRecord & "," & Trim(Str(Round(rsChk("n_except").Value, 0)))
       tmpRecord = tmpRecord & "," & Trim(Str(Round(rsChk("n_edf").Value, 0)))
       tmpRecord = tmpRecord & "," & Trim(Str(Round(rsChk("n_eduamount").Value, 0)))
       tmpRecord = tmpRecord & "," & Trim(Str(Round(rsChk("n_intamount").Value, 0)))
       tmpRecord = tmpRecord & "," & Trim(Str(Round(rsChk("n_preamount").Value, 0)))
       tmpRecord = tmpRecord & "," & Trim(Str(Round(rsChk("n_paye").Value, 0)))

       rsCsv.AddNew
       vSeq = vSeq + 1
       rsCsv("n_seq").Value = vSeq
       rsCsv("c_csv").Value = Trim(tmpRecord)
       rsCsv.Update
       
       rsChk.MoveNext
    Loop
    
    Screen.MousePointer = vbDefault
    g_SaveFlagNull = False
    CON.CommitTrans
    
    If vCompShortName <> "" Then
       comDialog.FileName = "EMOULMENT_" & Trim(vCompShortName) & ".CSV"
    Else
       comDialog.FileName = "EMOULMENT.CSV"
    End If
    comDialog.ShowSave
    vFilePath = comDialog.FileName
    
    Screen.MousePointer = vbHourglass
    Call CsvFileExport_Process(vFilePath)
    Screen.MousePointer = vbDefault

    MsgBox "Transfered to CSV File format Completed Successfully ", vbInformation, "Information"

  Exit Sub

Err_Flag:
   Screen.MousePointer = vbDefault
   g_SaveFlagNull = False
   CON.RollbackTrans
   MsgBox "Error while Generating ACII File Process " & Err.Description & vbCrLf & rsChk("c_empno").Value
End Sub

Private Function AmountToBankChar(ByVal nNetAmount As Double, nNetFloorAmount As Double)
  Dim tmpStr As String
      tmpStr = Trim(Format(nNetFloorAmount, "00000000")) & Trim(Format(IIf(Round((nNetAmount - nNetFloorAmount) * 100, 0) = 100, 99, Round((nNetAmount - nNetFloorAmount) * 100, 0)), "00"))
      AmountToBankChar = Trim(tmpStr)
End Function


Private Sub Save_SCBPaytr_Discp()
  Dim tmpPeriod As String
  
    If Trim(Txtc_Month) = "" Or Trim(Txtc_Year) = "" Then
       MsgBox "Please enter period to generate discrepancies", vbInformation, "Information"
       Exit Sub
    End If
    
    tmpPeriod = Format(Val(Txtc_Month), "00") & Right(Format(Val(Txtc_Year), "0000"), 2)
        
    g_Sql = "truncate table scbpaytr_discp"
    CON.Execute g_Sql
    
    ' Not account no.
    g_Sql = "Insert into scbpaytr_discp (c_empno, c_remarks) " & _
            "select a.c_empno, ' Not a Bank Account Number' " & _
            "from pr_emp_mst a, pr_salary_mst b " & _
            "where a.c_empno = b.c_empno and a.c_rec_sta = 'A' and b.n_period = " & vPayPeriod & " and " & _
            "b.c_paytype = 'BA' and a.c_bankcode = 'B03' and len(a.c_acctno) <= 5"
    CON.Execute g_Sql
    
    ' Barclays accout no.
    g_Sql = "Insert into scbpaytr_discp (c_empno, c_remarks) " & _
            "select a.c_empno, ' Not a Barclays A/c No. Should be 9 Digit.' " & _
            "from pr_emp_mst a, pr_salary_mst b " & _
            "where a.c_empno = b.c_empno and a.c_rec_sta = 'A' and b.n_period = " & vPayPeriod & " and " & _
            "b.c_paytype = 'BA' and a.c_bankcode = 'B03' and a.c_bankcode is null and len(a.c_acctno) <> 9  "
    CON.Execute g_Sql
    
    ' SBM 14 digit
    g_Sql = "Insert into scbpaytr_discp (c_empno, c_remarks) " & _
            "select a.c_empno, ' Not a SBM A/c No. Should be 14 Digit.' " & _
            "from pr_emp_mst a, pr_salary_mst b " & _
            "where a.c_empno = b.c_empno and a.c_rec_sta = 'A' and b.n_period = " & vPayPeriod & " and " & _
            "b.c_paytype = 'BA' and a.c_bankcode = 'B03' and len(a.c_acctno) <> 14 and " & _
            "(a.c_bankcode is null or a.c_bankcode = '11')  "
    CON.Execute g_Sql
    
    ' SBM 10 digit
    g_Sql = "Insert into scbpaytr_discp (c_empno, c_remarks) " & _
            "select a.c_empno, ' Not a Baroda / S.E. Asian A/c No. Should be 10 Digit.' " & _
            "from pr_emp_mst a, pr_salary_mst b " & _
            "where a.c_empno = b.c_empno and a.c_rec_sta = 'A' and b.n_period = " & vPayPeriod & " and " & _
            "b.c_paytype = 'BA' and a.c_bank = 'B03' and len(a.c_acctno) <> 10 and " & _
            "(a.c_bankcode = '02' or a.c_bankcode = '10')  "
    CON.Execute g_Sql
    
    ' SBM 12 digit
    g_Sql = "Insert into scbpaytr_discp (c_empno, c_remarks) " & _
            "select a.c_empno, ' Not a Hong Kong / IOIB A/c No. Should be 12 Digit.' " & _
            "from pr_emp_mst a, pr_salary_mst b " & _
            "where a.c_empno = b.c_empno and a.c_rec_sta = 'A' and b.n_period = " & vPayPeriod & " and " & _
            "b.c_paytype = 'BA' and a.c_bank = 'B03' and len(a.c_acctno) <> 12 and " & _
            "(a.c_bankcode = '07' or a.c_bankcode = '08')  "
    CON.Execute g_Sql
    
    ' SBM 11 digit
    g_Sql = "Insert into scbpaytr_discp (c_empno, c_remarks) " & _
            "select a.c_empno, ' Not a First City A/c No. Should be 11 Digit.' " & _
            "from pr_emp_mst a, pr_salary_mst b " & _
            "where a.c_empno = b.c_empno and a.c_rec_sta = 'A' and b.n_period = " & vPayPeriod & " and " & _
            "b.c_paytype = 'BA' and a.c_bank = 'B03' and len(a.c_acctno) <> 11 and " & _
            "a.c_bankcode = '05' "
    CON.Execute g_Sql
    
    ' SBM 9 digit
    g_Sql = "Insert into scbpaytr_discp (c_empno, c_remarks) " & _
            "select a.c_empno, ' Not a MCB / Barclays A/c No. Should be 9 Digit.' " & _
            "from pr_emp_mst a, pr_salary_mst b " & _
            "where a.c_empno = b.c_empno and a.c_rec_sta = 'A' and b.n_period = " & vPayPeriod & " and " & _
            "b.c_paytype = 'BA' and a.c_bank = 'B03' and len(a.c_acctno) <> 9 and " & _
            "(a.c_bankcode = '03' or a.c_bankcode = '09')  "
    CON.Execute g_Sql
    
    ' Duplicate Account No.
    g_Sql = "Insert into scbpaytr_discp (c_empno, c_remarks) " & _
            "select c_empno, 'Duplicate Account No.' " & _
            "from pr_emp_mst where c_rec_sta = 'A' and c_acctno in " & _
            "(select c_acctno from pr_emp_mst where c_rec_sta = 'A' and c_empno in " & _
            "(select c_empno from pr_salary_mst " & _
            "where c_paytype = 'BA' and n_period = " & vPayPeriod & ") " & _
            "group by c_acctno having count(c_acctno) > 1) "
    CON.Execute g_Sql
    
End Sub

Private Sub GenViewExitButtonStatus(ByVal vPreViewVisible As Boolean)
    If vPreViewVisible = False Then
       Btn_Preview.Visible = False
       Btn_Generate.Left = 4300
       Btn_Exit.Left = 5910
    End If
End Sub

Private Sub PAYE_Yr_Retn_RepFilter()
   vF1 = "": vF2 = "": vF3 = "": vF4 = "": vF5 = ""
   
   vF1 = "{PR_PAYE_RETN_MRA_DTL.N_YEAR} = " & Trim(Txtc_Year)
   
   If Trim(Cmb_Company) <> "" Then
      vF2 = "{PR_PAYE_RETN_MRA_DTL.C_COMPANY} = '" & Trim(Right(Trim(Cmb_Company), 7)) & "'"
   End If
   
   If Trim(Cmb_EmpType) <> "" Then
      vF3 = "{PR_PAYE_RETN_MRA_DTL.C_EMPTYPE} = '" & Trim(Cmb_EmpType) & "'"
   End If
   
   SelFor = ReportFilterOption(vF1, vF2, vF3, vF4, vF5)

End Sub


