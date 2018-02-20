VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Begin VB.Form frm_AddPay_Details 
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
   Begin VB.Frame Frame2 
      Height          =   900
      Left            =   120
      TabIndex        =   30
      Top             =   -30
      Width           =   11970
      Begin VB.TextBox Txtc_No 
         Height          =   300
         Left            =   6480
         TabIndex        =   6
         Top             =   240
         Width           =   1230
      End
      Begin VB.CommandButton Btn_Copy 
         Height          =   720
         Left            =   9795
         Picture         =   "frm_AddPay_Details.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Print 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   3435
         Picture         =   "frm_AddPay_Details.frx":360F
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Delete 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   2715
         Picture         =   "frm_AddPay_Details.frx":6BFD
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Clear 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   1725
         Picture         =   "frm_AddPay_Details.frx":A2A7
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Exit 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   4470
         Picture         =   "frm_AddPay_Details.frx":D917
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Save 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   30
         Picture         =   "frm_AddPay_Details.frx":10F77
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_View 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   1005
         Picture         =   "frm_AddPay_Details.frx":14600
         Style           =   1  'Graphical
         TabIndex        =   1
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
         TabIndex        =   32
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
         TabIndex        =   31
         Top             =   555
         Width           =   2955
      End
   End
   Begin VB.TextBox Txtn_NoEmp 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   1560
      TabIndex        =   17
      Top             =   8730
      Width           =   720
   End
   Begin VB.TextBox Txtn_TotAmount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   6825
      TabIndex        =   18
      Top             =   8730
      Width           =   1545
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
      Height          =   990
      Left            =   120
      TabIndex        =   19
      Top             =   1275
      Width           =   11940
      Begin VB.CommandButton Btn_Display 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   11040
         Picture         =   "frm_AddPay_Details.frx":17CB0
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Display"
         Top             =   180
         Width           =   700
      End
      Begin VB.ComboBox Cmb_Dept 
         Height          =   315
         Left            =   7890
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   540
         Width           =   2850
      End
      Begin VB.ComboBox Cmb_Branch 
         Height          =   315
         Left            =   4275
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   540
         Width           =   1770
      End
      Begin VB.ComboBox Cmb_Company 
         Height          =   315
         Left            =   1110
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   540
         Width           =   1575
      End
      Begin VB.TextBox Txtc_Year 
         Height          =   300
         Left            =   4905
         MaxLength       =   4
         TabIndex        =   10
         Top             =   240
         Width           =   1125
      End
      Begin VB.TextBox Txtc_Month 
         Height          =   300
         Left            =   4275
         MaxLength       =   2
         TabIndex        =   9
         Top             =   240
         Width           =   600
      End
      Begin VB.TextBox Txtc_Remarks 
         Height          =   315
         Left            =   7890
         MaxLength       =   50
         TabIndex        =   11
         Top             =   225
         Width           =   2835
      End
      Begin VB.TextBox Txtc_Code 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   1110
         TabIndex        =   8
         Top             =   240
         Width           =   1560
      End
      Begin VB.Label Label6 
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
         Left            =   7455
         TabIndex        =   29
         Top             =   585
         Width           =   375
      End
      Begin VB.Label Label5 
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
         ForeColor       =   &H00404040&
         Height          =   210
         Left            =   3630
         TabIndex        =   28
         Top             =   585
         Width           =   570
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
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
         Left            =   285
         TabIndex        =   27
         Top             =   585
         Width           =   780
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
         Left            =   7095
         TabIndex        =   24
         Top             =   270
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
         Left            =   3660
         TabIndex        =   23
         Top             =   285
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
         TabIndex        =   22
         Top             =   285
         Width           =   435
      End
   End
   Begin MSComDlg.CommonDialog ComDialog 
      Left            =   12255
      Top             =   255
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin FPUSpreadADO.fpSpread Va_Details 
      Height          =   6345
      Left            =   120
      TabIndex        =   16
      Top             =   2280
      Width           =   11925
      _Version        =   458752
      _ExtentX        =   21034
      _ExtentY        =   11192
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
      SpreadDesigner  =   "frm_AddPay_Details.frx":1B3EE
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "No. of Employee"
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
      Left            =   165
      TabIndex        =   26
      Top             =   8745
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total"
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
      Left            =   6315
      TabIndex        =   25
      Top             =   8760
      Width           =   405
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
      Left            =   10350
      TabIndex        =   21
      Top             =   990
      Width           =   360
   End
   Begin VB.Label lbl_scr_name 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Additional Income / Deductions Details"
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
      TabIndex        =   20
      Top             =   1005
      Width           =   3150
   End
   Begin VB.Shape shp_scr_name 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   120
      Top             =   945
      Width           =   11955
   End
End
Attribute VB_Name = "frm_AddPay_Details"
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
    lbl_scr_name.Caption = "Additional Income / Deduction Entry"
    
    Clear_Spread Va_Details
    Call Spread_Lock
    Call Load_Spread
    Call TGControlProperty(Me)
    Call Spread_Row_Height(Va_Details)
        
    Txtc_Code.Enabled = False
    Txtn_NoEmp.Enabled = False
    Txtn_TotAmount.Enabled = False
End Sub

Private Sub Btn_Exit_Click()
    Unload Me
End Sub

Private Sub Btn_Clear_Click()
    Call CancelButtonClick
End Sub

Private Sub Btn_Delete_Click()
On Error GoTo ErrDel
   Dim RsChk As New ADODB.Recordset
   
    If Trim(Txtc_Code) = "" Then
       Exit Sub
    End If
    
    If ChkPeriodOpen(vPeriod, "W") Then
       If (MsgBox("Are you sure you want to delete ?", vbExclamation + vbYesNo, "Caution") = vbYes) Then
          CON.BeginTrans
          CON.Execute "update pr_addpay_mst set " & GetDelFlag & " where c_code = '" & Trim(Txtc_Code) & "'"
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
  
    SelFor = "{PR_ADDPAY_MST.C_CODE}='" & Trim(Txtc_Code) & "'"
    Call Print_Rpt(SelFor, "Pr_AddIncDed_Doc.rpt")
    Mdi_Ta_HrPay.CRY1.Action = 1
  
  Exit Sub

Err_Print:
    MsgBox "Error while Generating - " & Err.Description, vbInformation, "Information"
End Sub

Private Sub Btn_View_Click()
  Dim Search As New Search.MyClass, SerVar

    Search.Query = "select a.n_period Period, a.c_remarks Remarks, a.c_code Code " & _
                   "from pr_addpay_mst a, pr_payperiod_dtl b " & _
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
        
        Save_Pr_AddPay_Mst
        Save_Pr_AddPay_Dtl
     
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
  Dim i As Integer
  
  If Trim(Txtc_Month) = "" Then
     MsgBox "Period should not be empty", vbInformation, "Information"
     Txtc_Month.SetFocus
     Exit Function
  ElseIf Trim(Txtc_Year) = "" Then
     MsgBox "Period should not be empty", vbInformation, "Information"
     Txtc_Year.SetFocus
     Exit Function
  ElseIf IsColTwoDupValues(Va_Details, 1, 3) > 0 Then
     MsgBox "Duplicate entry found.", vbInformation, "Information"
     Va_Details.SetFocus
     Exit Function
  ElseIf Not ChkPeriodOpen(vPeriod, "W") Then
     Exit Function
  End If
  
  For i = 1 To Va_Details.DataRowCnt
      Va_Details.Row = i
      Va_Details.Col = 4
         If Val(Va_Details.Text) > 0 Then
            Va_Details.Col = 1
               If Trim(Va_Details.Text) = "" Then
                  MsgBox "Employee No. Should not be Empty", vbInformation, "Information"
                  Va_Details.SetFocus
                  Exit Function
               End If
            Va_Details.Col = 3
               If Trim(Va_Details.Text) = "" Then
                  MsgBox "Income / Deduction Type Should not be Empty", vbInformation, "Information"
                  Va_Details.SetFocus
                  Exit Function
               End If
         End If
  Next i
  
  ChkSave = True
End Function

Private Sub Save_Pr_AddPay_Mst()
    
       Set rs = Nothing
       g_Sql = "Select * from pr_addpay_mst where c_code = '" & Trim(Txtc_Code) & "' and c_rec_sta='A'"
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

Private Sub Save_Pr_AddPay_Dtl()
Dim i As Long
Dim tmpEmpNo, tmpSalary As String

       Set rs = Nothing
       g_Sql = "delete from pr_addpay_dtl where c_code = '" & Trim(Txtc_Code) & "'"
       CON.Execute (g_Sql)
       
       g_Sql = "Select * from pr_addpay_dtl where 1=2"
       rs.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
     
       For i = 1 To Va_Details.DataRowCnt
           Va_Details.Row = i
           Va_Details.Col = 1
              tmpEmpNo = Trim(Va_Details.Text)
           Va_Details.Col = 3
              tmpSalary = Trim(Va_Details.Text)
              
              If Trim(tmpEmpNo) <> "" And Trim(tmpSalary) <> "" Then
                 rs.AddNew
                 rs("c_code").Value = Is_Null(Txtc_Code, False)
                 rs("n_period").Value = Is_Null(vPeriod, True)
                 Va_Details.Col = 1
                    rs("c_empno").Value = Is_Null(Va_Details.Text, False)
                 Va_Details.Col = 3
                    rs("c_salary").Value = Is_Null(Right(Trim(Va_Details.Text), 7), False)
                 Va_Details.Col = 4
                    rs("n_amount").Value = Is_Null_D(Va_Details.Text, True)
                 Va_Details.Col = 5
                    rs("c_remarks").Value = Is_Null_D(Va_Details.Text, False)
                 rs.Update
              End If
       Next i
End Sub

Private Sub Start_Generate_New()
  Dim MaxNo As ADODB.Recordset
  Dim vCode As String
  
  vCode = Trim(Str(vPeriod - 200000))
  g_Sql = "Select max(substring(c_code,6,5)) from pr_addpay_mst where c_code like '" & vCode & "%'"
  Set MaxNo = CON.Execute(g_Sql)
  Txtc_Code = vCode + "/" + Format(Is_Null(MaxNo(0).Value, True) + 1, "00000")
End Sub

Private Sub Display_Records()
On Error GoTo Err_Display
  Dim DyDisp As New ADODB.Recordset
  Dim i, j As Long
  Dim vType As String
  
  Set DyDisp = Nothing
  g_Sql = "select * from pr_addpay_mst where c_code = '" & Trim(Txtc_Code) & "' and c_rec_sta='A'"
  DyDisp.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
  If DyDisp.RecordCount = 0 Then
     Exit Sub
  End If
  
  Txtc_Code = Is_Null(DyDisp("c_code").Value, False)
  vPeriod = Is_Null(DyDisp("n_period").Value, True)
  Txtc_Month = Right(Trim(Str(vPeriod)), 2)
  Txtc_Year = Left(Trim(Str(vPeriod)), 4)
  Txtc_Remarks = Is_Null(DyDisp("c_remarks").Value, False)
  
  ' // Details
   Set DyDisp = Nothing
   g_Sql = "select a.c_empno, b.c_name, b.c_othername, a.c_salary, a.n_amount, a.c_remarks, c.c_payname " & _
           "from pr_addpay_dtl a, pr_emp_mst b, pr_paystructure_dtl c " & _
           "where a.c_empno = b.c_empno and a.c_salary = c.c_salary and a.c_code = '" & Trim(Txtc_Code) & "' " & _
           "order by a.c_empno, c.c_payname "
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
             Va_Details.Text = Is_Null(DyDisp("c_payname").Value & Space(150) & DyDisp("c_salary").Value, False)
          Va_Details.Col = 4
             Va_Details.Text = Is_Null(DyDisp("n_amount").Value, True)
          Va_Details.Col = 5
             Va_Details.Text = Is_Null(DyDisp("c_remarks").Value, False)
          DyDisp.MoveNext
      Next i
      Call Calculate_Total
      Call Calculate_NoEmployee
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
End Sub

Private Function Check_Existing_Emp() As Boolean
  Dim RsChk As New ADODB.Recordset
            
     Set RsChk = Nothing
     Va_Details.Row = Va_Details.ActiveRow
     Va_Details.Col = 1
        g_Sql = "select c_empno, c_name, c_othername from pr_emp_mst " & _
                "where c_rec_sta = 'A' and (d_dol is null or d_dol >= '" & Is_Date(vPayPeriodFrom, "S") & "') and " & _
                "c_empno = '" & Trim(Va_Details.Text) & "'"
        RsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
        If RsChk.RecordCount > 0 Then
           Va_Details.Col = 1
              Va_Details.Text = Is_Null(RsChk("c_empno").Value, False)
           Va_Details.Col = 2
              Va_Details.Text = Proper(Is_Null(RsChk("c_name").Value, False)) & " " & Proper(Is_Null(RsChk("c_othername").Value, False))
           
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
        If i = 2 Or i = 3 Then
           Va_Details.Lock = True
        Else
           Va_Details.Lock = False
        End If
    Next i
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
                    "from pr_addpay_mst  where c_rec_sta='A' " & tmpFilter
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
    
    If Trim(Txtc_Remarks) <> "" Then
      ' Call SpreadScrollAllow(Me, True)
    End If

End Sub

Private Function Check_Existing_Salary() As Boolean
  Dim RsChk As New ADODB.Recordset
  Dim tmpEmpNo, tmpSalary, tmpOldSalary As String
     
     Va_Details.Row = Va_Details.ActiveRow
     Va_Details.Col = 1
        tmpEmpNo = Trim(Va_Details.Text)
     Va_Details.Col = 3
        tmpSalary = Right(Trim(Va_Details.Text), 7)
     Va_Details.Col = 5
        tmpOldSalary = Trim(Va_Details.Text)
        
     If Trim(tmpEmpNo) <> "" Then
        Set RsChk = Nothing
        g_Sql = "select * from pr_addpay_mst a, pr_addpay_dtl b " & _
                "where a.c_code = b.c_code and a.c_rec_sta ='A' and b.n_period = " & vPeriod & " and " & _
                "b.c_empno = '" & Trim(tmpEmpNo) & "' and b.c_salary = '" & Trim(tmpSalary) & "' and " & _
                "a.c_code <> '" & Trim(Txtc_Code) & "' "
        RsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
        If RsChk.RecordCount > 0 Then
           Va_Details.Col = 3
           Va_Details.Text = ""
           MsgBox "Entries are already available for this employee " & vbCrLf & _
                  "Code - " & Is_Null(RsChk("c_code").Value, False), vbInformation, "Information"
           Check_Existing_Salary = False
           Exit Function
        End If
     End If
     Check_Existing_Salary = True
End Function

Private Sub Va_Details_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
   If BlockCol = 3 Or BlockCol = 4 Then
      Call SpreadBlockCopy(Va_Details, BlockCol, BlockRow, BlockCol2, BlockRow2)
   End If
End Sub

Private Sub Va_Details_Change(ByVal Col As Long, ByVal Row As Long)
  
  If Col = 1 Or Col = 2 Or Col = 3 Then
     Va_Details.Row = Row
     Va_Details.Col = 1
        If Trim(Va_Details.Text) <> "" Then
           Call IsColTwoDupValues(Va_Details, 1, 3)
        End If
  End If
End Sub

Private Sub Va_Details_Click(ByVal Col As Long, ByVal Row As Long)
    Call SpreadColSort(Va_Details, Col, Row)
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
                    "from pr_emp_mst where d_dol is null"
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
     Search.Query = "select c_payname PaymentType, c_salary Code from pr_paystructure_dtl where c_syscal = 'N' "
     Search.CheckFields = "Code, PaymentType "
     Search.ReturnField = "Code, PaymentType "
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
  
  If Col = 1 Or Col = 2 Then
     Va_Details.Row = Row
     Va_Details.Col = 1
        If Trim(Va_Details.Text) <> "" Then
           If vPeriod > 0 Then
              If Not Check_Existing_Emp Then
                 Cancel = True
              End If
           End If
        End If
        If Col = 1 Then
           Call Calculate_NoEmployee
        End If
  
  ElseIf Col = 3 Then
     Va_Details.Row = Row
     Va_Details.Col = 3
        If Trim(Va_Details.Text) <> "" Then
           If Not Check_Existing_Salary Then
              Cancel = True
           End If
        End If
  ElseIf Col = 4 Then
     Call Calculate_Total
  End If
End Sub

Private Sub Calculate_Total()
  Dim i As Long
      
      Txtn_TotAmount = 0
      For i = 1 To Va_Details.DataRowCnt
          Va_Details.Row = i
          Va_Details.Col = 4
             Txtn_TotAmount = Is_Null_D(Txtn_TotAmount, True) + Is_Null_D(Va_Details.Text, True)
      Next i
      Txtn_TotAmount = Format(Is_Null_D(Txtn_TotAmount, True), g_Nformat)
End Sub

Private Sub Calculate_NoEmployee()
  Dim i As Long
      
      Txtn_NoEmp = 0
      For i = 1 To Va_Details.DataRowCnt
          Va_Details.Row = i
          Va_Details.Col = 1
             If Trim(Va_Details.Text) <> "" Then
                Txtn_NoEmp = Val(Txtn_NoEmp) + 1
             End If
      Next i
      Txtn_NoEmp = Format(Val(Txtn_NoEmp), g_Nformat0)
End Sub

Private Sub Disable_Controls()
    Txtn_TotAmount.Enabled = False
    Txtn_NoEmp.Enabled = False
    Txtc_No.Enabled = False
End Sub

Private Sub Load_Spread()
   Call LoadComboCompany(Me)
   Call LoadComboBranch(Me)
   Call LoadComboDept(Me)
End Sub

Private Sub Btn_Display_Click()
  Dim RsChk As New ADODB.Recordset
  Dim i As Integer, j As Integer
  Dim vF1, vF2, vF3 As String
  
      vF1 = Trim(Right(Trim(Cmb_Company), 7))
      vF2 = Trim(Cmb_Branch)
      vF3 = Trim(Cmb_Dept)
      
      If ReportFilterOption(vF1, vF2, vF3) = "" Then
         MsgBox "Please Select Filter to Display Employee Details", vbInformation, "Information"
         Cmb_Company.SetFocus
         Exit Sub
     End If
      
      If Va_Details.DataRowCnt > 0 Then
         If MsgBox("Do you want to added to the entries?", vbYesNo + vbDefaultButton2, "Confirmation") = vbNo Then
            Va_Details.SetFocus
            Exit Sub
         End If
      End If
      
      g_Sql = "select c_empno, c_name, c_othername from pr_emp_mst where d_dol is null "

      If Trim(Cmb_Company) <> "" Then
         g_Sql = g_Sql & " and c_company = '" & Trim(Right(Trim(Cmb_Company), 7)) & "'"
      End If
      
      If Trim(Cmb_Branch) <> "" Then
         g_Sql = g_Sql & " and c_branch = '" & Trim(Cmb_Branch) & "'"
      End If
     
      If Trim(Cmb_Dept) <> "" Then
         g_Sql = g_Sql & " and c_dept = '" & Trim(Cmb_Dept) & "'"
      End If
     
      g_Sql = g_Sql & " Order by c_empno"
      
      Set RsChk = Nothing
      RsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
      Va_Details.MaxRows = RsChk.RecordCount + 25
      If RsChk.RecordCount > 0 Then
         j = Va_Details.DataRowCnt
         For i = 1 To RsChk.RecordCount
                j = j + 1
                Va_Details.Row = j
                Va_Details.Col = 1
                Va_Details.Text = Is_Null(RsChk("c_empno").Value, False)
                Va_Details.Col = 2
                Va_Details.Text = Proper(Is_Null(RsChk("c_name").Value, False) & " " & Is_Null(RsChk("c_othername").Value, False))
                RsChk.MoveNext
         Next i
      End If
      Call Spread_Row_Height(Va_Details)
End Sub

Private Sub Assign_PayPeriodDate()
  Dim RsChk As New ADODB.Recordset

    Set RsChk = Nothing
    g_Sql = "select d_fromdate, d_todate from pr_payperiod_dtl where n_period = " & vPeriod & " and c_type = 'W' "
    RsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If RsChk.RecordCount > 0 Then
       vPayPeriodFrom = RsChk("d_fromdate").Value
       vPayPeriodTo = RsChk("d_todate").Value
    End If
End Sub

Private Sub Btn_Copy_Click()
  Dim Search As New Search.MyClass, SerVar
  Dim tmpFilter As String
  
    tmpFilter = InputBox("Please input period to view", "Filter Option")
    If Trim(tmpFilter) <> "" Then
       tmpFilter = " and n_period like '" & Trim(tmpFilter) & "%'"
    End If
    
    Search.Query = "select n_period Period, c_remarks Remarks, c_code Code " & _
                   "from pr_addpay_mst where c_rec_sta='A' " & tmpFilter
    Search.CheckFields = "Code"
    Search.ReturnField = "Code"
    SerVar = Search.Search(, , CON)
    If Len(Search.col1) <> 0 Then
        Call CancelButtonClick
        Txtc_Code = Search.col1
        Call Display_Records
        
        Txtc_Code = ""
        Txtc_Month = ""
        Txtc_Year = ""
        vPeriod = 0
        Txtc_Remarks = ""
        
        Txtc_Year = Format(Year(g_CurrentDate), "0000")
        Txtc_Month.SetFocus
    End If
End Sub


