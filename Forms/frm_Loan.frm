VERSION 5.00
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Object = "{D3400A36-5DD1-4308-8669-3D13D3C27560}#1.0#0"; "CS_DateControl.ocx"
Begin VB.Form frm_Loan 
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
      TabIndex        =   45
      Top             =   -30
      Width           =   11505
      Begin VB.CommandButton Btn_Print 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   3435
         Picture         =   "frm_Loan.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Delete 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   2715
         Picture         =   "frm_Loan.frx":35EE
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Clear 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   1725
         Picture         =   "frm_Loan.frx":6C98
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Exit 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   4470
         Picture         =   "frm_Loan.frx":A308
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Save 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   30
         Picture         =   "frm_Loan.frx":D968
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_View 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   1005
         Picture         =   "frm_Loan.frx":10FF1
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   700
      End
      Begin VB.Frame Frm_EmpStatus 
         Caption         =   "Loan Status"
         ForeColor       =   &H00C00000&
         Height          =   720
         Left            =   6045
         TabIndex        =   46
         Top             =   105
         Width           =   2925
         Begin VB.OptionButton Opt_All 
            Caption         =   "All"
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
            Left            =   1965
            TabIndex        =   7
            Top             =   330
            Width           =   735
         End
         Begin VB.OptionButton Opt_OS 
            Caption         =   "Outstanding"
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
            TabIndex        =   6
            Top             =   345
            Value           =   -1  'True
            Width           =   1320
         End
      End
   End
   Begin FPUSpreadADO.fpSpread Va_Loan 
      Height          =   4965
      Left            =   6480
      TabIndex        =   25
      Top             =   1380
      Width           =   5130
      _Version        =   458752
      _ExtentX        =   9049
      _ExtentY        =   8758
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
      MaxCols         =   8
      MaxRows         =   50
      ProcessTab      =   -1  'True
      SpreadDesigner  =   "frm_Loan.frx":146A1
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
      Height          =   3945
      Left            =   120
      TabIndex        =   26
      Top             =   1275
      Width           =   6285
      Begin VB.ComboBox Cmb_EPZRemarks 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   3450
         Width           =   4575
      End
      Begin VB.TextBox Txtc_Year 
         Height          =   300
         Left            =   2115
         MaxLength       =   4
         TabIndex        =   15
         Top             =   1710
         Width           =   840
      End
      Begin VB.TextBox Txtc_Month 
         Height          =   300
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   14
         Top             =   1710
         Width           =   525
      End
      Begin VB.ComboBox Cmb_Type 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1095
         Width           =   4575
      End
      Begin VB.TextBox Txtn_Balance 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   4740
         TabIndex        =   19
         Top             =   2595
         Width           =   1380
      End
      Begin VB.TextBox Txtn_PaidAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   4740
         TabIndex        =   18
         Top             =   2295
         Width           =   1380
      End
      Begin VB.TextBox Txtn_PaidInstall 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   1560
         TabIndex        =   17
         Top             =   2295
         Width           =   1380
      End
      Begin VB.TextBox Txtc_LoanCode 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   1560
         TabIndex        =   8
         Top             =   240
         Width           =   1125
      End
      Begin VB.TextBox Txtn_NoInstall 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   4710
         MaxLength       =   3
         TabIndex        =   13
         Top             =   1425
         Width           =   1410
      End
      Begin VB.TextBox Txtn_LoanAmount 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   12
         Top             =   1410
         Width           =   1410
      End
      Begin VB.TextBox Txtc_Remarks 
         Height          =   300
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   20
         Top             =   3150
         Width           =   4575
      End
      Begin VB.TextBox Txtc_EmployeeName 
         Height          =   300
         Left            =   1560
         TabIndex        =   10
         Top             =   540
         Width           =   4560
      End
      Begin CS_DateControl.DateControl Dtp_Date 
         Height          =   345
         Left            =   4245
         TabIndex        =   9
         Top             =   225
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   609
      End
      Begin CS_DateControl.DateControl Dtp_PaidOn 
         Height          =   345
         Left            =   4710
         TabIndex        =   16
         Top             =   1725
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   609
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "EPZ Remarks"
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
         TabIndex        =   44
         Top             =   3510
         Width           =   1095
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Loan Paid On"
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
         Left            =   3570
         TabIndex        =   40
         Top             =   1770
         Width           =   1065
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Loan Type"
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
         Left            =   615
         TabIndex        =   39
         Top             =   1140
         Width           =   855
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000005&
         X1              =   15
         X2              =   7200
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Amount Balance"
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
         Left            =   3330
         TabIndex        =   38
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Paid Amount"
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
         Left            =   3615
         TabIndex        =   37
         Top             =   2340
         Width           =   1050
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Installment Paid"
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
         Left            =   150
         TabIndex        =   36
         Top             =   2340
         Width           =   1320
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Loan Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   210
         Left            =   585
         TabIndex        =   35
         Top             =   285
         Width           =   885
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   15
         X2              =   7200
         Y1              =   3015
         Y2              =   3015
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Deduction Starts"
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
         Left            =   105
         TabIndex        =   34
         Top             =   1755
         Width           =   1365
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Date"
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
         Left            =   3795
         TabIndex        =   33
         Top             =   285
         Width           =   360
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "No. of Installment"
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
         Left            =   3195
         TabIndex        =   32
         Top             =   1485
         Width           =   1440
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         X1              =   15
         X2              =   7210
         Y1              =   930
         Y2              =   930
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
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
         Left            =   720
         TabIndex        =   31
         Top             =   3210
         Width           =   750
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Loan Amount"
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
         Left            =   360
         TabIndex        =   30
         Top             =   1455
         Width           =   1110
      End
      Begin VB.Label lbl_add 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Employee Name"
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
         Left            =   150
         TabIndex        =   29
         Top             =   585
         Width           =   1320
      End
   End
   Begin VB.Frame Fme_AdjLoan 
      Caption         =   "Load Paid by Cash / Loan Adjustments"
      ForeColor       =   &H00C000C0&
      Height          =   1065
      Left            =   120
      TabIndex        =   41
      Top             =   5280
      Width           =   6285
      Begin VB.TextBox Txtc_AdjRemarks 
         Height          =   300
         Left            =   1035
         MaxLength       =   50
         TabIndex        =   24
         Top             =   615
         Width           =   3960
      End
      Begin VB.TextBox Txtc_AdjMonth 
         Height          =   300
         Left            =   1035
         MaxLength       =   2
         TabIndex        =   22
         Top             =   315
         Width           =   510
      End
      Begin VB.TextBox Txtc_AdjYear 
         Height          =   300
         Left            =   1530
         MaxLength       =   4
         TabIndex        =   23
         Top             =   315
         Width           =   735
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
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
         Left            =   180
         TabIndex        =   43
         Top             =   660
         Width           =   750
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
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
         Left            =   390
         TabIndex        =   42
         Top             =   345
         Width           =   540
      End
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
      Left            =   10155
      TabIndex        =   28
      Top             =   960
      Width           =   360
   End
   Begin VB.Label lbl_scr_name 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Employee Loan/Advance Details"
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
      TabIndex        =   27
      Top             =   990
      Width           =   2610
   End
   Begin VB.Shape shp_scr_name 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   360
      Left            =   120
      Top             =   915
      Width           =   11490
   End
End
Attribute VB_Name = "frm_Loan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As New ADODB.Recordset

Private Sub Form_Load()
    lbl_date.Caption = Format(Date, "dd-mmm-yyyy")
    Dtp_Date.Text = Is_Date(Date, "D")
      
    Call Load_Combo
    Call Spread_Lock
    Call TGControlProperty(Me)
    Call Spread_Row_Height(Va_Loan)
    
    Txtc_LoanCode.Enabled = False
    Txtn_PaidAmount.Enabled = False
    Txtn_PaidInstall.Enabled = False
    Txtn_Balance.Enabled = False
End Sub

Private Sub Btn_Clear_Click()
    Clear_Controls Me
    Clear_Spread Va_Loan
    Cmb_Type.ListIndex = 0
    Dtp_Date.SetFocus

    Txtc_LoanCode.Enabled = False
    Txtn_PaidAmount.Enabled = False
    Txtn_PaidInstall.Enabled = False
    Txtn_Balance.Enabled = False
End Sub

Private Sub Btn_Exit_Click()
    Unload Me
End Sub

Private Sub Btn_Delete_Click()
On Error GoTo ErrDel
  Dim RsChk As New ADODB.Recordset
  
    If Trim(Txtc_LoanCode) = "" Then
       Exit Sub
    End If

    Set RsChk = Nothing
    g_Sql = "select n_loanpaid from pr_loan_mst where c_loancode = '" & Trim(Txtc_LoanCode) & "'"
    RsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If RsChk.RecordCount > 0 Then
       If Is_Null(RsChk("n_loanpaid").Value, True) > 0 Then
          MsgBox "Transaction over. No access to Delete", vbInformation, "Information"
          Exit Sub
       End If
    End If

    If (MsgBox("Are you sure you want to delete ?", vbExclamation + vbYesNo, "Caution") = vbYes) Then
       CON.BeginTrans
       CON.Execute "update pr_loan_mst set " & GetDelFlag & " where c_loancode='" & Trim(Txtc_LoanCode) & "'"
       CON.CommitTrans
    End If
    Call Btn_Clear_Click
  Exit Sub

ErrDel:
    MsgBox "Error while Deleting - " + Err.Description, vbCritical, "Critical"
    CON.RollbackTrans
End Sub

Private Sub Btn_Print_Click()
On Error GoTo Err_Print
  Dim SelFor As String
   
'   SelFor = "{MADA_PR_LOAN_MST.C_LOANCODE}='" & Trim(Txtc_LoanCode) & "'"
'   Call Print_Rpt(SelFor, "Pr_Loan_List.rpt")
'   Mdi_Report.CRY1.Action = 1
  
  Exit Sub

Err_Print:
    MsgBox "Error while Generating - " + Err.Description, vbInformation, "Information"
End Sub

Private Sub Cmb_Type_Click()
    If Cmb_Type.ListIndex = 0 Then
       Cmb_EPZRemarks.ListIndex = 0
    End If
End Sub

Private Sub Btn_View_Click()
  Dim Search As New Search.MyClass, SerVar
  Dim tmpFilter As String
  
    If Opt_OS.Value = True Then
       tmpFilter = tmpFilter & " and a.n_loanamount <> a.n_loanpaid"
    End If
  
    Search.Query = "select a.c_loancode Code, a.c_empno EmpNo, b.c_name Name, b.c_othername OtherName, " & _
                   "a.n_loanamount LoanAmount, n_noinstall Installment, a.n_loanpaid LoanPaid " & _
                   "from pr_loan_mst a, pr_emp_mst b " & _
                   "where a.c_empno = b.c_empno and a.c_rec_sta = 'A' "
    Search.CheckFields = "Code"
    Search.ReturnField = "Code"
    SerVar = Search.Search(, , CON)
    If Len(Search.col1) <> 0 Then
       Txtc_LoanCode = Search.col1
       Call Display_Records
        
        Va_Loan.Row = -1
        Va_Loan.Col = -1
           Va_Loan.Lock = True
        Va_Loan.Enabled = True
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
        Save_Pr_Loan_Mst
        Save_Pr_Loan_Dtl
        Save_Pr_Loan_Adj_Dtl
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
 Dim i, tmpQty As Double
  
  If Not IsDate(Dtp_Date.Text) Then
     MsgBox "Date Should not be empty", vbInformation, "Information"
     Dtp_Date.SetFocus
     Exit Function
  ElseIf Trim(Txtc_EmployeeName) = "" Then
     MsgBox "Employee Name Should not be empty", vbInformation, "Information"
     Txtc_EmployeeName.SetFocus
     Exit Function
  ElseIf Trim(Cmb_Type) = "" Then
     MsgBox "Loan type Should not be empty", vbInformation, "Information"
     Cmb_Type.SetFocus
     Exit Function
  ElseIf Is_Null_D(Txtn_LoanAmount, True) <= 0 Then
     MsgBox "Loan Amount should not be Zero", vbInformation, "Information"
     Txtn_LoanAmount.SetFocus
     Exit Function
  ElseIf Trim(Txtc_Month) = "" Or Trim(Txtc_Year) = "" Then
     MsgBox "Deduction Starts Period should not be empty", vbInformation, "Information"
     Txtc_Month.SetFocus
     Exit Function
  ElseIf Not IsDate(Dtp_PaidOn.Text) Then
     MsgBox "Loan Paid On should not be empty.", vbInformation, "Information"
     Dtp_PaidOn.SetFocus
     Exit Function
  End If
  
  If Trim(Txtc_AdjMonth) <> "" And Trim(Txtc_AdjYear) <> "" Then
     If Trim(Txtc_AdjRemarks) = "" Then
        MsgBox "Loan Adjustment Remarks should not be Empty", vbInformation, "Information"
        Txtc_AdjRemarks.SetFocus
        Exit Function
     End If
  End If
    
  Va_Loan.Row = 1
  Va_Loan.Col = 1
     If Val(Txtc_Month) <> Val(Va_Loan.Text) Then
        MsgBox "Loan Deduction Starts is not tallied. Please check your entry", vbInformation, "Information"
        Txtc_Month.SetFocus
        Exit Function
     End If
  Va_Loan.Row = 1
  Va_Loan.Col = 2
     If Val(Txtc_Year) <> Val(Va_Loan.Text) Then
        MsgBox "Loan Deduction Starts is not tallied. Please check your entry", vbInformation, "Information"
        Txtc_Year.SetFocus
        Exit Function
     End If
    
  tmpQty = 0
  For i = 1 To Va_Loan.DataRowCnt
      Va_Loan.Row = i
      Va_Loan.Col = 3
         tmpQty = tmpQty + Is_Null_D(Va_Loan.Text, True)
  Next i
  If Is_Null_D(Txtn_LoanAmount, True, True) <> Is_Null_D(tmpQty, True, True) Then
     MsgBox "Loan Amount is not equal to installments", vbInformation, "Information"
     Txtn_LoanAmount.SetFocus
     Exit Function
  End If
  
  If Right(Trim(Cmb_Type), 1) = 1 And Cmb_EPZRemarks = "" Then
    MsgBox "Select EPZ Remarks"
    Cmb_EPZRemarks.SetFocus
    Exit Function
  End If
    
  ChkSave = True
  
End Function

Private Sub Save_Pr_Loan_Mst()
    Set rs = Nothing
    g_Sql = "Select * from pr_loan_mst where c_loancode = '" & Trim(Txtc_LoanCode) & "'"
    rs.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
    
    If rs.RecordCount = 0 Then
       rs.AddNew
       Call Start_Generate_New
       rs("d_created").Value = GetDateTime
       rs("c_usr_id").Value = g_UserName
    Else
       rs("d_modified").Value = GetDateTime
       rs("c_musr_id").Value = g_UserName
    End If
       
       rs("c_loancode").Value = Is_Null(Txtc_LoanCode, False)
       rs("d_date").Value = IIf(IsDate(Dtp_Date.Text), Format(Dtp_Date.Text, "dd/mm/yyyy"), Null)
       rs("c_empno").Value = Is_Null(Right(Trim(Txtc_EmployeeName), 6), False)
       rs("c_type").Value = Is_Null(Right(Trim(Cmb_Type), 1), False)
       
       rs("n_loanamount").Value = Is_Null_D(Txtn_LoanAmount, True)
       rs("n_noinstall").Value = Is_Null(Txtn_NoInstall, True)
       rs("d_deductstarts").Value = Is_Date("25-" & Format(Txtc_Month, "00") & "-" & Format(Txtc_Year, "0000"), "S")
       rs("d_paidon").Value = Is_Date(Dtp_PaidOn.Text, "S")
       rs("c_remarks").Value = Is_Null(Txtc_Remarks, False)
       rs("c_epzremarks").Value = Is_Null(Cmb_EPZRemarks, False)
       rs("c_rec_sta").Value = "A"

       rs.Update
End Sub

Private Sub Save_Pr_Loan_Dtl()
Dim i As Long
Dim tmpYear As Integer, tmpMonth As Integer

       g_Sql = "delete from pr_loan_dtl where c_loancode = '" & Trim(Txtc_LoanCode) & "'"
       CON.Execute (g_Sql)
       
       Set rs = Nothing
       g_Sql = "Select * from pr_loan_dtl where 1=2"
       rs.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
     
       For i = 1 To Va_Loan.DataRowCnt
           rs.AddNew
           Va_Loan.Row = i
           rs("c_loancode").Value = Is_Null(Txtc_LoanCode, False)
           Va_Loan.Col = 1
              tmpMonth = Is_Null_D(Va_Loan.Text, True)
           Va_Loan.Col = 2
              tmpYear = Is_Null_D(Va_Loan.Text, True)
              rs("n_period").Value = Is_Null(Format(tmpYear, "0000") & Format(tmpMonth, "00"), True)
           Va_Loan.Col = 3
              rs("n_amount").Value = Is_Null_D(Va_Loan.Text, True)
           Va_Loan.Col = 4
              rs("n_paidamount").Value = Is_Null_D(Va_Loan.Text, True)
           Va_Loan.Col = 5
              rs("c_flag").Value = Is_Null(Va_Loan.Text, False)
           Va_Loan.Col = 6
              rs("c_adjremarks").Value = Is_Null(Va_Loan.Text, False)
           Va_Loan.Col = 7
              rs("c_adjby").Value = Is_Null(Va_Loan.Text, False)
           Va_Loan.Col = 8
              rs("d_adjon").Value = Is_Date(Va_Loan.Text, "S")
              
           rs("d_recdon").Value = Is_Date("25-" & Format(tmpMonth, "00") & "-" & Format(tmpYear, "0000"), "S")
           
           If i = 1 Then
              g_Sql = "update pr_loan_mst set n_startperiod = " & Is_Null(Format(tmpYear, "0000") & Format(tmpMonth, "0000"), True) & " " & _
                      "where c_loancode = '" & Is_Null(Txtc_LoanCode, False) & "'"
              CON.Execute g_Sql
           End If
           
        rs.Update
       Next i
End Sub

Private Sub Save_Pr_Loan_Adj_Dtl()
  Dim tmpPeriod As Long
  
    If Trim(Txtc_AdjMonth) <> "" And Trim(Txtc_AdjYear) <> "" Then
       tmpPeriod = Is_Null(Format(Val(Txtc_AdjYear), "0000") & Format(Val(Txtc_AdjMonth), "00"), True)
       
       g_Sql = "Update pr_loan_dtl Set c_flag = 'O', n_paidamount = n_amount, c_adjremarks = '" & Trim(Txtc_AdjRemarks) & "', " & _
               "c_adjby = '" & Trim(g_UserName) & "', d_adjon = '" & GetDateTime & "' " & _
               "Where c_loancode = '" & Trim(Txtc_LoanCode) & "' and n_period = " & tmpPeriod
       CON.Execute g_Sql
       
       'g_Sql = "MADA_PR_LOAN_DTL_REUPD 'A', 'A'"
       'CON.Execute g_Sql
    End If
End Sub

Private Sub Start_Generate_New()
  Dim MaxNo As ADODB.Recordset
  g_Sql = "Select max(substring(c_loancode,6,5)) from pr_loan_mst " & _
          "where substring(c_loancode,3,2) = '" & Right(Format(Year(g_CurrentDate), "0000"), 2) & "'"
  Set MaxNo = CON.Execute(g_Sql)
  Txtc_LoanCode = "L/" + Right(Format(Year(g_CurrentDate), "0000"), 2) + "/" + Format(Is_Null(MaxNo(0).Value, True) + 1, "00000")
End Sub

Private Sub Display_Records()
On Error GoTo Err_Display
  Dim DyDisp As New ADODB.Recordset
  Dim i, j As Long
  
  Set DyDisp = Nothing
  g_Sql = "select * from pr_loan_mst where c_loancode = '" & Trim(Txtc_LoanCode) & "' and c_rec_sta='A'"
  DyDisp.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
  
  Txtc_LoanCode = Is_Null(DyDisp("c_loancode").Value, False)
  Dtp_Date.Text = Is_Date(DyDisp("d_date").Value, "D")
  Txtc_EmployeeName = Is_Null(DyDisp("c_empno").Value, False)
  Call Txtc_EmployeeName_Validate(True)
  
  For i = 0 To Cmb_Type.ListCount - 1
      If Right(Trim(Cmb_Type.List(i)), 1) = Is_Null(DyDisp("c_type").Value, False) Then
         Cmb_Type.ListIndex = i
         Exit For
      End If
  Next i
  
  Txtn_LoanAmount = Format_Num(Is_Null(DyDisp("n_loanamount").Value, True))
  Txtn_NoInstall = Is_Null(DyDisp("n_noinstall").Value, True)
  Txtc_Month = Format(Month(DyDisp("d_deductstarts").Value), "00")
  Txtc_Year = Format(Year(DyDisp("d_deductstarts").Value), "0000")
  Dtp_PaidOn.Text = Is_Date(DyDisp("d_paidon").Value, "D")
  
  Txtn_PaidAmount = Format_Num(Is_Null(DyDisp("n_loanpaid").Value, True))
  Txtn_PaidInstall = Is_Null(DyDisp("n_paidinstall").Value, True)
  Txtn_Balance = Format_Num(Is_Null_D(Txtn_LoanAmount, True) - Is_Null_D(Txtn_PaidAmount, True))
  Txtc_Remarks = Is_Null(DyDisp("c_remarks").Value, False)
  
  For i = 0 To Cmb_EPZRemarks.ListCount - 1
      If Trim(Cmb_EPZRemarks.List(i)) = Is_Null(Trim(DyDisp("c_epzremarks").Value), False) Then
         Cmb_EPZRemarks.ListIndex = i
         Exit For
      End If
  Next i

 
  ' // Details
    
   Set DyDisp = Nothing
   g_Sql = "select * from pr_loan_dtl where c_loancode='" & Trim(Txtc_LoanCode) & "' order by n_period "
   DyDisp.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    
   If DyDisp.RecordCount > 0 Then
      Va_Loan.MaxRows = DyDisp.RecordCount + 1
      DyDisp.MoveFirst
      For i = 1 To DyDisp.RecordCount
          Va_Loan.Row = i
          Va_Loan.Col = 1
             Va_Loan.Text = Right(Trim(Str(DyDisp("n_period").Value)), 2)
          Va_Loan.Col = 2
             Va_Loan.Text = Left(Trim(Str(DyDisp("n_period").Value)), 4)
          Va_Loan.Col = 3
             Va_Loan.Text = Is_Null(DyDisp("n_amount").Value, True)
          Va_Loan.Col = 4
             Va_Loan.Text = Is_Null(DyDisp("n_paidamount").Value, True)
          Va_Loan.Col = 5
             Va_Loan.Text = Is_Null(DyDisp("c_flag").Value, False)
          Va_Loan.Col = 6
             Va_Loan.Text = Is_Null(DyDisp("c_adjremarks").Value, False)
          Va_Loan.Col = 7
             Va_Loan.Text = Is_Null(DyDisp("c_adjby").Value, False)
          Va_Loan.Col = 8
             Va_Loan.Text = Is_DateSpread(DyDisp("d_adjon").Value, False)
          DyDisp.MoveNext
      Next i
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
      Search.Query = "select c_empno EmpNo, c_name Name, c_othername OtherName, c_branch Branch, c_dept Dept " & _
                     "from pr_emp_mst where d_dol is null"
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
   
   If Trim(Txtc_EmployeeName) <> "" Then
      Set RsChk = Nothing
      g_Sql = "select c_empno, c_name, c_othername, c_title " & _
              "from pr_emp_mst where c_rec_sta='A' and c_empno = '" & Trim(Right(Trim(Txtc_EmployeeName), 7)) & "'"
      RsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
      If RsChk.RecordCount > 0 Then
         Txtc_EmployeeName = Is_Null(RsChk("c_name").Value, False) & " " & Is_Null(RsChk("c_othername").Value, False) & Space(100) & RsChk("c_empno").Value
      Else
         MsgBox "Employee details not found. Press <F2> to select", vbInformation, "Information"
         Txtc_EmployeeName.SetFocus
         Cancel = True
      End If
   End If
End Sub

Private Sub Assign_DeductPeriod()
  Dim i, vMonth, vYear As Integer
   
   If Trim(Txtc_Month) = "" Or Trim(Txtc_Year) = "" Then
      Exit Sub
   End If
   
      Clear_Spread Va_Loan
      Va_Loan.MaxRows = 1000
      vMonth = Is_Null_D(Txtc_Month, True)
      vYear = Is_Null_D(Txtc_Year, True)
      For i = 1 To Is_Null_D(Txtn_NoInstall, True)
          Va_Loan.Row = i
          Va_Loan.Col = 1
             If vMonth > 12 Then
                Va_Loan.Text = 1
                vMonth = 1
                vYear = vYear + 1
             Else
                Va_Loan.Text = vMonth
             End If
             vMonth = vMonth + 1
          Va_Loan.Col = 2
             Va_Loan.Text = vYear
          Va_Loan.Col = 3
             Va_Loan.Text = Round(Is_Null_D(Txtn_LoanAmount, True) / Is_Null_D(Txtn_NoInstall, True), 2)
      Next i
      Va_Loan.MaxRows = Va_Loan.DataRowCnt + 1
      Call Spread_Row_Height(Va_Loan)
End Sub

Private Sub Txtn_LoanAmount_KeyPress(KeyAscii As Integer)
   Call OnlyNumeric(Txtn_LoanAmount, KeyAscii, 10, 2)
End Sub

Private Sub Txtn_LoanAmount_Validate(Cancel As Boolean)
    Txtn_LoanAmount = Format_Num(Txtn_LoanAmount)
End Sub

Private Sub Txtn_NoInstall_KeyPress(KeyAscii As Integer)
   Call OnlyNumeric(Txtn_NoInstall, KeyAscii, 3)
End Sub

Private Sub Va_Loan_KeyDown(KeyCode As Integer, Shift As Integer)
   
  If ((Shift And 1) = 1) And KeyCode = vbKeyInsert And g_Admin Then
     Va_Loan.Row = Va_Loan.ActiveRow
     Va_Loan.Action = 7
     Va_Loan.Row = Va_Loan.ActiveRow
     Va_Loan.Col = 1
     Va_Loan.Action = 0
  ElseIf ((Shift And 1) = 1) And KeyCode = vbKeyDelete And g_Admin Then
     Va_Loan.Row = Va_Loan.ActiveRow
     Va_Loan.Col = 4
        If Is_Null_D(Va_Loan.Text, True) > 0 Then
           MsgBox "Already Paid Loan. No access to Delete", vbInformation, "Information"
        Else
           Va_Loan.Row = Va_Loan.ActiveRow
           Va_Loan.Col = 1
           Va_Loan.Action = 5
        End If
  End If
End Sub

Private Sub Va_Loan_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    If Col = 1 Then
       Va_Loan.Row = Row
       Va_Loan.Col = 1
          If Trim(Va_Loan.Text) <> "" Then
             If Is_Null_D(Va_Loan.Text, True) <= 0 Or Is_Null_D(Va_Loan.Text, True) > 12 Then
                Va_Loan.Text = ""
                MsgBox "Not a valid month", vbInformation, "Information"
                Va_Loan.SetFocus
                Cancel = True
             End If
          End If
    End If
End Sub

Private Sub Spread_Lock()
  Dim i As Integer
  
    For i = 1 To Va_Loan.MaxCols
        Va_Loan.Row = -1
        Va_Loan.Col = i
           If g_Admin And (i = 3) Then
              Va_Loan.Lock = False
           Else
              Va_Loan.Lock = True
           End If
    Next i
End Sub

Private Sub Load_Combo()
  Dim rsCombo As New ADODB.Recordset

    Cmb_Type.Clear
    Cmb_Type.AddItem "Company Loan" & Space(100) & "0"
    Cmb_Type.AddItem "EPZ Loan" & Space(100) & "1"
    
    Cmb_EPZRemarks.Clear
    Cmb_EPZRemarks.AddItem ""
    Cmb_EPZRemarks.AddItem "House Enhancing"
    Cmb_EPZRemarks.AddItem "Computer"
    Cmb_EPZRemarks.AddItem "Medical"
    Cmb_EPZRemarks.AddItem "Marriage"
    Cmb_EPZRemarks.AddItem "Multipurpose 1"
    Cmb_EPZRemarks.AddItem "Multipurpose 2"
    Cmb_EPZRemarks.AddItem "Holiday 1"
    Cmb_EPZRemarks.AddItem "Holiday 2"
    Cmb_EPZRemarks.AddItem "School Materials"
    Cmb_EPZRemarks.AddItem "Educational"
    
End Sub

Private Sub Txtc_Month_KeyPress(KeyAscii As Integer)
    Call OnlyNumeric(Txtc_Month, KeyAscii, 2)
End Sub

Private Sub Txtc_Month_Validate(Cancel As Boolean)
  Dim tmpPeriod As String
    If Trim(Txtc_Month) <> "" Then
       If Is_Null_D(Txtc_Month, True) <= 0 Or Is_Null_D(Txtc_Month, True) > 13 Then
          MsgBox "Not a valid month", vbInformation, "Information"
          Txtc_Month.SetFocus
          Cancel = True
       ElseIf Len(Trim(Txtc_Month)) <> 2 Then
          MsgBox "Month Should be 2 digit. Example 01, 02 etc.", vbInformation, "Information"
          Txtc_Month.SetFocus
          Cancel = True
       End If
    End If
    
    Call Assign_DeductPeriod
    
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
  Call Assign_DeductPeriod
End Sub

Private Sub Txtc_AdjMonth_KeyPress(KeyAscii As Integer)
    Call OnlyNumeric(Txtc_AdjMonth, KeyAscii, 2)
End Sub

Private Sub Txtc_AdjMonth_Validate(Cancel As Boolean)
  Dim tmpPeriod As String
    If Trim(Txtc_AdjMonth) <> "" Then
       If Is_Null_D(Txtc_AdjMonth, True) <= 0 Or Is_Null_D(Txtc_AdjMonth, True) > 13 Then
          MsgBox "Not a valid month", vbInformation, "Information"
          Txtc_AdjMonth.SetFocus
          Cancel = True
       Else
          Txtc_AdjMonth = Format(Val(Txtc_AdjMonth), "00")
       End If
    End If
End Sub

Private Sub Txtc_AdjYear_KeyPress(KeyAscii As Integer)
    Call OnlyNumeric(Txtc_AdjYear, KeyAscii, 4)
End Sub

Private Sub Txtc_AdjYear_Validate(Cancel As Boolean)
  If Trim(Txtc_AdjYear) <> "" Then
     If Len(Txtc_AdjYear) <> 4 Then
        MsgBox "Not a valid year", vbInformation, "Information"
        Txtc_AdjYear.SetFocus
        Cancel = True
     End If
  End If
End Sub

