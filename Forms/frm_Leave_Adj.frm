VERSION 5.00
Object = "{D3400A36-5DD1-4308-8669-3D13D3C27560}#1.0#0"; "CS_DateControl.ocx"
Begin VB.Form frm_Leave_Adj 
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
   Begin VB.Frame Frame2 
      Height          =   900
      Left            =   120
      TabIndex        =   28
      Top             =   -30
      Width           =   9120
      Begin VB.CommandButton Btn_Print 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   3435
         Picture         =   "frm_Leave_Adj.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Delete 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   2715
         Picture         =   "frm_Leave_Adj.frx":35EE
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Clear 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   1725
         Picture         =   "frm_Leave_Adj.frx":6C98
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Exit 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   4470
         Picture         =   "frm_Leave_Adj.frx":A308
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Save 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   30
         Picture         =   "frm_Leave_Adj.frx":D968
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_View 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   1005
         Picture         =   "frm_Leave_Adj.frx":10FF1
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   700
      End
      Begin VB.Frame Frm_EmpStatus 
         Caption         =   "Adjusted On"
         ForeColor       =   &H00C00000&
         Height          =   720
         Left            =   6045
         TabIndex        =   29
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
            Left            =   2010
            TabIndex        =   7
            Top             =   330
            Width           =   765
         End
         Begin VB.OptionButton Opt_Year 
            Caption         =   "Current Year"
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
            Width           =   1515
         End
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
      Height          =   3840
      Left            =   120
      TabIndex        =   17
      Top             =   1275
      Width           =   6285
      Begin VB.TextBox Txtn_Year 
         Height          =   300
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   11
         Top             =   1170
         Width           =   1155
      End
      Begin VB.ComboBox Cmb_Type 
         Height          =   315
         Left            =   4710
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2670
         Width           =   1440
      End
      Begin VB.ComboBox Cmb_Leave 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1470
         Width           =   4575
      End
      Begin VB.TextBox Txtc_Code 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   1560
         TabIndex        =   8
         Top             =   240
         Width           =   1125
      End
      Begin VB.TextBox Txtn_Adjusted 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   4710
         TabIndex        =   14
         Top             =   2355
         Width           =   1410
      End
      Begin VB.TextBox Txtn_Balance 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   4710
         TabIndex        =   13
         Top             =   2040
         Width           =   1410
      End
      Begin VB.TextBox Txtc_Remarks 
         Height          =   300
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   16
         Top             =   3360
         Width           =   4575
      End
      Begin VB.TextBox Txtc_EmployeeName 
         Height          =   300
         Left            =   1560
         TabIndex        =   10
         Top             =   870
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
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Leave Year"
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
         Left            =   555
         TabIndex        =   30
         Top             =   1215
         Width           =   915
      End
      Begin VB.Label Label14 
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
         Height          =   210
         Left            =   4245
         TabIndex        =   27
         Top             =   2730
         Width           =   405
      End
      Begin VB.Label Label12 
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
         Height          =   210
         Left            =   975
         TabIndex        =   26
         Top             =   1515
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00404080&
         Height          =   210
         Left            =   1035
         TabIndex        =   25
         Top             =   285
         Width           =   435
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   15
         X2              =   7200
         Y1              =   3165
         Y2              =   3165
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
         TabIndex        =   24
         Top             =   285
         Width           =   360
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Adjusted"
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
         Left            =   3885
         TabIndex        =   23
         Top             =   2415
         Width           =   750
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         X1              =   15
         X2              =   7210
         Y1              =   690
         Y2              =   690
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
         TabIndex        =   22
         Top             =   3420
         Width           =   750
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Leave Balance at the time of adjusted"
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
         Left            =   1545
         TabIndex        =   21
         Top             =   2085
         Width           =   3075
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
         TabIndex        =   20
         Top             =   915
         Width           =   1320
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
      Left            =   4695
      TabIndex        =   19
      Top             =   990
      Width           =   360
   End
   Begin VB.Label lbl_scr_name 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Leave Encash / Adjustment"
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
      TabIndex        =   18
      Top             =   990
      Width           =   2235
   End
   Begin VB.Shape shp_scr_name 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   360
      Left            =   120
      Top             =   915
      Width           =   6300
   End
End
Attribute VB_Name = "frm_Leave_Adj"
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
    Call TGControlProperty(Me)
    Cmb_Type.ListIndex = 0
    
    Txtc_Code.Enabled = False
    Txtn_Balance.Enabled = False
    
End Sub

Private Sub Btn_Clear_Click()
    Clear_Controls Me
    Cmb_Type.ListIndex = 0
    Dtp_Date.Text = Is_Date(Date, "D")
    Dtp_Date.SetFocus
    Call Enable_Controls(Me, True)

    Txtc_Code.Enabled = False
    Txtn_Balance.Enabled = False
End Sub

Private Sub Btn_Exit_Click()
    Unload Me
End Sub

Private Sub Btn_Delete_Click()
On Error GoTo ErrDel
  Dim rsChk As New ADODB.Recordset
  
    If Trim(Txtc_Code) = "" Then
       Exit Sub
    End If

    Set rsChk = Nothing
    g_Sql = "select d_date from pr_leave_adj where c_code = '" & Trim(Txtc_Code) & "'"
    rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If rsChk.RecordCount > 0 Then
       If Year(rsChk("d_date").Value) < Year(g_CurrentDate) Then
          MsgBox "No access to delete previouse year record", vbInformation, "Information"
          Exit Sub
       End If
    End If

    If (MsgBox("Are you sure you want to delete ?", vbExclamation + vbYesNo, "Caution") = vbYes) Then
       CON.BeginTrans
       CON.Execute "update pr_leave_adj set " & GetDelFlag & " where c_code='" & Trim(Txtc_Code) & "'"
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
   
'   SelFor = "{MADA_pr_leave_adj.c_code}='" & Trim(Txtc_code) & "'"
'   Call Print_Rpt(SelFor, "Pr_Loan_List.rpt")
'   Mdi_Report.CRY1.Action = 1
  
  Exit Sub

Err_Print:
    MsgBox "Error while Generating - " + Err.Description, vbInformation, "Information"
End Sub


Private Sub Btn_View_Click()
  Dim Search As New Search.MyClass, SerVar
  Dim tmpFilter As String
  
    If Opt_Year.Value = True Then
       tmpFilter = tmpFilter & " and year(d_date) = " & Year(g_CurrentDate)
    End If
  
    Search.Query = "select a.c_code Code, a.c_empno EmpNo, b.c_name Name, b.c_othername OtherName, " & _
                   "a.c_leave, a.n_adjusted Adjusted " & _
                   "from pr_leave_adj a, pr_emp_mst b " & _
                   "where a.c_empno = b.c_empno and a.c_rec_sta = 'A' and b.c_rec_sta = 'A' " & tmpFilter
    Search.CheckFields = "Code"
    Search.ReturnField = "Code"
    SerVar = Search.Search(, , CON)
    If Len(Search.col1) <> 0 Then
       Txtc_Code = Search.col1
       Call Display_Records
       
       Txtc_EmployeeName.Enabled = False
       Txtn_Year.Enabled = False
       Cmb_Leave.Enabled = False
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
        Save_Pr_Leave_Adj
     CON.CommitTrans
     
     g_Sql = "HR_LEAVE_REUPDATE_PROC " & Val(Txtn_Year) & ", '" & Trim(Right(Trim(Txtc_EmployeeName), 7)) & "'"
     CON.Execute g_Sql
     
    
     g_SaveFlagNull = False
     Screen.MousePointer = vbDefault
    
     MsgBox "Record Saved Successfully", vbInformation, "Information"
     
     Txtc_EmployeeName.Enabled = False
     Txtn_Year.Enabled = False
     Cmb_Leave.Enabled = False
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
  ElseIf Val(Txtn_Year) <= 0 Then
     MsgBox "Not a valid Leave Year", vbInformation, "Information"
     Txtn_Year.SetFocus
     Exit Function
  ElseIf Trim(Cmb_Leave) = "" Then
     MsgBox "Leave Should not be empty", vbInformation, "Information"
     Cmb_Leave.SetFocus
     Exit Function
  ElseIf Trim(Cmb_Type) = "" Then
     MsgBox "Type Should not be empty", vbInformation, "Information"
     Cmb_Type.SetFocus
     Exit Function
  ElseIf Is_Null_D(Txtn_Adjusted, True) = 0 Then
     MsgBox "Adjusted should not be Zero", vbInformation, "Information"
     Txtn_Adjusted.SetFocus
     Exit Function
  ElseIf Is_Null_D(Txtn_Balance, True) < Is_Null_D(Txtn_Adjusted, True) Then
     MsgBox "Leave adjustment should not more than leave balance", vbInformation, "Information"
     Txtn_Adjusted.SetFocus
     Exit Function
  End If
  
  ChkSave = True
End Function

Private Sub Save_Pr_Leave_Adj()
    Set rs = Nothing
    g_Sql = "Select * from pr_leave_adj where c_code = '" & Trim(Txtc_Code) & "'"
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
       
       rs("c_code").Value = Is_Null(Txtc_Code, False)
       rs("d_date").Value = IIf(IsDate(Dtp_Date.Text), Format(Dtp_Date.Text, "dd/mm/yyyy"), Null)
       rs("c_empno").Value = Is_Null(Right(Trim(Txtc_EmployeeName), 7), False)
       rs("c_leave").Value = Is_Null(Right(Trim(Cmb_Leave), 7), False)
       rs("n_year").Value = Is_Null_D(Txtn_Year, True)
       rs("c_type").Value = Is_Null(Right(Trim(Cmb_Type), 1), False)
       
       rs("n_balance").Value = Is_Null_D(Txtn_Balance, True)
       rs("n_adjusted").Value = Is_Null_D(Txtn_Adjusted, True)
       rs("c_remarks").Value = Is_Null(Txtc_Remarks, False)
       rs("c_rec_sta").Value = "A"

       rs.Update
End Sub

Private Sub Start_Generate_New()
  Dim MaxNo As ADODB.Recordset
  g_Sql = "Select max(substring(c_code,6,5)) from pr_leave_adj " & _
          "where substring(c_code,3,2) = '" & Right(Format(Year(g_CurrentDate), "0000"), 2) & "'"
  Set MaxNo = CON.Execute(g_Sql)
  Txtc_Code = "A/" & Right(Format(Year(g_CurrentDate), "0000"), 2) & "/" + Format(Is_Null(MaxNo(0).Value, True) + 1, "00000")
End Sub

Private Sub Display_Records()
On Error GoTo Err_Display
  Dim DyDisp As New ADODB.Recordset
  Dim i, j As Long
  
  Set DyDisp = Nothing
  g_Sql = "select * from pr_leave_adj where c_code = '" & Trim(Txtc_Code) & "' and c_rec_sta='A'"
  DyDisp.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
  If DyDisp.RecordCount > 0 Then
     Txtc_Code = Is_Null(DyDisp("c_code").Value, False)
     Dtp_Date.Text = Is_Date(DyDisp("d_date").Value, "D")
     Txtc_EmployeeName = Is_Null(DyDisp("c_empno").Value, False)
     Call Txtc_EmployeeName_Validate(True)
     
     For i = 0 To Cmb_Leave.ListCount - 1
         If Trim(Right(Trim(Cmb_Leave.List(i)), 7)) = Is_Null(DyDisp("c_leave").Value, False) Then
            Cmb_Leave.ListIndex = i
            Exit For
         End If
     Next i
    
     For i = 0 To Cmb_Type.ListCount - 1
         If Right(Trim(Cmb_Type.List(i)), 1) = Is_Null(DyDisp("c_type").Value, False) Then
            Cmb_Type.ListIndex = i
            Exit For
         End If
     Next i
    
     Txtn_Year = Is_Null(DyDisp("n_year").Value, True)
     Txtn_Balance = Format_Num(Is_Null(DyDisp("n_balance").Value, True))
     Txtn_Adjusted = Format_Num(Is_Null(DyDisp("n_adjusted").Value, True))
     Txtc_Remarks = Is_Null(DyDisp("c_remarks").Value, False)
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
 
   If KeyCode = vbKeyDelete Then
      Txtc_EmployeeName = ""
   End If
 
   If KeyCode = vbKeyF2 Then
      Search.Query = "select c_empno EmpNo, c_name Name, c_othername OtherName, c_branch Branch, c_dept Dept " & _
                     "from pr_emp_mst where d_dol is null and c_rec_sta = 'A'"
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
   
   If Trim(Txtc_EmployeeName) <> "" Then
      Set rsChk = Nothing
      g_Sql = "select c_empno, c_name, c_othername, c_title " & _
              "from pr_emp_mst where c_rec_sta='A' and c_empno = '" & Trim(Right(Trim(Txtc_EmployeeName), 7)) & "'"
      rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
      If rsChk.RecordCount > 0 Then
         Txtc_EmployeeName = Is_Null(rsChk("c_name").Value, False) & " " & Is_Null(rsChk("c_othername").Value, False) & Space(100) & rsChk("c_empno").Value
         Call GetLeaveBalance
      Else
         MsgBox "Employee details not found. Press <F2> to select", vbInformation, "Information"
         Txtc_EmployeeName.SetFocus
         Cancel = True
      End If
   End If
End Sub

Private Sub Cmb_Leave_Click()
    Call GetLeaveBalance
End Sub

Private Sub Txtn_Adjusted_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDelete Then
      Txtn_Adjusted = ""
   End If
End Sub

Private Sub Txtn_Adjusted_KeyPress(KeyAscii As Integer)
   Call OnlyNumeric(Txtn_Adjusted, KeyAscii, 10, 2)
End Sub

Private Sub Load_Combo()
  Dim rsCombo As New ADODB.Recordset
  Dim i As Integer
  
    Cmb_Type.Clear
    Cmb_Type.AddItem "Cash" & Space(100) & "C"
    Cmb_Type.AddItem "Adjust" & Space(100) & "A"

    
    Set rsCombo = Nothing
    g_Sql = "select c_leave, c_leavename from pr_leave_mst where c_leave in ('CL','SL','VL') " & _
            "order by c_leavename "
    rsCombo.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    Cmb_Leave.Clear
    Cmb_Leave.AddItem ""
    For i = 0 To rsCombo.RecordCount - 1
        Cmb_Leave.AddItem rsCombo("c_leavename").Value & Space(100) & rsCombo("c_leave").Value
        rsCombo.MoveNext
    Next i
    
End Sub

Private Sub GetLeaveBalance()
  Dim rsChk As New ADODB.Recordset

    If Trim(Txtc_EmployeeName) <> "" And Trim(Cmb_Leave) <> "" Then
       Set rsChk = Nothing
       g_Sql = "select n_clbal from pr_emp_leave_dtl " & _
               "where year(d_prfrom) = " & Val(Txtn_Year) & " and c_empno = '" & Trim(Right(Trim(Txtc_EmployeeName), 7)) & "' and " & _
               "c_leave = '" & Trim(Right(Trim(Cmb_Leave), 7)) & "'"
       rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
       If rsChk.RecordCount > 0 Then
          Txtn_Balance = Format_Num(Is_Null(rsChk("n_clbal").Value, True))
       End If
    End If
End Sub

Private Sub Txtn_Adjusted_Validate(Cancel As Boolean)
    Txtn_Adjusted = Format_Num(Is_Null_D(Txtn_Adjusted, True))
    If Is_Null_D(Txtn_Balance, True) < Is_Null_D(Txtn_Adjusted, True) Then
       MsgBox "Leave adjustment should not more than leave balance", vbInformation, "Information"
       Txtn_Adjusted.SetFocus
       Cancel = True
    End If
End Sub

Private Sub Txtn_Year_KeyPress(KeyAscii As Integer)
    Call OnlyNumeric(Txtn_Year, KeyAscii, 4)
End Sub

Private Sub Txtn_Year_Validate(Cancel As Boolean)
  If Trim(Txtn_Year) <> "" Then
     If Len(Txtn_Year) <> 4 Then
        MsgBox "Not a valid year", vbInformation, "Information"
        Txtn_Year.SetFocus
        Cancel = True
     Else
        Call GetLeaveBalance
     End If
  End If
End Sub
