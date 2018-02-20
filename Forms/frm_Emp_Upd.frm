VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Begin VB.Form frm_Emp_Upd 
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   11130
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3840
   ScaleWidth      =   11130
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   900
      Left            =   120
      TabIndex        =   28
      Top             =   -30
      Width           =   19260
      Begin VB.CommandButton Btn_Print 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   1710
         Picture         =   "frm_Emp_Upd.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Clear 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   975
         Picture         =   "frm_Emp_Upd.frx":35EE
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Exit 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   5145
         Picture         =   "frm_Emp_Upd.frx":6C5E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Save 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   30
         Picture         =   "frm_Emp_Upd.frx":A2BE
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Restore 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   3435
         Picture         =   "frm_Emp_Upd.frx":D947
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Hide 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   2700
         Picture         =   "frm_Emp_Upd.frx":1100D
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Export 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   4410
         Picture         =   "frm_Emp_Upd.frx":14712
         Style           =   1  'Graphical
         TabIndex        =   6
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
      TabIndex        =   20
      Top             =   1290
      Width           =   19245
      Begin VB.CommandButton Btn_Display 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   17385
         Picture         =   "frm_Emp_Upd.frx":17D59
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Display"
         Top             =   150
         Width           =   700
      End
      Begin VB.Frame Frm_EmpStatus 
         Caption         =   "Employee Status"
         ForeColor       =   &H00C00000&
         Height          =   705
         Left            =   13005
         TabIndex        =   27
         Top             =   150
         Width           =   3360
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
            Left            =   2400
            TabIndex        =   15
            Top             =   285
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
            Left            =   1275
            TabIndex        =   14
            Top             =   285
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
            Left            =   135
            TabIndex        =   13
            Top             =   285
            Width           =   885
         End
      End
      Begin VB.ComboBox Cmb_Desig 
         Height          =   315
         Left            =   5685
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   525
         Width           =   2370
      End
      Begin VB.ComboBox Cmb_Branch 
         Height          =   315
         Left            =   1635
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   525
         Width           =   2370
      End
      Begin VB.ComboBox Cmb_Dept 
         Height          =   315
         Left            =   5685
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   210
         Width           =   2370
      End
      Begin VB.ComboBox Cmb_Company 
         Height          =   315
         Left            =   1635
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   210
         Width           =   2370
      End
      Begin VB.ComboBox Cmb_EmpType 
         Height          =   315
         Left            =   9750
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   210
         Width           =   2370
      End
      Begin VB.TextBox Txtc_EmployeeName 
         Height          =   300
         Left            =   9750
         TabIndex        =   12
         Top             =   525
         Width           =   2370
      End
      Begin VB.Label Label1 
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
         Left            =   5130
         TabIndex        =   26
         Top             =   585
         Width           =   465
      End
      Begin VB.Label Label4 
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
         Left            =   975
         TabIndex        =   25
         Top             =   570
         Width           =   570
      End
      Begin VB.Label Label7 
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
         Left            =   5220
         TabIndex        =   24
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
         Left            =   765
         TabIndex        =   23
         Top             =   255
         Width           =   780
      End
      Begin VB.Label Label5 
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
         Left            =   9270
         TabIndex        =   22
         Top             =   262
         Width           =   405
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
         Left            =   9225
         TabIndex        =   21
         Top             =   570
         Width           =   465
      End
   End
   Begin FPUSpreadADO.fpSpread Va_Details 
      Height          =   7260
      Left            =   120
      TabIndex        =   17
      Top             =   2280
      Width           =   19245
      _Version        =   458752
      _ExtentX        =   33946
      _ExtentY        =   12806
      _StockProps     =   64
      AutoClipboard   =   0   'False
      ColsFrozen      =   4
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
      MaxCols         =   67
      MaxRows         =   100
      ProcessTab      =   -1  'True
      SpreadDesigner  =   "frm_Emp_Upd.frx":1B497
   End
   Begin MSComDlg.CommonDialog ComDialog 
      Left            =   19590
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Left            =   17700
      TabIndex        =   19
      Top             =   1020
      Width           =   375
   End
   Begin VB.Label lbl_scr_name 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Employee Details Updation"
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
      TabIndex        =   18
      Top             =   1005
      Width           =   2175
   End
   Begin VB.Shape shp_scr_name 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   120
      Top             =   960
      Width           =   19245
   End
End
Attribute VB_Name = "frm_Emp_Upd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mnuOption As String
Private rs As New ADODB.Recordset
Private vF1 As String, vF2 As String, vF3 As String, vF4 As String, vF5 As String
Private SelFor As String, RepTitle As String, RepDate As String, sDate As String

Private Sub Form_Load()
    lbl_date.Caption = Format(Date, "dd-mmm-yyyy")
    If mnuOption = "T" Then
       lbl_scr_name.Caption = "Transport Planning Updation"
    Else
       lbl_scr_name.Caption = "Employee Details Updation"
    End If
    
    Clear_Spread Va_Details
    Call Spread_Lock
    Call Load_Combo
    Call TGControlProperty(Me)
    Call Spread_Row_Height(Va_Details)
    Call Spread_ColHide
    
    If mnuOption = "T" Then
       Btn_Hide.Visible = False
       Btn_Restore.Visible = False
    End If
    Opt_Active.Value = True
    
'    Btn_View.Visible = False
'    Btn_Delete.Visible = False
End Sub

Private Sub Btn_Clear_Click()
    Clear_Controls Me
    Clear_Spread Va_Details
End Sub

Private Sub Btn_Exit_Click()
    Unload Me
End Sub

Private Sub Btn_Save_Click()
On Error GoTo ErrSave
      
     If ChkSave = False Then
        Exit Sub
     End If

     g_SaveFlagNull = True
     Screen.MousePointer = vbHourglass

     CON.BeginTrans
         Call Update_Pr_Emp_Mst
     CON.CommitTrans

     Screen.MousePointer = vbDefault
     g_SaveFlagNull = False

     MsgBox "Record Saved Successfully", vbInformation, "Information"
     Exit Sub

ErrSave:
     CON.RollbackTrans
     g_SaveFlagNull = False
     Screen.MousePointer = vbDefault
     MsgBox "Error while Saving - " + Err.Description, vbCritical, "Critical"
End Sub

Private Function ChkSave() As Boolean
  ChkSave = True
End Function

Private Sub Update_Pr_Emp_Mst()
  Dim i As Long
  Dim tmpEmpNo As String

    For i = 1 To Va_Details.DataRowCnt
        Va_Details.Row = i
        Va_Details.Col = 1
           If Va_Details.Value = 1 Then
              Va_Details.Col = 2
                 tmpEmpNo = Trim(Va_Details.Text)
                
              If Trim(tmpEmpNo) <> "" Then
                 Set rs = Nothing
                 g_Sql = "select * from pr_emp_mst where c_rec_sta = 'A' and c_empno = '" & tmpEmpNo & "'"
                 rs.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
                   
                 Va_Details.Col = 3
                    rs("c_title").Value = Is_Null(Va_Details.Text, False)
                 Va_Details.Col = 4
                    rs("c_name").Value = Is_Null(Va_Details.Text, False)
                 Va_Details.Col = 5
                    rs("c_othername").Value = Is_Null(Va_Details.Text, False)
                 Va_Details.Col = 6
                    rs("d_dob").Value = Is_Date(Va_Details.Text, "S")
                 Va_Details.Col = 7
                    rs("c_sex").Value = Proper(Is_Null(Va_Details.Text, False))
                 Va_Details.Col = 8
                    rs("c_expatriate").Value = Proper(Is_Null(Va_Details.Text, False))
                 Va_Details.Col = 9
                    rs("c_nationality").Value = Proper(Is_Null(Va_Details.Text, False))
                 
                 Va_Details.Col = 11
                    rs("c_socsecno").Value = Is_Null(Va_Details.Text, False)
                 Va_Details.Col = 12
                    rs("c_nicno").Value = Is_Null(Va_Details.Text, False)
                 Va_Details.Col = 13
                    rs("c_matstatus").Value = Is_Null(Va_Details.Text, False)
                 Va_Details.Col = 14
                    rs("c_bloodgroup").Value = Is_Null(Va_Details.Text, False)
        
                 Va_Details.Col = 15
                    rs("c_address").Value = Is_Null(Replace(Va_Details.Text, ",", " "), False)
                 Va_Details.Col = 16
                    rs("c_phone").Value = Is_Null(Va_Details.Text, False)
                 Va_Details.Col = 17
                    rs("c_email").Value = Is_Null(Replace(Va_Details.Text, ",", " "), False)
                 Va_Details.Col = 18
                    rs("c_familydetails").Value = Is_Null(Replace(Va_Details.Text, ",", " "), False)
                 Va_Details.Col = 19
                    rs("c_qualification").Value = Is_Null(Replace(Va_Details.Text, ",", " "), False)
                 Va_Details.Col = 20
                    rs("c_specialistin").Value = Is_Null(Replace(Va_Details.Text, ",", " "), False)
                 Va_Details.Col = 21
                    rs("c_additionalinfo").Value = Is_Null(Replace(Va_Details.Text, ",", " "), False)
                   
                 Va_Details.Col = 23
                    rs("c_company").Value = Is_Null(Right(Trim(Va_Details.Text), 7), False)
                 Va_Details.Col = 24
                    rs("c_branch").Value = Is_Null(Va_Details.Text, False)
                 Va_Details.Col = 25
                    rs("c_dept").Value = Is_Null(Va_Details.Text, False)
                 Va_Details.Col = 26
                    rs("c_desig").Value = Is_Null(Va_Details.Text, False)
                 
                 Va_Details.Col = 27
                    rs("d_doj").Value = Is_Date(Va_Details.Text, "S")
                 Va_Details.Col = 28
                    rs("d_dol").Value = Is_Date(Va_Details.Text, "S")
                 Va_Details.Col = 29
                    rs("c_skillset").Value = Is_Null(Va_Details.Text, False)
                 Va_Details.Col = 30
                    rs("c_daywork").Value = Is_Null(Right(Trim(Va_Details.Text), 3), False)
                 Va_Details.Col = 31
                    rs("c_shiftcode").Value = Is_Null(Right(Trim(Va_Details.Text), 3), False)
                 Va_Details.Col = 32
                    rs("c_line").Value = Is_Null(Va_Details.Text, False)
                   
                 Va_Details.Col = 34
                    rs("n_mldays").Value = Is_Null(Va_Details.Text, True)
                 Va_Details.Col = 35
                    rs("d_mlfrom").Value = Is_Date(Va_Details.Text, "S")
                 Va_Details.Col = 36
                    rs("d_mlto").Value = Is_Date(Va_Details.Text, "S")
                 Va_Details.Col = 37
                    rs("c_tpflag").Value = Is_Null(Va_Details.Text, False)
                 Va_Details.Col = 38
                    rs("c_tpmode").Value = Is_Null(Replace(Va_Details.Text, ",", " "), False)
                 Va_Details.Col = 39
                    rs("c_town").Value = Is_Null(Replace(Va_Details.Text, ",", " "), False)
                 Va_Details.Col = 40
                    rs("c_road").Value = Is_Null(Replace(Va_Details.Text, ",", " "), False)
                 Va_Details.Col = 41
                    rs("c_pad").Value = Is_Null(Replace(Va_Details.Text, ",", " "), False)
                   
                 Va_Details.Col = 43
                    rs("c_emptype").Value = Proper(Is_Null(Va_Details.Text, False))
                 Va_Details.Col = 44
                    rs("c_stafftype").Value = Is_Null(Right(Trim(Va_Details.Text), 1), False)
                 Va_Details.Col = 45
                    rs("c_salarytype").Value = Is_Null(Right(Trim(Va_Details.Text), 2), False)
                 Va_Details.Col = 46
                    rs("c_paytype").Value = Is_Null(Right(Trim(Va_Details.Text), 2), False)
                   
                 Va_Details.Col = 47
                    rs("c_bank").Value = Is_Null(Right(Trim(Va_Details.Text), 3), False)
                 Va_Details.Col = 48
                    rs("c_bankcode").Value = Is_Null(Va_Details.Text, False)
                 Va_Details.Col = 49
                    rs("c_acctno").Value = Is_Null(Va_Details.Text, False)
                 Va_Details.Col = 50
                    rs("c_itno").Value = Is_Null(Va_Details.Text, False)
                 
                 Va_Details.Col = 53
                    rs("c_tatype").Value = Is_Null(Right(Trim(Va_Details.Text), 1), False)
                 Va_Details.Col = 54
                    rs("n_carbenefit").Value = Is_Null_D(Va_Details.Text, True)
                 Va_Details.Col = 55
                    rs("c_edfcat").Value = Is_Null(Left(Trim(Va_Details.Text), 1), False)
                 Va_Details.Col = 56
                    rs("n_edfamount").Value = Is_Null_D(Va_Details.Text, True)
                 Va_Details.Col = 57
                    rs("n_eduamount").Value = Is_Null_D(Va_Details.Text, True)
                 Va_Details.Col = 58
                    rs("n_intamount").Value = Is_Null_D(Va_Details.Text, True)
                 Va_Details.Col = 59
                    rs("n_preamount").Value = Is_Null_D(Va_Details.Text, True)
                 Va_Details.Col = 60
                    rs("n_othamount").Value = Is_Null_D(Va_Details.Text, True)
                   
                 Va_Details.Col = 62
                    rs("c_clockcard").Value = Is_Null(Va_Details.Text, True)
                 Va_Details.Col = 63
                    rs("c_payerelief").Value = Is_Null(Va_Details.Text, False)
                 Va_Details.Col = 64
                    rs("c_npfdeduct").Value = Is_Null(Va_Details.Text, False)
                 Va_Details.Col = 65
                    rs("c_nobonus").Value = Is_Null(Va_Details.Text, False)
                 Va_Details.Col = 66
                    rs("c_mealallow").Value = Is_Null(Va_Details.Text, True)
                 Va_Details.Col = 67
                    rs("c_nopay").Value = Is_Null(Va_Details.Text, True)
                                    
                 rs("c_usr_id").Value = g_UserName
                 rs("d_modified").Value = GetDateTime
                 rs.Update
             End If
          End If
    Next i
End Sub

Private Sub Btn_Print_Click()
  Dim tmpStr As String, RepOpt As String
  
    tmpStr = "": RepOpt = ""
    SelFor = "": RepTitle = "": RepDate = "": sDate = ""
    vF1 = "": vF2 = "": vF3 = "": vF4 = "": vF5 = ""
 
    SelFor = EmpMaster_RepFilter
   
    If mnuOption = "T" Then
      RepTitle = MakeReportHead(Me, "Employee Transport List", True)
      RepDate = MakeReportSubHead(Me)
      Call Print_Rpt(SelFor, "Pr_Transport_Emp_List.rpt")
    
    Else
      RepTitle = MakeReportHead(Me, "Employee List", True)
      RepDate = MakeReportSubHead(Me)
      Call Print_Rpt(SelFor, "Pr_Emp_List.rpt")
    End If
    
    If Trim(RepTitle) <> "" Then
       Mdi_Ta_HrPay.CRY1.Formulas(1) = "ReportHead='" & UCase(RepTitle) & "'"
    End If
    
    If Trim(RepDate) <> "" Then
       Mdi_Ta_HrPay.CRY1.Formulas(2) = "RepHeadDate='" & RepDate & "'"
    End If
    Mdi_Ta_HrPay.CRY1.Action = 1
End Sub

Private Sub Txtc_EmployeeName_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim Search As New Search.MyClass, SerVar
   
   If KeyCode = vbKeyDelete Then
      Txtc_EmployeeName = ""
   End If
   
   If KeyCode = vbKeyF2 Then
      Search.Query = "select c_empno EmpNo, c_name Name, c_othername OtherName, c_dept Dept, " & _
                   "c_desig Desig, c_branch Branch, c_emptype Type " & _
                   "from pr_emp_mst where c_rec_sta = 'A' "
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
     g_Sql = "select c_empno, c_name, c_othername, c_company, c_branch, c_dept, c_emptype, " & _
             "d_doj, d_dol, c_stafftype, c_daywork from pr_emp_mst " & _
             "where c_rec_sta = 'A' and c_empno = '" & Trim(Right(Trim(Txtc_EmployeeName), 7)) & "'"
     rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
     If rsChk.RecordCount > 0 Then
        Txtc_EmployeeName = Is_Null(rsChk("c_name").Value, False) & " " & Is_Null(rsChk("c_othername").Value, False) & Space(100) & Is_Null(rsChk("c_empno").Value, False)
        Call DisplayComboCompany(Me, Is_Null(rsChk("c_company").Value, False))
        Call DisplayComboBranch(Me, Is_Null(rsChk("c_branch").Value, False))
        Call DisplayComboDept(Me, Is_Null(rsChk("c_dept").Value, False))
    
        For i = 0 To Cmb_EmpType.ListCount - 1
          If Trim(Cmb_EmpType.List(i)) = Is_Null(rsChk("c_emptype").Value, False) Then
             Cmb_EmpType.ListIndex = i
             Exit For
          End If
        Next i
     Else
        MsgBox "Employee not found. Press <F2> to select.", vbInformation, "Information"
        Cancel = True
     End If
  End If
End Sub

Private Sub Spread_Lock()
  Dim i As Integer

   For i = 1 To Va_Details.MaxCols
       Va_Details.Row = -1
       Va_Details.Col = i
       If i = 2 Or i = 34 Or i = 49 Or i = 56 Then
          Va_Details.Lock = True
       Else
          Va_Details.Lock = False
       End If
   Next i
End Sub

Private Sub Display_Employee()
  Dim rsDisp As New ADODB.Recordset
  Dim i As Integer, j As Integer

    Set rsDisp = Nothing
    g_Sql = "select a.c_empno, a.c_title, a.c_name, a.c_othername,  a.d_dob, a.c_sex, a.c_nationality, a.c_expatriate, a.c_socsecno, " & _
            "a.c_nicno,  a.c_matstatus, a.c_bloodgroup, a.c_address, a.c_phone, a.c_email,a.c_familydetails, a.c_qualification, " & _
            "a.c_specialistin, a.c_additionalinfo, a.c_company, a.c_branch, a.c_dept, a.c_desig, a.d_doj, a.d_dol, a.c_skillset, " & _
            "a.c_daywork, a.c_shiftcode, a.c_line, a.c_clockidno, a.n_mldays, a.d_mlfrom, a.d_mlto, a.c_tpflag, a.c_tpmode, a.c_pad, " & _
            "a.c_town, a.c_road, a.c_paytype, a.c_bank, a.c_bankcode, a.c_acctno, a.c_itno, a.c_nopay, a.c_emptype, a.c_stafftype, " & _
            "a.c_tatype, a.c_salarytype, a.n_carbenefit, a.c_edfcat, a.n_edfamount, a.n_eduamount, a.n_intamount, " & _
            "a.n_preamount, a.n_othamount, a.c_clockcard, a.c_payerelief, a.c_npfdeduct, a.c_nobonus, a.c_mealallow, " & _
            "b.c_companyname, c.c_bankname " & _
            "from pr_emp_mst a left outer join pr_bankmast c on a.c_bank = c.c_code, pr_company_mst b  " & _
            "Where a.c_company = b.c_company and a.c_rec_sta = 'A' "
            
    If Trim(Cmb_Company) <> "" Then
       g_Sql = g_Sql & " and a.c_company = '" & Right(Trim(Cmb_Company), 7) & "'"
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
    If Trim(Txtc_EmployeeName) <> "" Then
       g_Sql = g_Sql & " and a.c_empno = '" & Trim(Right(Trim(Txtc_EmployeeName), 7)) & "'"
    End If
    If Opt_Active.Value = True Then
       g_Sql = g_Sql & " and a.d_dol is null "
    ElseIf Opt_Left.Value = True Then
       g_Sql = g_Sql & " and a.d_dol is not null "
    End If
    
    rsDisp.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    Va_Details.MaxRows = rsDisp.RecordCount + 50
    
    If rsDisp.RecordCount > 0 Then
       For i = 1 To rsDisp.RecordCount
           Va_Details.Row = i
           Va_Details.Col = 2
              Va_Details.Text = Is_Null(rsDisp("c_empno").Value, False)
           Va_Details.Col = 3
              Va_Details.Text = Is_Null(rsDisp("c_title").Value, False)
           Va_Details.Col = 4
              Va_Details.Text = Is_Null(rsDisp("c_name").Value, False)
           Va_Details.Col = 5
              Va_Details.Text = Is_Null(rsDisp("c_othername").Value, False)
              
           Va_Details.Col = 6
              Va_Details.Text = Is_DateSpread(rsDisp("d_dob").Value, False)
           Va_Details.Col = 7
              Va_Details.Text = Is_Null(rsDisp("c_sex").Value, False)
           Va_Details.Col = 8
              Va_Details.Text = Is_Null(rsDisp("c_expatriate").Value, False)
           Va_Details.Col = 9
              Va_Details.Text = Is_Null(rsDisp("c_nationality").Value, False)
           
           Va_Details.Col = 11
              Va_Details.Text = Is_Null(rsDisp("c_socsecno").Value, False)
           Va_Details.Col = 12
              Va_Details.Text = Is_Null(rsDisp("c_nicno").Value, False)
           Va_Details.Col = 13
              Va_Details.Text = Is_Null(rsDisp("c_matstatus").Value, False)
           Va_Details.Col = 14
              Va_Details.Text = Is_Null(rsDisp("c_bloodgroup").Value, False)
           
           Va_Details.Col = 15
              Va_Details.Text = Is_Null(rsDisp("c_address").Value, False)
           Va_Details.Col = 16
              Va_Details.Text = Is_Null(rsDisp("c_phone").Value, False)
           Va_Details.Col = 17
              Va_Details.Text = Is_Null(rsDisp("c_email").Value, False)
           Va_Details.Col = 18
              Va_Details.Text = Is_Null(rsDisp("c_familydetails").Value, False)
           Va_Details.Col = 19
              Va_Details.Text = Is_Null(rsDisp("c_qualification").Value, False)
           Va_Details.Col = 20
              Va_Details.Text = Is_Null(rsDisp("c_specialistin").Value, False)
           Va_Details.Col = 21
              Va_Details.Text = Is_Null(rsDisp("c_additionalinfo").Value, False)
              
           Va_Details.Col = 23
              j = SelComboString(Va_Details, Is_Null(rsDisp("c_company").Value, False), i, 23, True)
              Va_Details.TypeComboBoxIndex = j
           Va_Details.Col = 24
              Va_Details.Text = Is_Null(rsDisp("c_branch").Value, False)
           Va_Details.Col = 25
              Va_Details.Text = Is_Null(rsDisp("c_dept").Value, False)
           Va_Details.Col = 26
              Va_Details.Text = Is_Null(rsDisp("c_desig").Value, False)
           
           Va_Details.Col = 27
              Va_Details.Text = Is_DateSpread(rsDisp("d_doj").Value, False)
           Va_Details.Col = 28
              Va_Details.Text = Is_DateSpread(rsDisp("d_dol").Value, False)
              
           Va_Details.Col = 29
              Va_Details.Text = Is_Null(rsDisp("c_skillset").Value, False)
           Va_Details.Col = 30
              j = SelComboString(Va_Details, Is_Null(rsDisp("c_daywork").Value, False), i, 30, True)
              Va_Details.TypeComboBoxIndex = j
           Va_Details.Col = 31
              j = SelComboString(Va_Details, Is_Null(rsDisp("c_shiftcode").Value, False), i, 31, True)
              Va_Details.TypeComboBoxIndex = j
           Va_Details.Col = 32
              Va_Details.Text = Is_Null(rsDisp("c_line").Value, False)
           
           Va_Details.Col = 34
              Va_Details.Text = Spread_IntFormat(rsDisp("n_mldays").Value, True)
           Va_Details.Col = 35
              Va_Details.Text = Is_DateSpread(rsDisp("d_mlfrom").Value, False)
           Va_Details.Col = 36
              Va_Details.Text = Is_DateSpread(rsDisp("d_mlto").Value, False)
              
           Va_Details.Col = 37
              Va_Details.Text = Is_Null(rsDisp("c_tpflag").Value, False)
           Va_Details.Col = 38
              Va_Details.Text = Is_Null(rsDisp("c_tpmode").Value, False)
           Va_Details.Col = 39
              Va_Details.Text = Is_Null(rsDisp("c_town").Value, False)
           Va_Details.Col = 40
              Va_Details.Text = Is_Null(rsDisp("c_road").Value, False)
           Va_Details.Col = 41
              Va_Details.Text = Is_Null(rsDisp("c_pad").Value, False)
              
           Va_Details.Col = 43
              Va_Details.Text = Is_Null(rsDisp("c_emptype").Value, False)
           Va_Details.Col = 44
              j = SelComboString(Va_Details, Is_Null(rsDisp("c_stafftype").Value, False), i, 44, True)
              Va_Details.TypeComboBoxIndex = j
           Va_Details.Col = 45
              j = SelComboString(Va_Details, Is_Null(rsDisp("c_salarytype").Value, False), i, 45, True)
              Va_Details.TypeComboBoxIndex = j
           Va_Details.Col = 46
              j = SelComboString(Va_Details, Is_Null(rsDisp("c_paytype").Value, False), i, 46, True)
              Va_Details.TypeComboBoxIndex = j
              
           Va_Details.Col = 47
              Va_Details.Text = Is_Null(rsDisp("c_bankname").Value, False) & Space(100) & Is_Null(rsDisp("c_bank").Value, False)
           Va_Details.Col = 48
              Va_Details.Text = Is_Null(rsDisp("c_bankcode").Value, False)
           Va_Details.Col = 49
              Va_Details.Text = Is_Null(rsDisp("c_acctno").Value, False)
           Va_Details.Col = 50
              Va_Details.Text = Is_Null(rsDisp("c_itno").Value, False)
           
              
           Va_Details.Col = 53
              j = SelComboString(Va_Details, Is_Null(rsDisp("c_tatype").Value, False), i, 53, True)
              Va_Details.TypeComboBoxIndex = j
           Va_Details.Col = 54
              Va_Details.Text = Spread_NumFormat(rsDisp("n_carbenefit").Value, True)
           
           Va_Details.Col = 55
              j = SelComboString(Va_Details, Is_Null(rsDisp("c_edfcat").Value, False) & " - ", i, 55, True)
              Va_Details.TypeComboBoxIndex = j
           Va_Details.Col = 56
              Va_Details.Text = Spread_NumFormat(rsDisp("n_edfamount").Value, True)
           Va_Details.Col = 57
              Va_Details.Text = Spread_NumFormat(rsDisp("n_eduamount").Value, True)
           Va_Details.Col = 58
              Va_Details.Text = Spread_NumFormat(rsDisp("n_intamount").Value, True)
           Va_Details.Col = 59
              Va_Details.Text = Spread_NumFormat(rsDisp("n_preamount").Value, True)
           Va_Details.Col = 60
              Va_Details.Text = Spread_NumFormat(rsDisp("n_othamount").Value, True)
              
           Va_Details.Col = 62
              Va_Details.Text = Is_Null(rsDisp("c_clockcard").Value, False)
           Va_Details.Col = 63
              Va_Details.Text = Is_Null(rsDisp("c_payerelief").Value, False)
           Va_Details.Col = 64
              Va_Details.Text = Is_Null(rsDisp("c_npfdeduct").Value, False)
           Va_Details.Col = 65
              Va_Details.Text = Is_Null(rsDisp("c_nobonus").Value, False)
           Va_Details.Col = 66
              Va_Details.Text = Is_Null(rsDisp("c_mealallow").Value, False)
           Va_Details.Col = 67
              Va_Details.Text = Is_Null(rsDisp("c_nopay").Value, False)
    
          rsDisp.MoveNext
       Next i
    Else
       MsgBox "Employee are not found", vbInformation, "Information"
    End If
End Sub

Private Sub Btn_Display_Click()
 
   If Va_Details.DataRowCnt > 0 Then
      If MsgBox("The data's will be reset. Do you want to refresh?", vbYesNo + vbDefaultButton2, "Confirmation") = vbNo Then
         Exit Sub
      End If
   End If

   Clear_Spread Va_Details
   Call Display_Employee
   Call Spread_Row_Height(Va_Details)

End Sub

Private Sub Va_Details_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
  Dim i As Integer
  Dim tmpStr As String
  
    If BlockRow = -1 Then
       Exit Sub
    End If
    
    If BlockCol = 3 Or BlockCol = 7 Or BlockCol = 8 Or BlockCol = 9 Or BlockCol = 13 Or BlockCol = 14 Or BlockCol = 24 Or BlockCol = 25 Or BlockCol = 26 Or BlockCol = 29 Or BlockCol = 32 Or BlockCol = 38 Or BlockCol = 39 Or BlockCol = 40 Or BlockCol = 41 Or BlockCol = 47 Or BlockCol = 48 Then
       If MsgBox("Do you want to change", vbOKCancel + vbDefaultButton2, "Information") = vbOK Then
          Va_Details.Row = BlockRow
          Va_Details.Col = BlockCol
             tmpStr = Trim(Va_Details.Text)
             
          For i = BlockRow To BlockRow2
              Va_Details.Col = BlockCol
              Va_Details.Row = i
                 If Va_Details.Lock = False Then
                    Va_Details.Text = tmpStr
                    Va_Details.Col = 1
                       Va_Details.Value = 1
                 End If
          Next i
       End If
    End If
End Sub


Private Sub Va_Details_Change(ByVal Col As Long, ByVal Row As Long)
Dim vVar As Variant
 
    If Col >= 3 Then
       Va_Details.Row = Row
       Va_Details.Col = 1
          Va_Details.Value = 1
    End If
    
    If Col = 55 Then
       Va_Details.Row = Row
       Va_Details.Col = Col
          If Trim(Va_Details.Text) <> "" Then
             vVar = Split(Va_Details.Text, "~")
             Va_Details.Col = 56
                Va_Details.Text = Spread_NumFormat(Val(vVar(1)), True)
          End If
    End If
End Sub

Private Sub Va_Details_DblClick(ByVal Col As Long, ByVal Row As Long)
    Call SpreadColSort(Va_Details, Col, Row)
End Sub

Private Sub Va_Details_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim Search As New Search.MyClass, SerVar, SerArray
 
  If KeyCode = vbKeyDelete Then
     Call SpreadCellDataClear(Va_Details, Va_Details.ActiveRow, Va_Details.ActiveCol)
  
  ElseIf Va_Details.ActiveCol = 47 And KeyCode = vbKeyF2 Then
     Search.Query = "Select c_code Code, c_bankname BankName, c_bankcode BankCode from pr_bankmast where c_rec_sta='A'"
     Search.CheckFields = "Code, BankName, BankCode"
     Search.ReturnField = "Code, BankName, BankCode"
     SerVar = Search.Search(, , CON)
     SerArray = Split(SerVar, "~")
     If Len(Search.col1) <> 0 Then
        Va_Details.Row = Va_Details.ActiveRow
        Va_Details.Col = 47
           Va_Details.Text = Search.col2 & Space(100) & Search.col1
        Va_Details.Col = 48
           Va_Details.Value = SerArray(2)
        Va_Details.Col = 1
           Va_Details.Value = 1
     End If
  End If
End Sub

Private Sub Load_Combo()
  Dim rsCombo As New ADODB.Recordset
  Dim i As Integer
  Dim tmpStr As String

    Call LoadComboCompany(Me)
    Call LoadComboBranch(Me)
    Call LoadComboDept(Me)
    Call LoadComboEmpType(Me)
    Call LoadComboDesig(Me)
    
   
    'Company
    Set rsCombo = Nothing
    g_Sql = "Select c_company, c_companyname from pr_company_mst where c_rec_sta='A' order by c_companyname"
    rsCombo.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    For i = 1 To rsCombo.RecordCount
       If i = 1 Then
          tmpStr = rsCombo("c_companyname").Value & Space(100) & rsCombo("c_company").Value
       Else
          tmpStr = tmpStr & Chr$(9) & rsCombo("c_companyname").Value & Space(100) & rsCombo("c_company").Value
       End If
       rsCombo.MoveNext
    Next i

    If Len(tmpStr) > 0 Then
       Va_Details.Row = -1
       Va_Details.Col = 23
          Va_Details.TypeComboBoxList = tmpStr
    End If
    
    
    ' Daywork
    tmpStr = "5 Days" & Space(50) & "5D" & Chr$(9) & _
             "6 Days" & Space(50) & "6D" & Chr$(9) & _
             "5 Days - Shift" & Space(50) & "5DS" & Chr$(9) & _
             "6 Days - Shift" & Space(50) & "6DS"
    
    If Len(tmpStr) > 0 Then
        Va_Details.Row = -1
       Va_Details.Col = 30
          Va_Details.TypeComboBoxList = tmpStr
    End If

    
    ' Shift Name and Code
    Set rsCombo = Nothing
    g_Sql = "select c_shiftname, c_code from pr_shiftstructure_mst order by c_shiftname"
    rsCombo.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    For i = 1 To rsCombo.RecordCount
        If i = 1 Then
          tmpStr = rsCombo("c_shiftname").Value & Space(100) & rsCombo("c_code").Value
       Else
          tmpStr = tmpStr & Chr$(9) & rsCombo("c_shiftname").Value & Space(100) & rsCombo("c_code").Value
       End If
       rsCombo.MoveNext
    Next i

    If Len(tmpStr) > 0 Then
       Va_Details.Row = -1
       Va_Details.Col = 31
          Va_Details.TypeComboBoxList = tmpStr
    End If

    'Transport flag
    tmpStr = "Yes" & Chr$(9) & _
             "No"
             
    If Len(tmpStr) > 0 Then
        Va_Details.Row = -1
       Va_Details.Col = 37
          Va_Details.TypeComboBoxList = tmpStr
    End If
 
    
    'Staff Type
    tmpStr = "Flat" & Space(50) & "F" & Chr$(9) & _
             "OverTime" & Space(50) & "O"
             
    If Len(tmpStr) > 0 Then
       Va_Details.Row = -1
       Va_Details.Col = 44
          Va_Details.TypeComboBoxList = tmpStr
    End If
    
    'Salary Type
    tmpStr = "Monthly" & Space(50) & "ML" & Chr$(9) & _
             "Hourly" & Space(50) & "HR"
             
    If Len(tmpStr) > 0 Then
       Va_Details.Row = -1
       Va_Details.Col = 45
          Va_Details.TypeComboBoxList = tmpStr
    End If
    
    'Pay Type
    tmpStr = "Cash" & Space(50) & "CA" & Chr$(9) & _
             "Cheque" & Space(50) & "CH" & Chr$(9) & _
             "Bank A/c" & Space(50) & "BA"
             
    If Len(tmpStr) > 0 Then
       Va_Details.Row = -1
       Va_Details.Col = 46
          Va_Details.TypeComboBoxList = tmpStr
    End If
    
    'TA Type
    tmpStr = " " & Chr$(9) & _
             "Actual" & Space(50) & "A" & Chr$(9) & _
             "Fixed" & Space(50) & "F"
             
    If Len(tmpStr) > 0 Then
       Va_Details.Row = -1
       Va_Details.Col = 53
          Va_Details.TypeComboBoxList = tmpStr
    End If
    
    'Edf category
    Set rsCombo = Nothing
    g_Sql = "Select * from pr_edfmast order by c_category"
    rsCombo.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    For i = 1 To rsCombo.RecordCount
       If i = 1 Then
          tmpStr = rsCombo("c_category").Value & " - " & rsCombo("c_desp").Value & Space(100) & "~" & rsCombo("n_edfamt").Value
       Else
          tmpStr = tmpStr & Chr$(9) & rsCombo("c_category").Value & " - " & rsCombo("c_desp").Value & Space(100) & "~" & rsCombo("n_edfamt").Value
       End If
       rsCombo.MoveNext
    Next i
    
    If Len(tmpStr) > 0 Then
       Va_Details.Row = -1
       Va_Details.Col = 55
          Va_Details.TypeComboBoxList = tmpStr
    End If
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

Private Sub Spread_ColHide()
  Dim i As Integer, j As Integer
  
    Va_Details.Row = -1
    Va_Details.Col = 10
       Va_Details.ColHidden = True
    Va_Details.Col = 22
       Va_Details.ColHidden = True
    Va_Details.Col = 33
       Va_Details.ColHidden = True
    Va_Details.Col = 34
       Va_Details.ColHidden = True
    Va_Details.Col = 42
       Va_Details.ColHidden = True
    Va_Details.Col = 51
       Va_Details.ColHidden = True
    Va_Details.Col = 52
       Va_Details.ColHidden = True
    Va_Details.Col = 61
       Va_Details.ColHidden = True
       
 
 
    ' --- SIFB
    For i = 1 To Va_Details.MaxCols
        Va_Details.Row = -1
        Va_Details.Col = i
           If i = 10 Or i = 22 Or (i >= 33 And i <= 42) Or (i >= 45 And i <= 61) Or i >= 63 Then
              Va_Details.ColHidden = True
           Else
              Va_Details.ColHidden = False
           End If
    Next i
    ' --- SIFB
       
    If mnuOption = "T" Then
       For i = 1 To Va_Details.MaxCols
           Va_Details.Row = -1
           Va_Details.Col = i
              If (i <= 5) Or (i = 30 Or i = 31) Or (i >= 37 And i <= 41) Then
                 Va_Details.ColHidden = False
              Else
                 Va_Details.ColHidden = True
              End If
       Next i
    End If
    
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
     
     
     comDialog.FileName = "EmpDtl.CSV"
     comDialog.ShowSave
     vFileName = Trim(comDialog.FileName)
     
     If vFileName = "EmpDtl.CSV" Then  'user cancel the export to csv file in save dialog box.
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
    Screen.MousePointer = vbDefault
    MsgBox "Error while Generating - " + Err.Description, vbInformation, "Information"
    CON.RollbackTrans
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
        Va_Details.Col = 2
           If Trim(Va_Details.Text) <> "" Then
              vStr = "": vEmpNo = ""
              Ctr = Ctr + 1
              Va_Details.Row = i
              Va_Details.Col = 2
                 vEmpNo = Trim(Va_Details.Text)
                 vStr = Trim(Va_Details.Text)
              
              For j = 3 To Va_Details.MaxCols
                  Va_Details.Row = i
                  Va_Details.Col = j
                     If Va_Details.ColHidden = False Then
                        If j = 23 Or j = 30 Or j = 31 Or j = 44 Or j = 45 Or j = 46 Or j = 53 Or j = 55 Then  'combo
                           vStr = vStr & "," & Trim(Left(Trim(Va_Details.Text), 50))
                        
                        ElseIf j = 11 Or j = 12 Then  'long numbers
                           If i = 0 Then
                              vStr = vStr & "," & Trim(Va_Details.Text)
                           Else
                              If Trim(Va_Details.Text) = "" Then
                                 vStr = vStr & "," & Trim(Va_Details.Text)
                              Else
                                 vStr = vStr & ",'" & Trim(Va_Details.Text)
                              End If
                           End If
'
'                        ElseIf j = 54 Or j = 56 Or j = 57 Or j = 58 Or j = 59 Or j = 60 Then ' comma separated decimal values
'                           If i = 0 Then
'                              vStr = vStr & "," & Trim(Va_Details.Text)
'                           Else
'                              vStr = vStr & "," & Str(Is_Null_D(Va_Details.Text, True))
'                           End If
                           
                        ElseIf j = 62 Or j = 63 Or j = 64 Or j = 65 Or j = 66 Or j = 67 Then  'option box
                          If i = 0 Then
                              vStr = vStr & "," & Trim(Va_Details.Text)
                           Else
                              vStr = vStr & "," & IIf(Va_Details.Value = 1, "Yes", "No")
                           End If
                        
                        Else
                           vStr = vStr & "," & Replace(Trim(Va_Details.Text), ",", "")
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
    
    g_Sql = "update pr_export_csv set c_csv = replace(replace(c_csv, char(10), ''), char(13), '')"
    CON.Execute g_Sql
    
End Sub

Private Function EmpMaster_RepFilter() As String
  Dim vF1 As String, vF2 As String, vF3 As String, vF4 As String, vF5 As String
  Dim SelFor As String
  
     vF1 = "": vF2 = "": vF3 = "": vF4 = "": vF5 = ""
     SelFor = ""
     
     If Right(Trim(Cmb_Company), 7) <> "" Then
        vF1 = "{V_PR_EMP_MST.C_COMPANY}='" & Right(Trim(Cmb_Company), 7) & "'"
     End If
     
     If Trim(Cmb_Branch) <> "" Then
        vF2 = "{V_PR_EMP_MST.C_BRANCH}='" & Trim(Cmb_Branch) & "'"
     End If
     
     If Trim(Cmb_Dept) <> "" Then
        vF3 = "{V_PR_EMP_MST.C_DEPT}='" & Trim(Cmb_Dept) & "'"
     End If
     
     If Trim(Cmb_Desig) <> "" Then
        vF4 = "{V_PR_EMP_MST.C_DESIG}='" & Trim(Cmb_Desig) & "'"
     End If
     
     If Trim(Cmb_EmpType) <> "" Then
        vF5 = "{V_PR_EMP_MST.C_EMPTYPE}='" & Trim(Cmb_EmpType) & "'"
     End If
     
     SelFor = ReportFilterOption(vF1, vF2, vF3, vF4, vF5)
     vF1 = "": vF2 = "": vF3 = "": vF4 = "": vF5 = ""
     
     If Trim(Txtc_EmployeeName) <> "" Then
        vF1 = "{V_PR_EMP_MST.C_EMPNO}='" & Trim(Right(Trim(Txtc_EmployeeName), 7)) & "'"
     End If
     
     If Opt_Active.Value = True Then
        vF2 = "ISNULL({V_PR_EMP_MST.D_DOL})"
     ElseIf Opt_Left.Value = True Then
        vF2 = "NOT ISNULL({V_PR_EMP_MST.D_DOL})"
     End If
     
     If mnuOption = "T" Then
        vF3 = "{V_PR_EMP_MST.C_TPFLAG}='Yes'"
     End If
     
     EmpMaster_RepFilter = ReportFilterOption(SelFor, vF1, vF2, vF3, vF4)
     
End Function

