VERSION 5.00
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Begin VB.Form frm_Pay_Struture 
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
      TabIndex        =   14
      Top             =   -45
      Width           =   11925
      Begin VB.CommandButton Btn_Print 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   3435
         Picture         =   "frm_Pay_Structure.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Delete 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   2715
         Picture         =   "frm_Pay_Structure.frx":35EE
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Clear 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   1725
         Picture         =   "frm_Pay_Structure.frx":6C98
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Exit 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   4470
         Picture         =   "frm_Pay_Structure.frx":A308
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Save 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   30
         Picture         =   "frm_Pay_Structure.frx":D968
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_View 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   1005
         Picture         =   "frm_Pay_Structure.frx":10FF1
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   700
      End
   End
   Begin FPUSpreadADO.fpSpread Va_Details 
      Height          =   6495
      Left            =   120
      TabIndex        =   8
      Top             =   1905
      Width           =   11895
      _Version        =   458752
      _ExtentX        =   20981
      _ExtentY        =   11456
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
      MaxCols         =   9
      MaxRows         =   100
      ProcessTab      =   -1  'True
      SpreadDesigner  =   "frm_Pay_Structure.frx":146A1
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
      Height          =   675
      Left            =   120
      TabIndex        =   9
      Top             =   1230
      Width           =   11910
      Begin VB.TextBox Txtc_Company 
         Height          =   300
         Left            =   1305
         TabIndex        =   6
         Top             =   240
         Width           =   3765
      End
      Begin VB.TextBox Txtc_Remarks 
         Height          =   300
         Left            =   6960
         MaxLength       =   50
         TabIndex        =   7
         Top             =   225
         Width           =   4485
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
         Left            =   6135
         TabIndex        =   13
         Top             =   270
         Width           =   750
      End
      Begin VB.Label Label7 
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
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   450
         TabIndex        =   12
         Top             =   285
         Width           =   780
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
      Left            =   10620
      TabIndex        =   11
      Top             =   960
      Width           =   360
   End
   Begin VB.Label lbl_scr_name 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Pay Structure Details"
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
      TabIndex        =   10
      Top             =   960
      Width           =   1710
   End
   Begin VB.Shape shp_scr_name 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   120
      Top             =   915
      Width           =   11910
   End
End
Attribute VB_Name = "frm_Pay_Struture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As New ADODB.Recordset
 
Private Sub Form_Load()
    lbl_date.Caption = Format(Date, "dd-mmm-yyyy")
    Call Spread_Lock
    Call TGControlProperty(Me)
    Call Spread_Row_Height(Va_Details)
    Call Combo_Load
    Call Clear_Spread(Va_Details)
    Call Create_Default_PayStructure("COM0001")
End Sub

Private Sub Btn_Exit_Click()
    Unload Me
End Sub

Private Sub Btn_Clear_Click()
    Clear_Controls Me
    Clear_Spread Va_Details
End Sub

Private Sub Btn_Delete_Click()
On Error GoTo ErrDel

    If Trim(Txtc_Company) = "" Then
       Exit Sub
    End If
    
    MsgBox "No access to delete.", vbInformation, "Information"
Exit Sub

ErrDel:
    MsgBox "Error while Deleting - " + Err.Description, vbCritical, "Critical"
    CON.RollbackTrans
End Sub

Private Sub Btn_Print_Click()
On Error GoTo Err_Print
  Dim SelFor As String, RepTitle As String
  
   If Trim(Txtc_Company) = "" Then
      Exit Sub
   End If
   
   RepTitle = "Pay Component Structure"
   SelFor = "{PR_PAYSTRUCTURE_MST.C_COMPANY}='" & Right(Trim(Trim(Txtc_Company)), 7) & "'"
   Call Print_Rpt(SelFor, "Pr_PayStructure_List.rpt")
  
   If Trim(RepTitle) <> "" Then
      Mdi_Ta_HrPay.CRY1.Formulas(1) = "ReportHead='" & Trim(RepTitle) & "'"
   End If

   Mdi_Ta_HrPay.CRY1.Action = 1
  
  Exit Sub

Err_Print:
    MsgBox "Error while Generating - " + Err.Description, vbInformation, "Information"
End Sub

Private Sub Btn_View_Click()
  Dim Search As New Search.MyClass, SerVar
    
    Search.Query = "select b.c_companyname Company, a.c_remarks Remarks, a.c_company Code " & _
                   "from pr_paystructure_mst a, pr_company_mst b " & _
                   "where a.c_company = b.c_company and a.c_rec_sta='A'"
    Search.CheckFields = "Company, Code "
    Search.ReturnField = "Company, Code "
    SerVar = Search.Search(, , CON)
    If Len(Search.col1) <> 0 Then
       Txtc_Company = Search.col1 & Space(100) & Search.col2
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
        
        Save_PR_PayStructure_Mst
        Save_PR_PayStructure_Dtl
     
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
  If Trim(Txtc_Company) = "" Then
     MsgBox "Company should not be empty", vbInformation, "Information"
     Txtc_Company.SetFocus
     Exit Function
  ElseIf IsColDupValues(Va_Details, 2) > 0 Then
     MsgBox "Duplicate Salary type found. Please check the entries.", vbInformation, "Information"
     Va_Details.SetFocus
     Exit Function
  ElseIf IsColDupValues(Va_Details, 9) > 0 Then
     MsgBox "Duplicate code found. Please check the entries.", vbInformation, "Information"
     Va_Details.SetFocus
     Exit Function
  End If
  ChkSave = True
End Function

Private Sub Save_PR_PayStructure_Mst()
    Set rs = Nothing
    g_Sql = "Select * from pr_paystructure_mst where c_company = '" & Is_Null(Right(Trim(Txtc_Company), 7), False) & "'"
    rs.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
    If rs.RecordCount = 0 Then
       rs.AddNew
       rs("d_created").Value = GetDateTime
       rs("c_usr_id").Value = g_UserName
    Else
       rs("d_modified").Value = GetDateTime
       rs("c_musr_id").Value = g_UserName
    End If
    
       rs("c_company").Value = Is_Null(Right(Trim(Txtc_Company), 7), False)
       rs("c_remarks").Value = Is_Null(Txtc_Remarks, False)
       rs("c_rec_sta").Value = "A"
       rs.Update

End Sub

Private Sub Save_PR_PayStructure_Dtl()
Dim i As Long

        g_Sql = "delete from pr_paystructure_dtl where c_company = '" & Is_Null(Right(Trim(Txtc_Company), 7), False) & "'"
        CON.Execute g_Sql
        
        Set rs = Nothing
        g_Sql = "Select * from pr_paystructure_dtl where 1=1"
        rs.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
     
        For i = 1 To Va_Details.DataRowCnt
            Va_Details.Row = i
            Va_Details.Col = 2
               If Trim(Va_Details.Text) <> "" Then
                  rs.AddNew
                  Va_Details.Row = i
                     rs("c_company").Value = Is_Null(Right(Trim(Txtc_Company), 7), False)
                     rs("n_seq").Value = i
                  Va_Details.Col = 1
                     Va_Details.TypeComboBoxIndex = Va_Details.TypeComboBoxCurSel
                     rs("c_type").Value = Is_Null(Right(Va_Details.TypeComboBoxString, 1), False)
                  Va_Details.Col = 2
                     rs("c_payname").Value = Is_Null(Va_Details.Text, False)
                  Va_Details.Col = 3
                     rs("c_paye").Value = IIf(Va_Details.Value, "Y", "N")
                  Va_Details.Col = 4
                     rs("c_bonus").Value = IIf(Va_Details.Value, "Y", "N")
                  Va_Details.Col = 5
                     rs("c_master").Value = IIf(Va_Details.Value, "Y", "N")
                  Va_Details.Col = 6
                     rs("c_syscal").Value = IIf(Va_Details.Value, "Y", "N")
                  Va_Details.Col = 7
                     rs("c_category").Value = Is_Null(Va_Details.Text, False)
                  Va_Details.Col = 8
                     rs("c_remarks").Value = Is_Null(Va_Details.Text, False)
                  Va_Details.Col = 9
                     rs("c_salary").Value = Start_Generate_New(Trim(Va_Details.Text))
                  rs.Update
               End If
        Next i
End Sub

Private Function Start_Generate_New(ByVal vSalary As String) As String
  Dim MaxNo As ADODB.Recordset
  
  If Len(vSalary) = 7 Then
     Start_Generate_New = vSalary
  Else
     g_Sql = "Select max(right(c_salary,2)) from pr_paystructure_dtl where c_company = '" & Is_Null(Right(Trim(Txtc_Company), 7), False) & "' " & _
             "and c_salary like 'SAL01%' "
     Set MaxNo = CON.Execute(g_Sql)
     Start_Generate_New = "SAL01" & Format(Is_Null(MaxNo(0).Value, True) + 1, "00")
  End If
End Function

Private Sub Display_Records()
On Error GoTo Err_Display
  Dim DyDisp As New ADODB.Recordset
  Dim i, j As Long
  Dim vType As String
  
  Set DyDisp = Nothing
  g_Sql = "select a.*, b.c_companyname from pr_paystructure_mst a, pr_company_mst b " & _
          "where a.c_company = b.c_company and a.c_company= '" & Is_Null(Right(Trim(Txtc_Company), 7), False) & "' and a.c_rec_sta='A'"
  DyDisp.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
  If DyDisp.RecordCount > 0 Then
     Txtc_Company = Is_Null(DyDisp("c_companyname").Value, False) & Space(100) & Is_Null(DyDisp("c_company").Value, False)
     Txtc_Remarks = Is_Null(DyDisp("c_remarks").Value, False)
  End If
  
  ' // Details
   Set DyDisp = Nothing
   g_Sql = "select * from pr_paystructure_dtl " & _
           "where c_company = '" & Is_Null(Right(Trim(Txtc_Company), 7), False) & "' order by c_type, n_seq "
   DyDisp.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
   
   Va_Details.MaxRows = DyDisp.RecordCount + 10
    
   If DyDisp.RecordCount > 0 Then
      DyDisp.MoveFirst
      For i = 1 To DyDisp.RecordCount
          Va_Details.Row = i
          Va_Details.Col = 1
             If Is_Null(DyDisp("c_type").Value, False) = "1" Then
                vType = "Income" & Space(50) & "1"
             ElseIf Is_Null(DyDisp("c_type").Value, False) = "2" Then
                vType = "Deduction" & Space(50) & "2"
             ElseIf Is_Null(DyDisp("c_type").Value, False) = "3" Then
                vType = "Company Contribution" & Space(50) & "3"
             Else
                vType = ""
             End If
            j = SelComboString(Va_Details, vType, i, 1)
            Va_Details.TypeComboBoxIndex = j
          Va_Details.Col = 2
             Va_Details.Text = Is_Null(DyDisp("c_payname").Value, False)
          Va_Details.Col = 3
             Va_Details.Text = IIf(DyDisp("c_paye").Value = "Y", True, False)
          Va_Details.Col = 4
             Va_Details.Text = IIf(DyDisp("c_bonus").Value = "Y", True, False)
          Va_Details.Col = 5
             Va_Details.Text = IIf(DyDisp("c_master").Value = "Y", True, False)
          Va_Details.Col = 6
             Va_Details.Text = IIf(DyDisp("c_syscal").Value = "Y", True, False)
          Va_Details.Col = 7
             Va_Details.Text = Is_Null(DyDisp("c_category").Value, False)
          Va_Details.Col = 8
             Va_Details.Text = Is_Null(DyDisp("c_remarks").Value, False)
          Va_Details.Col = 9
             Va_Details.Text = Is_Null(DyDisp("c_salary").Value, False)
          DyDisp.MoveNext
      Next i
   End If
   Call Spread_Row_Height(Va_Details)

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

Private Sub Txtc_Company_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim Search As New Search.MyClass, SerVar
  If KeyCode = vbKeyF2 Then
                   
     Search.Query = "select c_companyname CompanyName, c_company Code from pr_company_mst where c_rec_sta='A' "
     Search.CheckFields = "CompanyName, Code"
     Search.ReturnField = "CompanyName, Code"
     SerVar = Search.Search(, , CON)
     If Len(Search.col1) <> 0 Then
        Txtc_Company.Text = Search.col1 & Space(100) & Search.col2
     End If
  End If
End Sub

Private Sub Txtc_Company_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Txtc_Company_Validate(Cancel As Boolean)
  Dim RsChk As New ADODB.Recordset
    If Trim(Txtc_Company) <> "" Then
       Set RsChk = Nothing
       g_Sql = "select * from pr_company_mst where c_company = '" & Right(Trim(Txtc_Company), 7) & "'"
       RsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
       If RsChk.RecordCount = 0 Then
          MsgBox "Company Name not found. Press <F2> to Select", vbInformation, "Information"
          Txtc_Company.SetFocus
          Cancel = True
       Else
          Txtc_Company = Is_Null(RsChk("c_companyname").Value, False) & Space(100) & Is_Null(RsChk("c_company").Value, False)
          Call Create_Default_PayStructure(Is_Null(RsChk("c_company").Value, False))
          Call Display_Records
       End If
    End If
End Sub

Private Sub Va_Details_KeyDown(KeyCode As Integer, Shift As Integer)
  
  If ((Shift And 1) = 1) And KeyCode = vbKeyInsert Then
     Call SpreadInsertRow(Va_Details, Va_Details.ActiveRow)
     Call Spread_Row_Height(Va_Details)
  
  ElseIf ((Shift And 1) = 1) And KeyCode = vbKeyDelete And g_Admin Then
     If MsgBox("Do you want to Delete", vbYesNo, "Confirmation") = vbYes Then
        Call SpreadDeleteRow(Va_Details, Va_Details.ActiveRow)
        Call Spread_Row_Height(Va_Details)
     End If
  End If
End Sub

Private Sub Va_Details_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
  Dim vPayCode As String
    vPayCode = GetNewPayCode()
    
    Va_Details.Row = Row
    Va_Details.Col = 2
       If Trim(Va_Details.Text) <> "" Then
          Va_Details.Col = 9
             If Trim(Va_Details.Text) = "" Then
                Va_Details.Text = vPayCode
             End If
       End If
End Sub

Private Sub Combo_Load()
  Dim Str As String
    
    Str = "Income" & Space(50) & "1"
    Str = Str & Chr$(9) & "Deduction" & Space(50) & "2"
    Str = Str & Chr$(9) & "Company Contribution" & Space(50) & "3"
   
    Va_Details.Row = -1
    Va_Details.Col = 1
    Va_Details.TypeComboBoxList = Str

End Sub

Private Sub Spread_Lock()
  Dim i As Integer
    
    For i = 1 To Va_Details.MaxCols
        Va_Details.Row = -1
        Va_Details.Col = i
           If i = 9 Then
              Va_Details.Lock = True
           Else
              Va_Details.Lock = False
           End If
    Next i
End Sub

Private Function GetNewPayCode() As String
  Dim i As Integer
  Dim MaxNo As Integer, vNo As Integer
  
    MaxNo = 0
  
    For i = 1 To Va_Details.DataRowCnt
        Va_Details.Row = i
        Va_Details.Col = 9
           vNo = Is_Null(Right(Trim(Va_Details.Text), 4), True)
           If vNo > MaxNo Then
              MaxNo = vNo
           End If
    Next i
    If MaxNo < 100 Then
       MaxNo = 100
    End If
    GetNewPayCode = "SAL" & Format(MaxNo + 1, "0000")
End Function
