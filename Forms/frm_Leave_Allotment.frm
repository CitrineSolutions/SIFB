VERSION 5.00
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Begin VB.Form frm_Leave_Allotment 
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6825
   ScaleWidth      =   8925
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   900
      Left            =   120
      TabIndex        =   13
      Top             =   -45
      Width           =   5535
      Begin VB.CommandButton Btn_View 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   1005
         Picture         =   "frm_Leave_Allotment.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Save 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   30
         Picture         =   "frm_Leave_Allotment.frx":36B0
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Exit 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   4470
         Picture         =   "frm_Leave_Allotment.frx":6D39
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Clear 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   1725
         Picture         =   "frm_Leave_Allotment.frx":A399
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Delete 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   2715
         Picture         =   "frm_Leave_Allotment.frx":DA09
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Print 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   3435
         Picture         =   "frm_Leave_Allotment.frx":110B3
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   700
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
      Height          =   4485
      Left            =   120
      TabIndex        =   11
      Top             =   1260
      Width           =   5535
      Begin VB.TextBox Txtn_YearFrom 
         Height          =   300
         Left            =   1905
         TabIndex        =   7
         Top             =   615
         Width           =   915
      End
      Begin FPUSpreadADO.fpSpread Va_Details 
         Height          =   3225
         Left            =   660
         TabIndex        =   8
         Top             =   1080
         Width           =   4320
         _Version        =   458752
         _ExtentX        =   7620
         _ExtentY        =   5689
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
         MaxCols         =   4
         MaxRows         =   7
         ProcessTab      =   -1  'True
         ScrollBars      =   0
         SpreadDesigner  =   "frm_Leave_Allotment.frx":146A1
         VisibleCols     =   1
      End
      Begin VB.TextBox Txtc_LeaveName 
         Height          =   300
         Left            =   1905
         TabIndex        =   6
         Top             =   315
         Width           =   3435
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Year Effective From"
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
         Left            =   210
         TabIndex        =   14
         Top             =   675
         Width           =   1605
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   -750
         X2              =   5415
         Y1              =   975
         Y2              =   960
      End
      Begin VB.Label Label9 
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
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   1305
         TabIndex        =   12
         Top             =   360
         Width           =   495
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
      Left            =   4140
      TabIndex        =   10
      Top             =   1005
      Width           =   360
   End
   Begin VB.Label lbl_scr_name 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Leave Entitle Slabs"
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
      TabIndex        =   9
      Top             =   1005
      Width           =   1545
   End
   Begin VB.Shape shp_scr_name 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   120
      Top             =   960
      Width           =   5535
   End
End
Attribute VB_Name = "frm_Leave_Allotment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As New ADODB.Recordset

Private Sub Form_Load()
    lbl_date.Caption = Format(Date, "dd-mmm-yyyy")
    Call TGControlProperty(Me)
    Call Spread_Row_Height(Va_Details, , 30)
    Call Clear_Spread(Va_Details)
    Call Create_Default_LeaveAllot
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
       
    MsgBox "No access to delete.", vbInformation, "Information"
Exit Sub

ErrDel:
    MsgBox "Error while Deleting - " + Err.Description, vbCritical, "Critical"
    CON.RollbackTrans
End Sub

Private Sub Btn_Print_Click()
On Error GoTo Err_Print
  Dim SelFor As String, RepTitle As String
  
   If Trim(Txtc_LeaveName) = "" Or Trim(Txtn_YearFrom) = "" Then
      Exit Sub
   End If
   
   RepTitle = "Leave Allotment Details"
   SelFor = "{PR_LEAVEALLOT_MST.C_LEAVE}='" & Trim(Right(Trim(Txtc_LeaveName), 7)) & "' AND {PR_LEAVEALLOT_MST.N_YEARFROM}= " & Val(Txtn_YearFrom)
   Call Print_Rpt(SelFor, "Pr_LeaveAllot_List.rpt")
  
   If Trim(RepTitle) <> "" Then
      Mdi_Ta_HrPay.CRY1.Formulas(1) = "ReportHead='" & UCase(Trim(RepTitle)) & "'"
   End If

   Mdi_Ta_HrPay.CRY1.Action = 1
  
  Exit Sub

Err_Print:
    MsgBox "Error while Generating - " + Err.Description, vbInformation, "Information"
End Sub

Private Sub Btn_View_Click()
  Dim Search As New Search.MyClass, SerVar

    Search.Query = "select b.c_leavename Leave, a.n_yearfrom YearFrom, a.c_leave Code " & _
                   "from pr_leaveallot_mst a, pr_leave_mst b where a.c_leave = b.c_leave and a.c_rec_sta='A'"
    Search.CheckFields = "YearFrom, Code"
    Search.ReturnField = "YearFrom, Code"
    SerVar = Search.Search(, , CON)
    If Len(Search.col2) <> 0 Then
        Txtc_LeaveName = Search.col2
        Txtn_YearFrom = Search.col1
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
        
        Save_Pr_LeaveAllot_Mst
        Save_Pr_LeaveAllot_Dtl
        
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
  
  If Trim(Txtc_LeaveName) = "" Then
     MsgBox "Leave should not be empty", vbInformation, "Information"
     Txtc_LeaveName.SetFocus
     Exit Function
  
  ElseIf Trim(Txtn_YearFrom) = "" Then
     MsgBox "Year From should not be empty", vbInformation, "Information"
     Txtn_YearFrom.SetFocus
     Exit Function
  End If
  
  ChkSave = True
End Function

Private Sub Save_Pr_LeaveAllot_Mst()
    Set rs = Nothing
    g_Sql = "Select * from pr_leaveallot_mst where c_leave = '" & Trim(Right(Trim(Txtc_LeaveName), 7)) & "' and " & _
            "n_yearfrom = " & Is_Null(Txtn_YearFrom, True) & " and c_rec_sta='A'"
    rs.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
    
    If rs.RecordCount = 0 Then
        rs.AddNew
        rs("c_usr_id").Value = g_UserName
        rs("d_created").Value = GetDateTime
    Else
       rs("c_musr_id").Value = g_UserName
       rs("d_modified").Value = GetDateTime
    End If
    
    rs("c_leave").Value = Is_Null(Right(Trim(Txtc_LeaveName), 7), False)
    rs("n_yearfrom").Value = Is_Null(Txtn_YearFrom, True)
    rs("c_rec_sta").Value = "A"
    rs.Update
End Sub

Private Sub Save_Pr_LeaveAllot_Dtl()
Dim i As Long, tmp As Long

  g_Sql = "delete from pr_leaveallot_dtl where c_leave = '" & Trim(Right(Trim(Txtc_LeaveName), 7)) & "' and " & _
          "n_yearfrom = " & Is_Null(Txtn_YearFrom, True)
  CON.Execute (g_Sql)
  
  Set rs = Nothing
  g_Sql = "Select * from pr_leaveallot_dtl where 1=2"
  rs.Open g_Sql, CON, adOpenDynamic, adLockOptimistic

  For i = 1 To Va_Details.DataRowCnt
      
      Va_Details.Row = i
      Va_Details.Col = 1
         tmp = Is_Null(Va_Details.Text, True)
      Va_Details.Col = 2
         tmp = tmp + Is_Null(Va_Details.Text, True)
         
      If tmp > 0 Then
         rs.AddNew
         rs("c_leave").Value = Is_Null(Right(Trim(Txtc_LeaveName), 7), False)
         rs("n_yearfrom").Value = Is_Null(Txtn_YearFrom, True)
        
         Va_Details.Row = i
         Va_Details.Col = 1
            rs("n_from").Value = Is_Null(Va_Details.Text, True)
         Va_Details.Col = 2
            rs("n_to").Value = Is_Null(Va_Details.Text, True)
         Va_Details.Col = 3
            rs("n_allot").Value = Is_Null(Va_Details.Text, True)
         Va_Details.Col = 4
            rs("n_max").Value = Is_Null(Va_Details.Text, True)
         
         rs.Update
      End If
      
  Next i
  
End Sub

Private Sub Display_Records()
On Error GoTo Err_Display
  Dim DyDisp As New ADODB.Recordset
  Dim i, j As Long
  Dim vType As String
  
    Set DyDisp = Nothing
    g_Sql = "select * from pr_leaveallot_mst where c_leave = '" & Trim(Right(Trim(Txtc_LeaveName), 7)) & "' and " & _
            "n_yearfrom = " & Is_Null(Txtn_YearFrom, True) & " and c_rec_sta='A'"
    DyDisp.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    
    Txtc_LeaveName = Is_Null(DyDisp("c_leave").Value, False)
    Txtn_YearFrom = Is_Null(DyDisp("n_yearfrom").Value, True)
    Call Txtc_LeaveName_Validate(True)
    
    Call Get_Leave_Allot
    
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


Private Sub Txtc_LeaveName_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim Search As New Search.MyClass, SerVar
    If KeyCode = vbKeyDelete Then
       Txtc_LeaveName = ""
    End If
   
    If KeyCode = vbKeyF2 Then
       Search.Query = "select c_leavename Leave, c_leave Code " & _
                      "from pr_leave_mst "
       Search.CheckFields = "Leave, Code"
       Search.ReturnField = "Leave, Code"
       SerVar = Search.Search(, , CON)
       If Len(Search.col1) <> 0 Then
          Txtc_LeaveName = Search.col1 & Space(100) & Search.col2
       End If
    End If
End Sub

Private Sub Txtc_LeaveName_Validate(Cancel As Boolean)
  Dim rsChk As New ADODB.Recordset
  
    If Trim(Txtc_LeaveName) <> "" Then
       Set rsChk = Nothing
       g_Sql = "select * from pr_leave_mst where c_leave = '" & Trim(Right(Trim(Txtc_LeaveName.Text), 7)) & "'"
       rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
       If rsChk.RecordCount > 0 Then
          Txtc_LeaveName = Is_Null(rsChk("c_leavename").Value, False) & Space(100) & Is_Null(rsChk("c_leave").Value, False)
          Call Get_Leave_Allot
       Else
          Txtc_LeaveName.SetFocus
          MsgBox "Leave details not found in the Master. Press <F2> to select", vbInformation, "Information"
          Cancel = True
       End If
    End If
End Sub

Private Sub Get_Leave_Allot()
  Dim DyDisp As New ADODB.Recordset
  Dim i As Integer
  
     If Trim(Txtc_LeaveName) <> "" And Trim(Txtn_YearFrom) <> "" Then
        Call Clear_Spread(Va_Details)
        Set DyDisp = Nothing
        g_Sql = "select * from pr_leaveallot_dtl where c_leave = '" & Trim(Right(Trim(Txtc_LeaveName), 7)) & "' and " & _
                "n_yearfrom = " & Is_Null(Txtn_YearFrom, True) & "  " & _
                "order by n_from, n_to "
        DyDisp.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
         
        If DyDisp.RecordCount > 0 Then
           DyDisp.MoveFirst
           For i = 1 To DyDisp.RecordCount
               Va_Details.Row = i
               Va_Details.Col = 1
                  Va_Details.Text = Is_Null(DyDisp("n_from").Value, True)
               Va_Details.Col = 2
                  Va_Details.Text = Is_Null(DyDisp("n_to").Value, True)
               Va_Details.Col = 3
                  Va_Details.Text = Is_Null(DyDisp("n_allot").Value, True)
               Va_Details.Col = 4
                  Va_Details.Text = Is_Null(DyDisp("n_max").Value, True)
               DyDisp.MoveNext
           Next i
        End If
     End If
End Sub


Private Sub Txtn_YearFrom_Validate(Cancel As Boolean)
    Call Get_Leave_Allot
End Sub

Private Sub Va_Details_KeyDown(KeyCode As Integer, Shift As Integer)
    If ((Shift And 1) = 1) And KeyCode = vbKeyInsert Then
        Call SpreadInsertRow(Va_Details, Va_Details.ActiveRow)
        Call Spread_Row_Height(Va_Details, , 30)
    ElseIf ((Shift And 1) = 1) And KeyCode = vbKeyDelete And g_Admin Then
        If MsgBox("Do you want to Delete", vbYesNo, "Confirmation") = vbYes Then
            Call SpreadDeleteRow(Va_Details, Va_Details.ActiveRow)
            Call Spread_Row_Height(Va_Details, , 30)
        End If
    ElseIf KeyCode = vbKeyDelete Then
        Call SpreadCellDataClear(Va_Details, Va_Details.ActiveRow, Va_Details.ActiveCol)
    End If
End Sub
