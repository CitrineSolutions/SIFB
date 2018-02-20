VERSION 5.00
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Begin VB.Form frm_Shift_Structure 
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
      TabIndex        =   14
      Top             =   -45
      Width           =   5535
      Begin VB.CommandButton Btn_View 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   1005
         Picture         =   "frm_Shift_Structure.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Save 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   30
         Picture         =   "frm_Shift_Structure.frx":36B0
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Exit 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   4470
         Picture         =   "frm_Shift_Structure.frx":6D39
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Clear 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   1725
         Picture         =   "frm_Shift_Structure.frx":A399
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Delete 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   2715
         Picture         =   "frm_Shift_Structure.frx":DA09
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Print 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   3435
         Picture         =   "frm_Shift_Structure.frx":110B3
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
      Begin FPUSpreadADO.fpSpread Va_Details 
         Height          =   3225
         Left            =   285
         TabIndex        =   8
         Top             =   1080
         Width           =   4770
         _Version        =   458752
         _ExtentX        =   8414
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
         SpreadDesigner  =   "frm_Shift_Structure.frx":146A1
         VisibleCols     =   1
      End
      Begin VB.TextBox Txtc_ShiftName 
         Height          =   300
         Left            =   1380
         MaxLength       =   25
         TabIndex        =   7
         Top             =   540
         Width           =   3720
      End
      Begin VB.TextBox Txtc_Code 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   1380
         TabIndex        =   6
         Top             =   240
         Width           =   885
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
         Caption         =   "Shift Name"
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
         Left            =   375
         TabIndex        =   13
         Top             =   585
         Width           =   885
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
         ForeColor       =   &H00404080&
         Height          =   210
         Left            =   825
         TabIndex        =   12
         Top             =   285
         Width           =   435
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
      Caption         =   "Shift Structure Details"
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
      Width           =   1800
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
Attribute VB_Name = "frm_Shift_Structure"
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
    Call Clear_Spread(Va_Details)
    Call Create_Default_ShiftStructure
End Sub

Private Sub Form_Activate()
    Txtc_Code.Enabled = False
    Call Display_Week
End Sub

Private Sub Btn_Exit_Click()
    Unload Me
End Sub

Private Sub Btn_Clear_Click()
    Clear_Controls Me
    Clear_Spread Va_Details
    Call Display_Week
    Txtc_Code.Enabled = False
End Sub

Private Sub Btn_Delete_Click()
On Error GoTo ErrDel
        
    If Trim(Txtc_Code) = "" Then
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
  
   If Trim(Txtc_Code) = "" Then
      Exit Sub
   End If
   
   RepTitle = "Shift Structure"
   SelFor = "{PR_SHIFTSTRUCTURE_MST.C_CODE}='" & Trim(Txtc_Code) & "'"
   Call Print_Rpt(SelFor, "Pr_ShiftStructure_List.rpt")
  
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

    Search.Query = "select c_shiftname ShiftName, c_code Code " & _
                   "from pr_shiftstructure_mst where c_rec_sta='A'"
    Search.CheckFields = "ShiftName, Code"
    Search.ReturnField = "ShiftName, Code"
    SerVar = Search.Search(, , CON)
    If Len(Search.col2) <> 0 Then
        Txtc_Code = Search.col2
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
        
        Save_PR_ShiftStructure_Mst
        Save_PR_ShiftStructure_Dtl
        
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
  
  If Trim(Txtc_ShiftName) = "" Then
     MsgBox "Shift Name should not be empty", vbInformation, "Information"
     Txtc_ShiftName.SetFocus
     Exit Function
  End If
    
  For i = 1 To 7
      Va_Details.Row = i
      Va_Details.Col = 2
         If Trim(Va_Details.Text) = "" Then
            MsgBox "Shift should not be empty.", vbInformation, "Information"
            Va_Details.SetFocus
            Exit Function
         End If
  Next i
   
  ChkSave = True
End Function

Private Sub Save_PR_ShiftStructure_Mst()
    Set rs = Nothing
    g_Sql = "Select * from pr_shiftstructure_mst where c_code = '" & Trim(Txtc_Code) & "' and c_rec_sta='A'"
    rs.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
    
    If rs.RecordCount = 0 Then
        rs.AddNew
        Call Start_Generate_New
        rs("c_usr_id").Value = g_UserName
        rs("d_created").Value = GetDateTime
    Else
       rs("c_musr_id").Value = g_UserName
       rs("d_modified").Value = GetDateTime
    End If
    
    rs("c_code").Value = Is_Null(Txtc_Code, False)
    rs("c_shiftname").Value = Is_Null(Txtc_ShiftName, False)
    rs("c_rec_sta").Value = "A"
    rs.Update
End Sub

Private Sub Save_PR_ShiftStructure_Dtl()
Dim i As Long

  g_Sql = "delete from pr_shiftstructure_dtl where c_Code = '" & Trim(Txtc_Code) & "'"
  CON.Execute (g_Sql)
  
  Set rs = Nothing
  g_Sql = "Select * from pr_shiftstructure_dtl where 1=2"
  rs.Open g_Sql, CON, adOpenDynamic, adLockOptimistic

  For i = 1 To 7
      rs.AddNew
      rs("c_code").Value = Is_Null(Txtc_Code, False)
      rs("n_slno").Value = Is_Null(i, True)
      
      Va_Details.Row = i
      Va_Details.Col = 1
         rs("n_wkday").Value = Is_Null(Right(Trim(Va_Details.Text), 1), True)
      Va_Details.Col = 2
         rs("c_shiftcode").Value = Is_Null(Va_Details.Text, False)
      Va_Details.Col = 3
         rs("c_wocode").Value = Is_Null(Va_Details.Text, False)
      Va_Details.Col = 4
         rs("c_altcode").Value = Is_Null(Va_Details.Text, False)
      
      rs.Update
  Next i
  
End Sub

Private Sub Start_Generate_New()
  Dim MaxNo As ADODB.Recordset
  g_Sql = "Select max(right(c_code,2)) from pr_shiftstructure_mst "
  Set MaxNo = CON.Execute(g_Sql)
  Txtc_Code = "S" & Format(Is_Null(MaxNo(0).Value, True) + 1, "00")
End Sub


Private Sub Display_Records()
On Error GoTo Err_Display
  Dim DyDisp As New ADODB.Recordset
  Dim i, j As Long
  Dim vType As String
  
  Set DyDisp = Nothing
  g_Sql = "select * from pr_shiftstructure_mst where c_code = '" & Trim(Txtc_Code) & "'"
  DyDisp.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
  
  Txtc_Code = Is_Null(DyDisp("c_code").Value, False)
  Txtc_ShiftName = Is_Null(DyDisp("c_shiftname").Value, False)
  
  ' // Details
    
   Set DyDisp = Nothing
   g_Sql = "select * from pr_shiftstructure_dtl where c_code = '" & Trim(Txtc_Code) & "' " & _
           "order by n_slno "
   DyDisp.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    
   If DyDisp.RecordCount > 0 Then
      DyDisp.MoveFirst
      Call Display_Week
      For i = 1 To DyDisp.RecordCount
          Va_Details.Row = i
          Va_Details.Col = 2
             Va_Details.Text = Is_Null(DyDisp("c_shiftcode").Value, False)
          Va_Details.Col = 3
             Va_Details.Text = Is_Null(DyDisp("c_wocode").Value, False)
          Va_Details.Col = 4
             Va_Details.Text = Is_Null(DyDisp("c_altcode").Value, False)
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

Private Sub Display_Week()
  Dim i As Integer
  
      Va_Details.Col = 1
      Va_Details.Row = 1
         Va_Details.Text = "Monday" & Space(100) & "1"
      Va_Details.Row = 2
         Va_Details.Text = "Tuesday" & Space(100) & "2"
      Va_Details.Row = 3
         Va_Details.Text = "Wednesday" & Space(100) & "3"
      Va_Details.Row = 4
         Va_Details.Text = "Thursday" & Space(100) & "4"
      Va_Details.Row = 5
         Va_Details.Text = "Friday" & Space(100) & "5"
      Va_Details.Row = 6
         Va_Details.Text = "Saturday" & Space(100) & "6"
      Va_Details.Row = 7
         Va_Details.Text = "Sunday" & Space(100) & "7"
End Sub

Private Sub Spread_Lock()
  Dim i As Integer
      For i = 1 To Va_Details.MaxCols
          Va_Details.Row = -1
          Va_Details.Col = i
          If i = 1 Then
             Va_Details.Lock = True
          Else
             Va_Details.Lock = False
          End If
      Next i
End Sub

Private Sub Va_Details_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    Call SpreadBlockCopy(Va_Details, BlockCol, BlockRow, BlockCol2, BlockRow2, False, False, True)
End Sub

Private Sub Va_Details_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim Search As New Search.MyClass, SerVar
    
    If KeyCode = vbKeyDelete Then
       Call SpreadCellDataClear(Va_Details, Va_Details.ActiveRow, Va_Details.ActiveCol)
    End If
    
    If Va_Details.ActiveCol <> 1 And KeyCode = vbKeyF2 Then
       Search.Query = "select c_shiftcode Shift, n_starthrs StartTime, n_endhrs EndTime, n_shifthrs WorkHrs " & _
                      "from pr_clock_shift "
       Search.CheckFields = "Shift"
       Search.ReturnField = "Shift"
       SerVar = Search.Search(, , CON)
       If Len(Search.col1) <> 0 Then
          Va_Details.Row = Va_Details.ActiveRow
          Va_Details.Col = Va_Details.ActiveCol
             Va_Details.Text = Search.col1
       End If
    End If
End Sub

Private Sub Va_Details_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
  Dim rsChk As New ADODB.Recordset
    
    If Col = 2 Or Col = 3 Or Col = 4 Then
       Va_Details.Row = Row
       Va_Details.Col = Col
          If Trim(Va_Details.Text) <> "" Then
             Set rsChk = Nothing
             g_Sql = "select c_shiftcode from pr_clock_shift where c_shiftcode = '" & Trim(Va_Details.Text) & "'"
             rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
             If rsChk.RecordCount > 0 Then
                Va_Details.Text = Is_Null(rsChk("c_shiftcode").Value, False)
             Else
                Va_Details.Text = ""
                MsgBox "Shift not found in the Master. Press <F2> to select", vbInformation, "Information"
                Cancel = True
             End If
          End If
    End If
 
End Sub


