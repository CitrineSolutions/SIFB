VERSION 5.00
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Begin VB.Form frm_PayPeriod 
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
      TabIndex        =   16
      Top             =   -45
      Width           =   8055
      Begin VB.CommandButton Btn_View 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   1005
         Picture         =   "frm_PayPeriod.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Save 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   30
         Picture         =   "frm_PayPeriod.frx":36B0
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Exit 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   4470
         Picture         =   "frm_PayPeriod.frx":6D39
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Clear 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   1725
         Picture         =   "frm_PayPeriod.frx":A399
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Delete 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   2715
         Picture         =   "frm_PayPeriod.frx":DA09
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Print 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   3435
         Picture         =   "frm_PayPeriod.frx":110B3
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   700
      End
   End
   Begin FPUSpreadADO.fpSpread Va_Pay 
      Height          =   5760
      Left            =   120
      TabIndex        =   9
      Top             =   2190
      Width           =   8055
      _Version        =   458752
      _ExtentX        =   14208
      _ExtentY        =   10160
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
      MaxCols         =   7
      MaxRows         =   13
      ProcessTab      =   -1  'True
      SpreadDesigner  =   "frm_PayPeriod.frx":146A1
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
      Height          =   945
      Left            =   120
      TabIndex        =   12
      Top             =   1230
      Width           =   8070
      Begin VB.ComboBox Cmb_Type 
         Height          =   315
         Left            =   3990
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   195
         Width           =   1725
      End
      Begin VB.TextBox Txtc_Remarks 
         Height          =   300
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   8
         Top             =   540
         Width           =   4635
      End
      Begin VB.TextBox Txtn_Year 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   6
         Top             =   202
         Width           =   1005
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         Caption         =   "Period For"
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
         Left            =   2805
         TabIndex        =   15
         Top             =   247
         Width           =   855
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
         Left            =   255
         TabIndex        =   14
         Top             =   585
         Width           =   750
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Year"
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
         Left            =   630
         TabIndex        =   13
         Top             =   247
         Width           =   375
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
      Height          =   180
      Left            =   6525
      TabIndex        =   11
      Top             =   960
      Width           =   360
   End
   Begin VB.Label lbl_scr_name 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Salary/Wages Calculation Period"
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
      Top             =   975
      Width           =   2625
   End
   Begin VB.Shape shp_scr_name 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   120
      Top             =   915
      Width           =   8055
   End
End
Attribute VB_Name = "frm_PayPeriod"
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
    Call Spread_Row_Height(Va_Pay)
    Call Clear_Spread(Va_Pay)
    Call Create_Default_PayPeriods(Year(g_CurrentDate))
    
    Cmb_Type.Clear
    Cmb_Type.AddItem "General " & Space(100) & "W"
End Sub

Private Sub Form_Activate()
    Cmb_Type.ListIndex = 0
End Sub

Private Sub Btn_Exit_Click()
    Unload Me
End Sub

Private Sub Btn_Clear_Click()
    Clear_Controls Me
    Clear_Spread Va_Pay
End Sub

Private Sub Btn_Delete_Click()
On Error GoTo ErrDel

    If Trim(Txtn_Year) = "" Then
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
  
   If Val(Txtn_Year) = 0 Then
      Exit Sub
   End If
   
   RepTitle = "Pay Period"
   SelFor = "{PR_PAYPERIOD_MST.N_YEAR}=" & Val(Txtn_Year)
   Call Print_Rpt(SelFor, "Pr_PayPeriod_List.rpt")
  
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
  
    Search.Query = "select n_year Year, c_remarks Remarks " & _
                   "from pr_payperiod_mst where c_rec_sta='A' "
    Search.CheckFields = "Year"
    Search.ReturnField = "Year"
    SerVar = Search.Search(, , CON)
    If Len(Search.col1) <> 0 Then
        Txtn_Year = Search.col1
        Call Display_Records
        Call Va_Pay_LeaveCell(3, 1, 3, 2, False)
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
        
        Save_Pr_PayPeriod_Mst
        Save_PR_PayPeriod_Dtl
    
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
  
  If Val(Txtn_Year) = 0 Then
     MsgBox "Please enter valid year", vbInformation, "Information"
     Txtn_Year.SetFocus
     Exit Function
  End If
    
  For i = 1 To 2
      Va_Pay.Row = i
      Va_Pay.Col = 2
         If Not IsDate(Va_Pay.Text) Then
            MsgBox "Date (From Date) is not a valid date", vbInformation, "Information"
            Va_Pay.SetFocus
            Exit Function
         End If
      Va_Pay.Col = 3
         If Not IsDate(Va_Pay.Text) Then
            MsgBox "Date (To Date) is not a valid date", vbInformation, "Information"
            Va_Pay.SetFocus
            Exit Function
         End If
  Next i
   
  ChkSave = True
End Function

Private Sub Save_Pr_PayPeriod_Mst()
    
    Set rs = Nothing
    g_Sql = "Select * from pr_payperiod_mst where n_year = '" & Is_Null(Txtn_Year, True) & "' and " & _
               "c_type = '" & Right(Trim(Cmb_Type), 1) & "' and  c_rec_sta='A' "
    rs.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
       
    If rs.RecordCount = 0 Then
       rs.AddNew
       rs("d_created").Value = GetDateTime
       rs("c_usr_id").Value = g_UserName
    Else
       rs("d_modified").Value = GetDateTime
       rs("c_musr_id").Value = g_UserName
    End If

       rs("n_year").Value = Is_Null(Txtn_Year, True)
       rs("c_type").Value = Is_Null(Right(Trim(Cmb_Type), 1), False)
       rs("d_fystart").Value = Trim(Str(Is_Null(Txtn_Year, True) - 1)) & "-07-01"
       rs("d_fyend").Value = Trim(Str(Is_Null(Txtn_Year, True))) & "-06-30"
       rs("c_remarks").Value = Is_Null(Txtc_Remarks, False)
       rs("c_rec_sta").Value = "A"
       rs.Update
End Sub

Private Sub Save_PR_PayPeriod_Dtl()
Dim i As Long
       
       g_Sql = "delete from pr_payperiod_dtl where n_year = " & Is_Null(Txtn_Year, True) & " and " & _
               "c_type = '" & Right(Trim(Cmb_Type), 1) & "'"
       CON.Execute (g_Sql)

       Set rs = Nothing
       g_Sql = "Select * from pr_payperiod_dtl where 1=2"
       rs.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
            
       For i = 1 To 13
           rs.AddNew
           Va_Pay.Row = i
           rs("n_year").Value = Is_Null(Txtn_Year, True)
           rs("c_type").Value = Is_Null(Right(Trim(Cmb_Type), 1), False)
           Va_Pay.Col = 1
              rs("n_period").Value = Val(Format(Txtn_Year, "0000") & Right(Trim(Va_Pay.Text), 2))
           Va_Pay.Col = 2
              rs("d_fromdate").Value = Is_Date(Va_Pay.Text, "S")
           Va_Pay.Col = 3
              rs("d_todate").Value = Is_Date(Va_Pay.Text, "S")
           Va_Pay.Col = 4
              rs("n_days").Value = Is_Null(Va_Pay.Text, True)
           Va_Pay.Col = 5
              rs("n_weeks").Value = Is_Null(Va_Pay.Text, True)
           Va_Pay.Col = 6
              rs("c_notes").Value = Is_Null(Va_Pay.Text, False)
           Va_Pay.Col = 7
              If Trim(Va_Pay.Text) = "" Then
                 rs("c_period_closed").Value = "N"
              Else
                 rs("c_period_closed").Value = Is_Null(Va_Pay.Text, False)
              End If
        rs.Update
       Next i
       
End Sub

Private Sub Display_Records()
On Error GoTo Err_Display
  Dim DyDisp As New ADODB.Recordset
  Dim i, j As Long
  Dim vType As String
  
  Set DyDisp = Nothing
  Cmb_Type.ListIndex = 0
  g_Sql = "select * from pr_payperiod_mst where n_year = " & Is_Null(Txtn_Year, True) & " and " & _
          "c_type = '" & Trim(Right(Trim(Cmb_Type), 1)) & "' and  c_rec_sta='A'"
  DyDisp.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
  
  Txtn_Year = Is_Null(DyDisp("n_year").Value, True)
  For i = 0 To Cmb_Type.ListCount - 1
      If Trim(Right(Cmb_Type.List(i), 1)) = DyDisp("c_type").Value Then
         Cmb_Type.ListIndex = i
         Exit For
      End If
  Next i
  Txtc_Remarks = Is_Null(DyDisp("c_remarks").Value, False)
  
  ' // Details
    
   Set DyDisp = Nothing
   g_Sql = "select * from pr_payperiod_dtl where n_year = " & Is_Null(Txtn_Year, True) & " and " & _
           "c_type = '" & Right(Trim(Cmb_Type), 1) & "' order by n_period "
   DyDisp.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    
   If DyDisp.RecordCount > 0 Then
      DyDisp.MoveFirst
      Call Display_Month
      For i = 1 To DyDisp.RecordCount
          Va_Pay.Row = i
          Va_Pay.Col = 2
             Va_Pay.Text = Is_DateSpread(DyDisp("d_fromdate").Value, False)
          Va_Pay.Col = 3
             Va_Pay.Text = Is_DateSpread(DyDisp("d_todate").Value, False)
          Va_Pay.Col = 4
             Va_Pay.Text = Is_Null(DyDisp("n_days").Value, True)
          Va_Pay.Col = 5
             Va_Pay.Text = Is_Null(DyDisp("n_weeks").Value, True)
          Va_Pay.Col = 6
             Va_Pay.Text = Is_Null(DyDisp("c_notes").Value, False)
          Va_Pay.Col = 7
             Va_Pay.Text = Is_Null(DyDisp("c_period_closed").Value, False)
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

Private Sub Txtn_Year_KeyPress(KeyAscii As Integer)
    Call OnlyNumeric(Txtn_Year, KeyAscii, 4)
End Sub

Private Sub Txtn_Year_Validate(Cancel As Boolean)
     If Trim(Txtn_Year) = "" Then
        Exit Sub
     End If
     
     If Val(Txtn_Year) < 2016 Or Val(Txtn_Year) > 2045 Then
        MsgBox "Year must be between 2016 and 2045", vbInformation, "Information"
        Txtn_Year.SetFocus
        Cancel = True
     Else
        Call Display_Month
        Call Create_Default_PayPeriods(Val(Txtn_Year))
        Call Display_Records
     End If
End Sub

Private Sub Display_Month()
  Dim i As Integer
  
      Va_Pay.Col = 1
      Va_Pay.Row = 1
         Va_Pay.Text = "January" & Space(100) & "01"
      Va_Pay.Row = 2
         Va_Pay.Text = "February" & Space(100) & "02"
      Va_Pay.Row = 3
         Va_Pay.Text = "March" & Space(100) & "03"
      Va_Pay.Row = 4
         Va_Pay.Text = "April" & Space(100) & "04"
      Va_Pay.Row = 5
         Va_Pay.Text = "May" & Space(100) & "05"
      Va_Pay.Row = 6
         Va_Pay.Text = "June" & Space(100) & "06"
      Va_Pay.Row = 7
         Va_Pay.Text = "July" & Space(100) & "07"
      Va_Pay.Row = 8
         Va_Pay.Text = "August" & Space(100) & "08"
      Va_Pay.Row = 9
         Va_Pay.Text = "September" & Space(100) & "09"
      Va_Pay.Row = 10
         Va_Pay.Text = "October" & Space(100) & "10"
      Va_Pay.Row = 11
         Va_Pay.Text = "November" & Space(100) & "11"
      Va_Pay.Row = 12
         Va_Pay.Text = "December" & Space(100) & "12"
      Va_Pay.Row = 13
         Va_Pay.Text = "E.O.Y Bonus" & Space(100) & "13"
         
      Va_Pay.Row = 13
      Va_Pay.Col = 2
         Va_Pay.Text = "01/01/" & Format(Is_Null(Txtn_Year, True), "0000")
      Va_Pay.Col = 3
         Va_Pay.Text = "31/12/" & Format(Is_Null(Txtn_Year, True), "0000")
End Sub

Private Sub Spread_Lock()
  Dim i As Integer
      For i = 1 To Va_Pay.MaxCols
          Va_Pay.Row = -1
          Va_Pay.Col = i
          If i = 1 Or i = 4 Or i = 5 Or i = 7 Then
             Va_Pay.Lock = True
          Else
             Va_Pay.Lock = False
          End If
      Next i
End Sub

Private Sub Va_Pay_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
  Dim i As Integer
  Dim vFromDate As Date, vToDate As Date
  
    If Col = 3 Then
       For i = 1 To 11
           Va_Pay.Row = i
           Va_Pay.Col = 3
              If IsDate(Va_Pay.Text) Then
                 vToDate = CDate(Va_Pay.Text)
              End If
           Va_Pay.Row = i + 1
           Va_Pay.Col = 2
              Va_Pay.Text = Is_DateSpread(DateAdd("d", 1, vToDate), False)
       Next i
    End If
  
    For i = 1 To 13
        Va_Pay.Row = i
        Va_Pay.Col = 2
           If IsDate(Va_Pay.Text) Then
              vFromDate = CDate(Va_Pay.Text)
           End If
        Va_Pay.Col = 3
           If IsDate(Va_Pay.Text) Then
              vToDate = CDate(Va_Pay.Text)
           End If
           
        If IsDate(vFromDate) And IsDate(vToDate) Then
           Va_Pay.Col = 4
              If DateDiff("d", vFromDate, vToDate) > 0 Then
                 Va_Pay.Text = DateDiff("d", vFromDate, vToDate) + 1
              Else
                 Va_Pay.Text = ""
              End If
              
           Va_Pay.Col = 5
              If DateDiff("ww", vFromDate, vToDate) > 0 Then
                  Va_Pay.Text = DateDiff("ww", vFromDate, vToDate)
              Else
                 Va_Pay.Text = ""
              End If
        End If
    Next i
End Sub
