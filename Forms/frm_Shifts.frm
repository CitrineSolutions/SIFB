VERSION 5.00
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Begin VB.Form frm_Shifts 
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   11115
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7110
   ScaleWidth      =   11115
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   900
      Left            =   120
      TabIndex        =   7
      Top             =   -60
      Width           =   14895
      Begin VB.CommandButton Btn_View 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   1005
         Picture         =   "frm_Shifts.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Save 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   30
         Picture         =   "frm_Shifts.frx":36B0
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Exit 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   4470
         Picture         =   "frm_Shifts.frx":6D39
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Clear 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   1725
         Picture         =   "frm_Shifts.frx":A399
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Delete 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   2715
         Picture         =   "frm_Shifts.frx":DA09
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Print 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   3435
         Picture         =   "frm_Shifts.frx":110B3
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   700
      End
   End
   Begin FPUSpreadADO.fpSpread Va_Details 
      Height          =   6240
      Left            =   120
      TabIndex        =   6
      Top             =   1215
      Width           =   14895
      _Version        =   458752
      _ExtentX        =   26273
      _ExtentY        =   11007
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
      MaxCols         =   19
      MaxRows         =   15
      ProcessTab      =   -1  'True
      SpreadDesigner  =   "frm_Shifts.frx":146A1
      VisibleCols     =   1
      VisibleRows     =   1
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
      Height          =   195
      Left            =   13425
      TabIndex        =   9
      Top             =   915
      Width           =   360
   End
   Begin VB.Label lbl_scr_name 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Clock Card Shift Details"
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
      TabIndex        =   8
      Top             =   915
      Width           =   1920
   End
   Begin VB.Shape shp_scr_name 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   300
      Left            =   120
      Top             =   870
      Width           =   14895
   End
End
Attribute VB_Name = "frm_Shifts"
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
    Call Spread_Row_Height(Va_Details, 15, 30)
    Call Clear_Spread(Va_Details)
    Call Create_Default_Shift
    Call Display_Records
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
    
    If Va_Details.DataRowCnt = 0 Then
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
  
  
   RepTitle = "Shift Detail"
   SelFor = ""
   Call Print_Rpt(SelFor, "Pr_Shift_List.rpt")
  
   If Trim(RepTitle) <> "" Then
      Mdi_Ta_HrPay.CRY1.Formulas(1) = "ReportHead='" & UCase(Trim(RepTitle)) & "'"
   End If

   Mdi_Ta_HrPay.CRY1.Action = 1
  
  Exit Sub

Err_Print:
    MsgBox "Error while Generating - " + Err.Description, vbInformation, "Information"
End Sub

Private Sub Btn_View_Click()
   Call Display_Records
End Sub

Private Sub Btn_Save_Click()
On Error GoTo ErrSave
 
     If ChkSave = False Then
        Exit Sub
     End If
     
     Screen.MousePointer = vbHourglass
     g_SaveFlagNull = True
     
     CON.BeginTrans
        
        Save_PR_Clock_Shift
        Update_PR_ShiftStructure_Dtl
     
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
  If IsColDupValues(Va_Details, 1) > 0 Then
     MsgBox "Duplicate code found. ", vbInformation, "Information"
     Va_Details.SetFocus
     Exit Function
  End If
  ChkSave = True
End Function

Private Sub Save_PR_Clock_Shift()
Dim i As Long

    g_Sql = "truncate table pr_clock_shift "
    CON.Execute (g_Sql)
    
    Set rs = Nothing
    g_Sql = "Select * from pr_clock_shift"
    rs.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
    
    For i = 1 To Va_Details.DataRowCnt
        Va_Details.Row = i
        Va_Details.Col = 1
        If Trim(Va_Details.Text) <> "" Then
           rs.AddNew
           Va_Details.Col = 1
              rs("c_shiftcode").Value = Is_Null(Va_Details.Text, False)
           Va_Details.Col = 2
              rs("n_starthrs").Value = Is_Null(Va_Details.Text, True)
              rs("starthrs").Value = TimeToMins(Is_Null(Va_Details.Text, True))
           Va_Details.Col = 3
              rs("n_endhrs").Value = Is_Null(Va_Details.Text, True)
              rs("endhrs").Value = TimeToMins(Is_Null(Va_Details.Text, True))
           
           Va_Details.Col = 4  'Mins
              rs("n_breakhrs").Value = Is_Null(Va_Details.Text, True)
              rs("breakhrs").Value = Is_Null(Va_Details.Text, True)
           Va_Details.Col = 5
              rs("n_shifthrs").Value = Is_Null(Va_Details.Text, True)
              rs("shifthrs").Value = TimeToMins(Is_Null(Va_Details.Text, True))
           Va_Details.Col = 6 'Mins
              rs("n_latehrs").Value = Is_Null(Va_Details.Text, True)
              rs("latehrs").Value = Is_Null(Va_Details.Text, True)
           
           Va_Details.Col = 7
              rs("n_break1").Value = Is_Null(Va_Details.Text, True)
              rs("break1").Value = TimeToMins(Is_Null(Va_Details.Text, True))
           Va_Details.Col = 8
              rs("n_mins1").Value = Is_Null(Va_Details.Text, True)
              rs("mins1").Value = Is_Null(Va_Details.Text, True)
           
           Va_Details.Col = 9
              rs("n_break2").Value = Is_Null(Va_Details.Text, True)
              rs("break2").Value = TimeToMins(Is_Null(Va_Details.Text, True))
           Va_Details.Col = 10
              rs("n_mins2").Value = Is_Null(Va_Details.Text, True)
              rs("mins2").Value = Is_Null(Va_Details.Text, True)
           
           Va_Details.Col = 11
              rs("n_break3").Value = Is_Null(Va_Details.Text, True)
              rs("break3").Value = TimeToMins(Is_Null(Va_Details.Text, True))
           Va_Details.Col = 12
              rs("n_mins3").Value = Is_Null(Va_Details.Text, True)
              rs("mins3").Value = Is_Null(Va_Details.Text, True)
           
           Va_Details.Col = 13
              rs("n_clmin").Value = Is_Null(Va_Details.Text, True)
              rs("clmin").Value = Is_Null(Va_Details.Text, True)
           Va_Details.Col = 14
              rs("n_cutoffhrs").Value = Is_Null(Va_Details.Text, True)
              rs("cutoffhrs").Value = TimeToMins(Is_Null(Va_Details.Text, True))
              rs("n_maxhrs").Value = Is_Null(Va_Details.Text, True)
              rs("maxhrs").Value = TimeToMins(Is_Null(Va_Details.Text, True))
           Va_Details.Col = 15
              rs("n_permhrs").Value = Is_Null(Va_Details.Text, True)
              rs("permhrs").Value = Is_Null(Va_Details.Text, True)
           
           Va_Details.Col = 16
              rs("n_maafter").Value = Is_Null(Va_Details.Text, True)
              rs("maafter").Value = TimeToMins(Is_Null(Va_Details.Text, True))
           Va_Details.Col = 17
              rs("n_naafter").Value = Is_Null(Va_Details.Text, True)
              rs("naafter").Value = TimeToMins(Is_Null(Va_Details.Text, True))
           
           Va_Details.Col = 18
              rs("c_desp").Value = Is_Null(Va_Details.Text, False)
           
           rs.Update
        End If
    Next i
End Sub

Private Sub Update_PR_ShiftStructure_Dtl()
Dim i As Long
Dim vNewCode As String, vOldCode As String
    
    For i = 1 To Va_Details.DataRowCnt
        Va_Details.Row = i
        Va_Details.Col = 1
        If Trim(Va_Details.Text) <> "" Then
           Va_Details.Col = 1
              vNewCode = Trim(Va_Details.Text)
           Va_Details.Col = 19
              vOldCode = Trim(Va_Details.Text)
        End If
    
        If vNewCode <> "" And vOldCode <> "" Then
        
           g_Sql = "Update pr_shiftstructure_dtl Set c_shiftcode = '" & vNewCode & "' Where c_shiftcode is not null and c_shiftcode = '" & vOldCode & "'"
           CON.Execute g_Sql
        
           g_Sql = "Update pr_shiftstructure_dtl Set c_wocode = '" & vNewCode & "' Where c_wocode is not null and c_wocode = '" & vOldCode & "'"
           CON.Execute g_Sql
        
           g_Sql = "Update pr_shiftstructure_dtl Set c_altcode = '" & vNewCode & "' Where c_altcode is not null and c_altcode = '" & vOldCode & "'"
           CON.Execute g_Sql
        
        End If
    Next i
End Sub

Private Sub Display_Records()
On Error GoTo Err_Display
  Dim DyDisp As New ADODB.Recordset
  Dim i, j As Long
  
  Set DyDisp = Nothing
  g_Sql = "select * from pr_clock_shift "
  DyDisp.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
  If DyDisp.RecordCount > 0 Then
     DyDisp.MoveFirst
     For i = 1 To DyDisp.RecordCount
        Va_Details.Row = i
        Va_Details.Col = 1
           Va_Details.Text = Is_Null(DyDisp("c_shiftcode").Value, False)
        Va_Details.Col = 2
           Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_starthrs").Value, True), True)
        Va_Details.Col = 3
           Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_endhrs").Value, True), True)
        
        Va_Details.Col = 4
           Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_breakhrs").Value, True), True)
        Va_Details.Col = 5
           Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_shifthrs").Value, True), True)
        Va_Details.Col = 6
           Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_latehrs").Value, True), True)
        
        Va_Details.Col = 7
           Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_break1").Value, True), True)
        Va_Details.Col = 8
           Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_mins1").Value, True), True)
        
        Va_Details.Col = 9
           Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_break2").Value, True), True)
        Va_Details.Col = 10
           Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_mins2").Value, True), True)
        
        Va_Details.Col = 11
           Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_break3").Value, True), True)
        Va_Details.Col = 12
           Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_mins3").Value, True), True)
        
        Va_Details.Col = 13
           Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_clmin").Value, True), True)
        Va_Details.Col = 14
           Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_cutoffhrs").Value, True), True)
        Va_Details.Col = 15
           Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_permhrs").Value, True), True)
        
        Va_Details.Col = 16
           Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_maafter").Value, True), True)
        Va_Details.Col = 17
           Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_naafter").Value, True), True)
        
        Va_Details.Col = 18
           Va_Details.Text = Is_Null(DyDisp("c_desp").Value, False)
        Va_Details.Col = 19
           Va_Details.Text = Is_Null(DyDisp("c_shiftcode").Value, False)
        
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

Private Sub Spread_Lock()
  Dim i As Integer
     
    For i = 1 To Va_Details.MaxCols
        Va_Details.Row = -1
        Va_Details.Col = i
        If i = 19 Then
           Va_Details.Lock = True
        Else
           Va_Details.Lock = False
        End If
    Next i
End Sub

Private Sub Va_Details_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyDelete Then
     Call SpreadCellDataClear(Va_Details, Va_Details.ActiveRow, Va_Details.ActiveCol)
  End If
End Sub
