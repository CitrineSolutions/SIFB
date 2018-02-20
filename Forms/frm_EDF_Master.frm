VERSION 5.00
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Begin VB.Form frm_EDF_Master 
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5460
   ScaleWidth      =   7665
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   900
      Left            =   120
      TabIndex        =   9
      Top             =   -45
      Width           =   5775
      Begin VB.CommandButton Btn_Print 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   3435
         Picture         =   "frm_EDF_Master.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Delete 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   2715
         Picture         =   "frm_EDF_Master.frx":35EE
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Clear 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   1725
         Picture         =   "frm_EDF_Master.frx":6C98
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Exit 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   4470
         Picture         =   "frm_EDF_Master.frx":A308
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Save 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   30
         Picture         =   "frm_EDF_Master.frx":D968
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_View 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   1005
         Picture         =   "frm_EDF_Master.frx":10FF1
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   700
      End
   End
   Begin FPUSpreadADO.fpSpread Va_Details 
      Height          =   4335
      Left            =   120
      TabIndex        =   6
      Top             =   1260
      Width           =   5775
      _Version        =   458752
      _ExtentX        =   10186
      _ExtentY        =   7646
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
      MaxCols         =   3
      MaxRows         =   10
      ProcessTab      =   -1  'True
      SpreadDesigner  =   "frm_EDF_Master.frx":146A1
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
      Left            =   4500
      TabIndex        =   8
      Top             =   945
      Width           =   390
   End
   Begin VB.Label lbl_scr_name 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "EDF Structure"
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
      TabIndex        =   7
      Top             =   960
      Width           =   1140
   End
   Begin VB.Shape shp_scr_name 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   285
      Left            =   120
      Top             =   930
      Width           =   5775
   End
End
Attribute VB_Name = "frm_EDF_Master"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As New ADODB.Recordset
Private p_mode As String

Private Sub Form_Load()
    lbl_date.Caption = Format(Date, "dd-mmm-yyyy")
    Call TGControlProperty(Me)
    Call Spread_Row_Height(Va_Details)
    Clear_Spread Va_Details
    Call Create_Default_EDFTypes
    Call Display_Records
End Sub

Private Sub Form_Activate()
    Va_Details.SetFocus
End Sub

Private Sub Btn_Clear_Click()
    Clear_Spread Va_Details
End Sub

Private Sub Btn_Delete_Click()
On Error GoTo ErrDel
    MsgBox "No access to delete", vbInformation, "Information"
   
   Exit Sub
ErrDel:
    MsgBox "Error while Deleting - " + Err.Description, vbCritical, "Critical"
    CON.RollbackTrans
End Sub

Private Sub Btn_Exit_Click()
    Unload Me
End Sub

Private Sub Btn_Print_Click()
On Error GoTo Err_Print
  Dim SelFor As String, RepTitle As String
   
   RepTitle = "EDF Category"
   SelFor = ""
   Call Print_Rpt(SelFor, "Pr_EDF_List.rpt")
  
   If Trim(RepTitle) <> "" Then
      Mdi_Ta_HrPay.CRY1.Formulas(1) = "ReportHead='" & Trim(RepTitle) & "'"
   End If

   Mdi_Ta_HrPay.CRY1.Action = 1
  
  Exit Sub

Err_Print:
    MsgBox "Error while Generating - " + Err.Description, vbInformation, "Information"
End Sub

Private Sub Btn_Save_Click()
On Error GoTo ErrSave
  
     If ChkSave = False Then
        Exit Sub
     End If
     
     Screen.MousePointer = vbHourglass
     g_SaveFlagNull = True
     
     CON.BeginTrans
        Save_Pr_Edfmast
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

Private Sub Btn_View_Click()
   Call Display_Records
End Sub

Private Sub tlb1_exitclick()
    Unload Me
End Sub

Private Function ChkSave() As Boolean
Dim i As Integer
  If IsColDupValues(Va_Details, 1) > 0 Then
     MsgBox "Duplicate code found. ", vbInformation, "Information"
     Va_Details.SetFocus
     Exit Function
  ElseIf Va_Details.DataRowCnt = 0 Then
     MsgBox "No records to save. ", vbInformation, "Information"
     Exit Function
  End If
  ChkSave = True
End Function

Private Sub Save_Pr_Edfmast()
Dim i As Long
    Set rs = Nothing
    g_Sql = "truncate table pr_edfmast "
    CON.Execute (g_Sql)
     
    g_Sql = "Select * from pr_edfmast where 1=2"
    rs.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
        
    For i = 1 To Va_Details.DataRowCnt
        Va_Details.Row = i
        Va_Details.Col = 1
        If Trim(Va_Details.Text) <> "" Then
           rs.AddNew
           Va_Details.Col = 1
              rs("c_category").Value = Is_Null(Va_Details.Text, False)
           Va_Details.Col = 2
              rs("c_desp").Value = Is_Null(Va_Details.Text, False)
           Va_Details.Col = 3
              rs("n_edfamt").Value = Is_Null_D(Va_Details.Text, True)
           rs.Update
        End If
    Next i
End Sub

Private Sub Display_Records()
On Error GoTo Err_Display
  Dim DyDisp As New ADODB.Recordset
  Dim i As Long
  
  Set DyDisp = Nothing
  g_Sql = "select * from pr_edfmast order by c_category "
  DyDisp.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
  If DyDisp.RecordCount > 0 Then
     DyDisp.MoveFirst
     For i = 1 To DyDisp.RecordCount
        Va_Details.Row = i
        Va_Details.Col = 1
           Va_Details.Text = Is_Null(DyDisp("c_category").Value, False)
        Va_Details.Col = 2
           Va_Details.Text = Is_Null(DyDisp("c_desp").Value, False)
        Va_Details.Col = 3
           Va_Details.Text = Is_Null_D(DyDisp("n_edfamt").Value, True)
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

Private Sub Va_Details_KeyDown(KeyCode As Integer, Shift As Integer)
  
  If ((Shift And 1) = 1) And KeyCode = vbKeyInsert Then
     Call SpreadInsertRow(Va_Details, Va_Details.ActiveRow)
     Call Spread_Row_Height(Va_Details)
  
  ElseIf ((Shift And 1) = 1) And KeyCode = vbKeyDelete Then
     Call SpreadDeleteRow(Va_Details, Va_Details.ActiveRow)
     Call Spread_Row_Height(Va_Details)
  
  ElseIf KeyCode = vbKeyDelete Then
     Call SpreadCellDataClear(Va_Details, Va_Details.ActiveRow, Va_Details.ActiveCol)
  
  End If
End Sub

