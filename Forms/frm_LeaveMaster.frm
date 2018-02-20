VERSION 5.00
Begin VB.Form frm_LeaveMaster 
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   3750
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   2565
   ScaleWidth      =   3750
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Cmb_OthLeaveFlag 
      Height          =   330
      Left            =   2910
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   3330
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   900
      Left            =   120
      TabIndex        =   18
      Top             =   -60
      Width           =   6900
      Begin VB.CommandButton Btn_View 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   1005
         Picture         =   "frm_LeaveMaster.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Save 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   30
         Picture         =   "frm_LeaveMaster.frx":36B0
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Exit 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   4470
         Picture         =   "frm_LeaveMaster.frx":6D39
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Clear 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   1725
         Picture         =   "frm_LeaveMaster.frx":A399
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Delete 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   2715
         Picture         =   "frm_LeaveMaster.frx":DA09
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Print 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   3435
         Picture         =   "frm_LeaveMaster.frx":110B3
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
      Height          =   3660
      Left            =   120
      TabIndex        =   12
      Top             =   1170
      Width           =   6915
      Begin VB.TextBox Txtn_YearAllot 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2775
         MaxLength       =   7
         TabIndex        =   8
         Top             =   1170
         Width           =   1230
      End
      Begin VB.TextBox Txtn_MaxAccdays 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2775
         MaxLength       =   7
         TabIndex        =   9
         Top             =   1470
         Width           =   1230
      End
      Begin VB.TextBox Txtc_Remarks 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2775
         MaxLength       =   50
         TabIndex        =   11
         Top             =   3105
         Width           =   3780
      End
      Begin VB.TextBox Txtc_LeaveName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2775
         MaxLength       =   50
         TabIndex        =   7
         Top             =   660
         Width           =   3780
      End
      Begin VB.TextBox Txtc_Leave 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   2775
         MaxLength       =   7
         TabIndex        =   6
         Top             =   360
         Width           =   1230
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "This leave will be consider as Other Leave for Calculating Vacation Leave "
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
         Height          =   630
         Left            =   315
         TabIndex        =   21
         Top             =   2040
         Width           =   2355
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Max Accumulate Days"
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
         Left            =   840
         TabIndex        =   20
         Top             =   1515
         Width           =   1785
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Yearly Allotment"
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
         Left            =   1260
         TabIndex        =   19
         Top             =   1215
         Width           =   1365
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Leave Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   210
         Left            =   1650
         TabIndex        =   17
         Top             =   405
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Leave Name"
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
         Left            =   1620
         TabIndex        =   16
         Top             =   705
         Width           =   1005
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         X1              =   75
         X2              =   7050
         Y1              =   2850
         Y2              =   2850
      End
      Begin VB.Label Label1 
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
         Left            =   1875
         TabIndex        =   15
         Top             =   3120
         Width           =   750
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
      Left            =   5550
      TabIndex        =   14
      Top             =   915
      Width           =   360
   End
   Begin VB.Label lbl_scr_name 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Leave Details"
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
      Left            =   285
      TabIndex        =   13
      Top             =   930
      Width           =   1095
   End
   Begin VB.Shape shp_scr_name 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   120
      Top             =   870
      Width           =   6885
   End
End
Attribute VB_Name = "frm_LeaveMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As New ADODB.Recordset

Private Sub Form_Load()
    lbl_date.Caption = Format(Date, "dd-mmm-yyyy")
    Enable_Controls Me, True
    Call TGControlProperty(Me)
    Call Load_Combo
    Call Create_Default_LeaveTypes
    
    Cmb_OthLeaveFlag.ListIndex = 0
End Sub

Private Sub Form_Activate()
    If Txtc_Leave.Enabled = True Then
       Txtc_Leave.SetFocus
    End If
End Sub

Private Sub Btn_Exit_Click()
    Unload Me
End Sub

Private Sub Btn_Clear_Click()
    Clear_Controls Me
    Txtc_Leave.Enabled = True
    Txtc_Leave.BackColor = vbWhite
End Sub

Private Sub Btn_Delete_Click()
On Error GoTo ErrDel
    
    If Trim(Txtc_Leave) = "" Then
       Exit Sub
    End If
    
    If g_Admin Then
        If (MsgBox("Are you sure you want to delete ?", vbYesNo, "Confirmation") = vbYes) Then
            CON.BeginTrans
            CON.Execute "update pr_leave_mst set " & GetDelFlag & " where c_leave = '" & Is_Null(Txtc_Leave, False) & "'"
            CON.CommitTrans
            Call Btn_Clear_Click
        End If
    Else
        MsgBox "No access to delete. Please check with Admin", vbInformation, "Information"
    End If
  
  Exit Sub

ErrDel:
    MsgBox "Error while Deleting - " + Err.Description, vbCritical, "Critical"
    CON.RollbackTrans
End Sub

Private Sub Btn_View_Click()
 Dim Search As New Search.MyClass, SerVar
 
    Search.Query = "Select c_leave Leave, c_leavename LeaveName, c_othleaveflag OtherLeaveFlag, c_remarks Remarks " & _
                   "from pr_leave_mst where c_rec_sta='A'"
    Search.CheckFields = "Leave"
    Search.ReturnField = "Leave"
    SerVar = Search.Search(, , CON)
    If Len(Search.col1) <> 0 Then
       Txtc_Leave = Search.col1
       Call Display_Records
       Txtc_Leave.Enabled = False
       Txtc_Leave.BackColor = &HE0E0E0
    End If
End Sub

Private Sub Btn_Print_Click()
On Error GoTo Err_Print
  Dim SelFor As String, RepTitle As String
  
  
   RepTitle = "Leave Master"
   SelFor = ""
   Call Print_Rpt(SelFor, "Pr_Leave_List.rpt")
  
   If Trim(RepTitle) <> "" Then
      Mdi_Ta_HrPay.CRY1.Formulas(1) = "ReportHead='" & UCase(Trim(RepTitle)) & "'"
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
  
    Save_Pr_Leave_Mst
  
  CON.CommitTrans
  
  g_SaveFlagNull = False
  Screen.MousePointer = vbDefault

  MsgBox "Record Saved Successfully", vbInformation, "Information"
  Btn_Clear_Click

 Exit Sub

ErrSave:
   CON.RollbackTrans
   g_SaveFlagNull = False
   Screen.MousePointer = vbDefault
   MsgBox "Error while Saving - " + Err.Description, vbCritical, "Critical"
End Sub

Private Function ChkSave() As Boolean
Dim rsChk As New ADODB.Recordset
  
  If Trim(Txtc_Leave) = "" Then
     MsgBox "Leave Code shoule not be empty", vbInformation, "Information"
     Txtc_Leave.SetFocus
     Exit Function
  ElseIf Trim(Txtc_LeaveName) = "" Then
     MsgBox "Leave Name shoule not be empty", vbInformation, "Information"
     Txtc_LeaveName.SetFocus
     Exit Function
  End If
  ChkSave = True
End Function

Private Sub Save_Pr_Leave_Mst()
    Set rs = Nothing
    g_Sql = "Select * from pr_leave_mst where c_leave = '" & Is_Null(Txtc_Leave, False) & "'"
    rs.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
    
    If rs.RecordCount = 0 Then
       rs.AddNew
       rs("c_usr_id").Value = g_UserName
       rs("d_created").Value = GetDateTime
    Else
       rs("d_modified").Value = GetDateTime
       rs("c_musr_id").Value = g_UserName
       
    End If
    rs("c_leave").Value = UCase(Is_Null(Txtc_Leave, False))
    rs("c_leavename").Value = Proper(Is_Null(Txtc_LeaveName, False))
    rs("n_yearallot").Value = Proper(Is_Null(Txtn_YearAllot, True))
    rs("n_maxaccdays").Value = Proper(Is_Null(Txtn_MaxAccdays, True))
    rs("c_othleaveflag").Value = Is_Null(Right(Trim(Cmb_OthLeaveFlag), 1), False)
    rs("c_remarks").Value = Is_Null(Txtc_Remarks, False)
    
    rs("c_rec_sta").Value = "A"
    rs.Update
    
End Sub

Private Sub Display_Records()
On Error GoTo ErrDisp
  Dim rsDisp As New ADODB.Recordset
  Dim i As Integer
    
    Set rsDisp = Nothing
    g_Sql = "select * from pr_leave_mst where c_rec_sta='A' and c_leave = '" & Is_Null(Txtc_Leave, False) & "'"
    rsDisp.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If rsDisp.RecordCount > 0 Then
       Txtc_Leave = Is_Null(rsDisp("c_leave").Value, False)
       Txtc_LeaveName = Is_Null(rsDisp("c_leavename").Value, False)
       Txtn_YearAllot = Is_Null(rsDisp("n_yearallot").Value, False)
       Txtn_MaxAccdays = Is_Null(rsDisp("n_maxaccdays").Value, False)
       Txtc_Remarks = Is_Null(rsDisp("c_remarks").Value, False)
       
       For i = 0 To Cmb_OthLeaveFlag.ListCount - 1
           If Right(Trim(Cmb_OthLeaveFlag.List(i)), 1) = Is_Null(rsDisp("c_othleaveflag").Value, False) Then
              Cmb_OthLeaveFlag.ListIndex = i
              Exit For
           End If
       Next i
    End If
  
  Exit Sub
  
ErrDisp:
   MsgBox "Display Error - " + Err.Description, vbCritical, "Critical"
End Sub

Private Sub Txtc_Leave_Validate(Cancel As Boolean)
  Dim rsChk As New ADODB.Recordset
    
     If Trim(Txtc_Leave) <> "" Then
        Txtc_Leave = UCase(Txtc_Leave)
        Set rsChk = Nothing
        g_Sql = "select c_leave from pr_leave_mst where c_leave = '" & Trim(Txtc_Leave) & "'"
        rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
        If rsChk.RecordCount > 0 Then
           MsgBox "The leave is already available and should be unique", vbInformation, "Information"
           Txtc_Leave.SetFocus
           Cancel = True
        End If
    End If
End Sub

Private Sub Txtn_MaxAccDays_KeyPress(KeyAscii As Integer)
    Call OnlyNumeric(Txtn_MaxAccdays, KeyAscii, 3, 0)
End Sub

Private Sub Txtn_YearAllot_KeyPress(KeyAscii As Integer)
    Call OnlyNumeric(Txtn_YearAllot, KeyAscii, 3, 0)
End Sub

Private Sub Load_Combo()
    
    Cmb_OthLeaveFlag.Clear
    Cmb_OthLeaveFlag.AddItem "No" & Space(100) & "N"
    Cmb_OthLeaveFlag.AddItem "Yes" & Space(100) & "Y"

End Sub

