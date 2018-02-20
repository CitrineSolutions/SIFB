VERSION 5.00
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Begin VB.Form frm_Users 
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
   Begin VB.Frame Frame2 
      Height          =   900
      Left            =   120
      TabIndex        =   24
      Top             =   -90
      Width           =   12030
      Begin VB.CommandButton Btn_View 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   1005
         Picture         =   "frm_Users.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Save 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   30
         Picture         =   "frm_Users.frx":36B0
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Exit 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   4470
         Picture         =   "frm_Users.frx":6D39
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Clear 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   1725
         Picture         =   "frm_Users.frx":A399
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Delete 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   2715
         Picture         =   "frm_Users.frx":DA09
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Print 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   3435
         Picture         =   "frm_Users.frx":110B3
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
      Height          =   4965
      Left            =   120
      TabIndex        =   16
      Top             =   1185
      Width           =   12030
      Begin VB.CheckBox Chk_Delete 
         Height          =   255
         Left            =   10245
         TabIndex        =   13
         Top             =   135
         Width           =   315
      End
      Begin VB.CheckBox Chk_Add 
         Height          =   255
         Left            =   9450
         TabIndex        =   12
         Top             =   135
         Width           =   315
      End
      Begin VB.CheckBox Chk_Admin 
         Height          =   255
         Left            =   8655
         TabIndex        =   11
         Top             =   135
         Width           =   315
      End
      Begin VB.CheckBox Chk_View 
         Height          =   255
         Left            =   11070
         TabIndex        =   14
         Top             =   135
         Width           =   315
      End
      Begin VB.TextBox Txtc_UserName 
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
         Left            =   1380
         MaxLength       =   10
         TabIndex        =   6
         Top             =   345
         Width           =   2625
      End
      Begin VB.TextBox Txtc_Name 
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
         Left            =   1380
         MaxLength       =   50
         TabIndex        =   7
         Top             =   645
         Width           =   2625
      End
      Begin VB.ComboBox Cmb_Branch 
         Height          =   330
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   945
         Width           =   2625
      End
      Begin VB.OptionButton Opt_Admin 
         Caption         =   "Admin"
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
         Left            =   2940
         TabIndex        =   10
         Top             =   1620
         Width           =   1095
      End
      Begin VB.OptionButton Opt_User 
         Caption         =   "User"
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
         Left            =   1350
         TabIndex        =   9
         Top             =   1620
         Width           =   855
      End
      Begin FPUSpreadADO.fpSpread Va_Details 
         Height          =   4395
         Left            =   4860
         TabIndex        =   15
         Top             =   390
         Width           =   6960
         _Version        =   458752
         _ExtentX        =   12277
         _ExtentY        =   7752
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
         MaxCols         =   5
         MaxRows         =   10
         SpreadDesigner  =   "frm_Users.frx":146A1
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Note : "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   210
         Left            =   195
         TabIndex        =   23
         Top             =   2430
         Width           =   510
      End
      Begin VB.Label Label1 
         Caption         =   "The new user password will be same as Login and the system will force the user to change the password at the first time of login."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   810
         Left            =   195
         TabIndex        =   22
         Top             =   2715
         Width           =   4005
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Login"
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
         Left            =   750
         TabIndex        =   21
         Top             =   390
         Width           =   465
      End
      Begin VB.Label Label11 
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
         Height          =   210
         Left            =   750
         TabIndex        =   20
         Top             =   690
         Width           =   465
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
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
         Height          =   210
         Left            =   645
         TabIndex        =   19
         Top             =   1005
         Width           =   570
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
      Left            =   10815
      TabIndex        =   18
      Top             =   915
      Width           =   360
   End
   Begin VB.Label lbl_scr_name 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Login User"
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
      TabIndex        =   17
      Top             =   900
      Width           =   900
   End
   Begin VB.Shape shp_scr_name 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   120
      Top             =   855
      Width           =   12015
   End
End
Attribute VB_Name = "frm_Users"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As New ADODB.Recordset
Private p_mode As String

Private Sub Chk_Admin_Click()
  Dim i As Integer
    If Chk_Admin.Value = 1 Then
        For i = 1 To Va_Details.DataRowCnt
            Va_Details.Row = i
            Va_Details.Col = 2
               Va_Details.Value = True
        Next i
    End If
End Sub

Private Sub Chk_Add_Click()
  Dim i As Integer
    If Chk_Add.Value = 1 Then
        For i = 1 To Va_Details.DataRowCnt
            Va_Details.Row = i
            Va_Details.Col = 3
               Va_Details.Value = True
        Next i
    End If
End Sub

Private Sub Chk_Delete_Click()
  Dim i As Integer
    If Chk_Delete.Value = 1 Then
        For i = 1 To Va_Details.DataRowCnt
            Va_Details.Row = i
            Va_Details.Col = 4
               Va_Details.Value = True
        Next i
    End If
End Sub

Private Sub Chk_View_Click()
  Dim i As Integer
    If Chk_View.Value = 1 Then
        For i = 1 To Va_Details.DataRowCnt
            Va_Details.Row = i
            Va_Details.Col = 5
               Va_Details.Value = True
        Next i
    End If
End Sub

Private Sub Form_Activate()
    Txtc_UserName.SetFocus
    Opt_User.Value = True
End Sub

Private Sub Form_Load()
    lbl_date.Caption = Format(Date, "dd-mmm-yyyy")
    Opt_Admin.Visible = False

    Call Spread_Lock
    Call TGControlProperty(Me)
    Clear_Spread Va_Details
    Call Load_Combo
    Call Create_Default_Screens
    Call Display_Screens
    Call Spread_Row_Height(Va_Details)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub Btn_Exit_Click()
    Unload Me
End Sub

Private Sub Btn_Clear_Click()
  Dim i As Integer
    
    Txtc_UserName.Enabled = True
    Txtc_UserName.Text = ""
    Txtc_Name.Text = ""
    Cmb_Branch.ListIndex = 0
    
    Chk_Admin.Value = 0
    Chk_Add.Value = 0
    Chk_Delete.Value = 0
    Chk_View.Value = 0
   
    For i = 1 To Va_Details.DataRowCnt
        Va_Details.Row = i
        Va_Details.Col = 2
           Va_Details.Value = 0
        Va_Details.Col = 3
           Va_Details.Value = 0
        Va_Details.Col = 4
           Va_Details.Value = 0
        Va_Details.Col = 5
           Va_Details.Value = 0
    Next i
End Sub

Private Sub Btn_Delete_Click()
On Error GoTo ErrDel

    If Trim(Txtc_UserName) = "" Then
       Exit Sub
    End If
    
    If Not g_Admin Then
       MsgBox "No access to Delete. Please contact Admin", vbInformation, "Information"
       Exit Sub
    End If

    If (MsgBox("Are you sure you want to delete ?", vbExclamation + vbYesNo, lbl_scr_name.Caption) = vbYes) Then
        CON.BeginTrans
        
        CON.Execute "delete from pr_user_mst where c_user_id = '" & Trim(Txtc_UserName) & "'"
        CON.Execute "delete from pr_user_dtl where c_user_id = '" & Trim(Txtc_UserName) & "'"
        
        CON.CommitTrans
        Call Btn_Clear_Click
    End If
  
  Exit Sub
ErrDel:
    MsgBox "Error while Deleting - " + Err.Description, vbCritical, "Critical"
    CON.RollbackTrans
End Sub


Private Sub Btn_View_Click()
 Dim Search As New Search.MyClass, SerVar
    Search.Query = "Select c_user_id Login, c_user_name Name, c_branch Branch from pr_user_mst "
    Search.CheckFields = "Login"
    Search.ReturnField = "Login"
    SerVar = Search.Search(, , CON)
    If Len(Search.col1) <> 0 Then
       Txtc_UserName = Search.col1
       Call Display_Records
    End If
End Sub

Private Sub Display_Records()
On Error GoTo Err_Display
Dim DyDisp As New ADODB.Recordset
Dim i As Integer, tmpStr As String
    
   Set DyDisp = Nothing
   g_Sql = "select * from pr_user_mst where c_user_id = '" & Trim(Txtc_UserName) & "'"
   DyDisp.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
   If DyDisp.RecordCount > 0 Then
      Txtc_UserName = Is_Null(DyDisp("c_user_id").Value, False)
      Txtc_UserName.Enabled = False
      
      Txtc_Name = Is_Null(DyDisp("c_user_name").Value, False)
      Call DisplayComboBranch(Me, Is_Null(DyDisp("c_branch").Value, False))
      If DyDisp("c_user_level").Value = "2" Then
         Opt_Admin.Visible = True
      Else
         Opt_User.Value = True
      End If
   End If
   
   ' User Form Details
   Set DyDisp = Nothing
   g_Sql = "select a.c_screen_id, a.c_screen_name, b.c_adminright, b.c_addright, b.c_delright, b.c_viewright " & _
           "from pr_screen_mst a left outer join pr_user_dtl b on a.c_screen_id = b.c_screen_id and b.c_user_id = '" & Trim(Txtc_UserName) & "' " & _
           "order by a.c_screen_name "
   
   DyDisp.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
   If DyDisp.RecordCount > 0 Then
      Va_Details.MaxRows = DyDisp.RecordCount
      DyDisp.MoveFirst
      For i = 1 To DyDisp.RecordCount
          Va_Details.Row = i
          Va_Details.Col = 1
             Va_Details.Text = Proper(Is_Null(DyDisp("c_screen_name").Value, False)) & Space(100) & Is_Null(DyDisp("c_screen_id").Value, False)
          Va_Details.Col = 2
             Va_Details.Value = Is_Null(DyDisp("c_adminright").Value, True)
          Va_Details.Col = 3
             Va_Details.Value = Is_Null(DyDisp("c_addright").Value, True)
          Va_Details.Col = 4
             Va_Details.Value = Is_Null(DyDisp("c_delright").Value, True)
          Va_Details.Col = 5
             Va_Details.Value = Is_Null(DyDisp("c_viewright").Value, True)
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

Private Sub Display_Screens()
 Dim DyDisp As New ADODB.Recordset
 Dim i As Integer
 
   Set DyDisp = Nothing
   g_Sql = "select c_screen_id, c_screen_name from pr_screen_mst order by c_screen_name "
   
   DyDisp.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
   If DyDisp.RecordCount > 0 Then
      Va_Details.MaxRows = DyDisp.RecordCount
      DyDisp.MoveFirst
      For i = 1 To DyDisp.RecordCount
          Va_Details.Row = i
          Va_Details.Col = 1
             Va_Details.Text = Proper(Is_Null(DyDisp("c_screen_name").Value, False)) & Space(100) & Is_Null(DyDisp("c_screen_id").Value, False)
          Va_Details.Col = 2
             Va_Details.Value = False
          Va_Details.Col = 3
             Va_Details.Value = False
          Va_Details.Col = 4
             Va_Details.Value = False
          Va_Details.Col = 5
             Va_Details.Value = False
          DyDisp.MoveNext
      Next i
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
        
        Save_User_Mst
        Save_User_Dtl
     
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
 Dim rsChk As New ADODB.Recordset
  
  If Trim(Txtc_UserName) = "" Then
     MsgBox "Login should not be empty", vbInformation, "Information"
     Txtc_UserName.SetFocus
     Exit Function
  ElseIf Trim(Txtc_Name) = "" Then
     MsgBox "Name should not be empty", vbInformation, "Information"
     Txtc_Name.SetFocus
     Exit Function
  End If
  
  ChkSave = True
End Function

Private Sub Save_User_Mst()
    
    Set rs = Nothing
    g_Sql = "Select * from pr_user_mst where c_user_id = '" & Trim(Txtc_UserName) & "'"
    rs.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
    If rs.RecordCount = 0 Then
        rs.AddNew
        rs("c_cusr_id").Value = g_UserName
        rs("d_created").Value = GetDateTime
    Else
       rs("c_musr_id").Value = g_UserName
       rs("d_modified").Value = GetDateTime
    End If
    
    rs("c_user_id") = Proper(Is_Null(UCase(Txtc_UserName), False))
    rs("c_user_name") = Proper(Is_Null(UCase(Txtc_Name), False))
    rs("c_branch").Value = Is_Null(Proper(Right(Trim(Cmb_Branch), 7)), False)
    
    If Is_Null(rs("n_pwd").Value, True) = 0 Then
       rs("c_user_pwd").Value = Is_Null(UCase(Txtc_UserName), False)
       rs("d_pwd_on").Value = GetDateTime
    End If
    
    If Opt_Admin.Value Then
      rs("c_user_level").Value = "2"
    Else
      rs("c_user_level").Value = "0"
    End If
    rs("c_rec_sta").Value = "A"

    rs.Update
End Sub


Private Sub Save_User_Dtl()
 Dim i As Long
    
       g_Sql = "delete from pr_user_dtl where c_user_id = '" & Trim(Txtc_UserName) & "'"
       CON.Execute (g_Sql)
    
       Set rs = Nothing
       g_Sql = "Select * from pr_user_dtl where 1=2"
       rs.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
       
       For i = 1 To Va_Details.DataRowCnt
           Va_Details.Row = i
           Va_Details.Col = 1
           If Trim(Va_Details.Text) <> "" Then
              rs.AddNew
              rs("c_user_id").Value = Is_Null(Txtc_UserName, False)
              Va_Details.Col = 1
                 rs("c_screen_id").Value = Is_Null(Right(Trim(Va_Details.Text), 25), False)
              Va_Details.Col = 2
                 If Va_Details.Value = True Then
                    rs("c_adminright").Value = "1"
                 Else
                    rs("c_adminright").Value = "0"
                 End If
              Va_Details.Col = 3
                 If Va_Details.Value = True Then
                    rs("c_addright").Value = "1"
                 Else
                    rs("c_addright").Value = "0"
                 End If
              Va_Details.Col = 4
                 If Va_Details.Value = True Then
                    rs("c_delright").Value = "1"
                 Else
                    rs("c_delright").Value = "0"
                 End If
              Va_Details.Col = 5
                 If Va_Details.Value = True Then
                    rs("c_viewright").Value = "1"
                 Else
                    rs("c_viewright").Value = "0"
                 End If
              rs.Update
           End If
       Next i
End Sub

Private Sub Txtc_UserName_Validate(Cancel As Boolean)
    Call Display_Records
End Sub

Private Sub Va_Details_DblClick(ByVal Col As Long, ByVal Row As Long)
   Call SpreadColSort(Va_Details, Col, Row)
End Sub

Private Sub Va_Details_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
   If Col = 2 Then
      Va_Details.Row = Row
      Va_Details.Col = 2
         If Va_Details.Value = True Then
            Va_Details.Col = 3
               Va_Details.Value = True
            Va_Details.Col = 4
               Va_Details.Value = True
            Va_Details.Col = 5
               Va_Details.Value = True
         End If
   End If
End Sub

Private Sub Load_Combo()
    Call LoadComboBranch(Me)
End Sub


Private Sub Spread_Lock()
    If g_Admin Then
       Opt_Admin.Visible = True
       Va_Details.Row = -1
       Va_Details.Col = -1
          Va_Details.Lock = False
    Else
       Opt_Admin.Visible = False
       Va_Details.Row = -1
       Va_Details.Col = -1
          Va_Details.Lock = True
    End If
End Sub





