VERSION 5.00
Begin VB.Form frm_Login 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " User Login"
   ClientHeight    =   3885
   ClientLeft      =   2145
   ClientTop       =   2820
   ClientWidth     =   6105
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Btn_Cancel 
      BackColor       =   &H00C0E0FF&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3330
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2385
      Width           =   1155
   End
   Begin VB.CommandButton Btn_Ok 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Login"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1710
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2385
      Width           =   1155
   End
   Begin VB.ComboBox Cmb_User 
      Height          =   345
      ItemData        =   "frm_Login.frx":0000
      Left            =   2265
      List            =   "frm_Login.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   870
      Width           =   2460
   End
   Begin VB.TextBox Txt_Pwd 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2265
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1455
      Width           =   2460
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "User Name"
      Height          =   225
      Left            =   1110
      TabIndex        =   5
      Top             =   915
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   225
      Left            =   1185
      TabIndex        =   4
      Top             =   1515
      Width           =   870
   End
End
Attribute VB_Name = "frm_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call Create_Default_System
    Call Load_Combo
    Call Display_Combo
    
    Btn_Ok.Enabled = False
    
    g_Author = False
    g_Admin = False
    
    g_FrmSupUser = False
    g_FrmAddRight = False
    g_FrmModRight = False
    g_FrmDelRight = False
    g_FrmViewRight = False
    
    If Trim(Left(Trim(g_Server), 9)) = "NITHIN-PC" Then
      Txt_Pwd.Text = "Admin123"
    End If
End Sub

Private Sub Form_Activate()
  Txt_Pwd.SetFocus
End Sub

Private Sub Btn_Ok_Click()
On Error GoTo Err_Dis
  Dim rsChk As New ADODB.Recordset
  Dim bDate As String, eDate As String, vODBCDatabase As String
  
    Screen.MousePointer = vbHourglass
    
    g_ReportPath = App.Path & "\Reports\"
  
    Set rsChk = Nothing
    g_Sql = "select * from sysusers where uid=0"
    rsChk.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
    If rsChk.RecordCount > 0 Then
       g_CurrentDate = Format(Date, "dd/mm/yyyy hh:mm:ss")
       Call StartDB
       
       If Not Chk_DateFormat() Then
          Screen.MousePointer = vbDefault
          Exit Sub
       End If
       
       If Not Check_User() Then
          Screen.MousePointer = vbDefault
          Exit Sub
       End If
       
       vODBCDatabase = GetKeyValueInOBDC("Database")
       
       If Trim(g_Database) <> Trim(vODBCDatabase) Then
          MsgBox "The ODBC Database is not " & g_Database & ",  Please change it.", vbInformation, "Information"
       End If
       
       Screen.MousePointer = vbDefault
       Unload Me
    Else
       Screen.MousePointer = vbDefault
       MsgBox "No Database found. Please contact Admin", vbInformation, "Information"
    End If
  Exit Sub
 
Err_Dis:
   CON.RollbackTrans
   Screen.MousePointer = vbDefault
   MsgBox "Critical Error - " + Err.Description
End Sub

Private Function Check_User() As Boolean
  Dim rsChk As New ADODB.Recordset

     Set rsChk = Nothing
     g_Sql = "select *  from pr_user_mst where c_user_id = '" & UCase(Trim(Cmb_User.Text)) & "' and " & _
             "c_user_pwd = '" & UCase(Trim(Txt_Pwd.Text)) & "'"
     rsChk.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
     If rsChk.RecordCount > 0 Then
        g_UserName = Proper(Trim(Cmb_User.Text))
        
        If Is_Null(rsChk("c_user_level").Value, True) = 2 Then
           g_Admin = True
        Else
           g_Admin = False
        End If
        
        If UCase(g_UserName) = "NITHIN" Then
           g_Author = True
        End If
        
        If Is_Null(rsChk("n_pwd").Value, True) = 0 Then
           Screen.MousePointer = vbDefault
           frm_ChangePwd.Show 1
        End If

        
        If DateDiff("d", Is_Date(rsChk("d_pwd_on").Value, "D"), g_CurrentDate) > 100 Then
           Screen.MousePointer = vbDefault
           frm_ChangePwd.Show 1
        End If

     Else
        MsgBox "Not a Valid User", vbInformation, "Information"
        Txt_Pwd.SetFocus
        Check_User = False
        Exit Function
     End If
     
     Check_User = True
End Function

Private Sub Btn_Cancel_Click()
    End
End Sub

Private Sub Txt_Pwd_Change()
    If Trim(Txt_Pwd.Text) = "" Then
        Btn_Ok.Enabled = False
    Else
        Btn_Ok.Enabled = True
    End If
End Sub

Private Sub Txt_Pwd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call Btn_Ok_Click
    End If
End Sub

Private Function Chk_DateFormat() As Boolean
   Dim rsChk As New ADODB.Recordset
   Dim ServerDate As String, SystemDate As String
   Dim tmpStr As String
   
      Set rsChk = Nothing
      g_Sql = "select getdate() "
      rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
      If rsChk.RecordCount > 0 Then
         ServerDate = Format(rsChk(0).Value, "dd/mm/yyyy")
      End If
    
      SystemDate = Left(Trim(Str(Now)), 10)
      
      If ServerDate <> SystemDate Then
         tmpStr = "Your system date should be DD/MM/YYYY format. Please change setting in Regional Settings in Control Panel"
         MsgBox tmpStr, vbInformation, "Information"
         Chk_DateFormat = False
         Exit Function
      End If
 
      Chk_DateFormat = True
End Function

Private Sub Load_Combo()
  Dim rsCombo As New ADODB.Recordset
  Dim i As Integer
  
    Set rsCombo = Nothing
    g_Sql = "Select * from pr_user_mst where c_rec_sta = 'A' order by c_user_id"
    rsCombo.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    Cmb_User.Clear
    Cmb_User.AddItem ""
    For i = 0 To rsCombo.RecordCount - 1
        Cmb_User.AddItem Proper(rsCombo("c_user_id").Value)
        rsCombo.MoveNext
    Next i
End Sub

Private Sub Display_Combo()
  Dim i As Integer
  
    If g_UserName = "" Then
       Cmb_User.ListIndex = 1
    Else
       For i = 0 To Cmb_User.ListCount - 1
          If Trim(Cmb_User.List(i)) = Trim(g_UserName) Then
             Cmb_User.ListIndex = i
             Exit For
          End If
       Next i
    End If
End Sub

Private Sub Save_Pr_Daily_Proc_Log()
  Dim rs As New ADODB.Recordset
  Dim vToDate As Date
  
    Set rs = Nothing
    g_Sql = "Select * from pr_daily_proc_log where d_date = '" & Is_Date(g_CurrentDate, "S") & "'"
    rs.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
    
    If rs.RecordCount = 0 Then
       rs.AddNew
       rs("d_date").Value = Is_Date(g_CurrentDate, "S")
       rs("c_usr_id").Value = g_UserName
       rs("d_created").Value = GetDateTime
       rs.Update
    End If
End Sub


Private Sub Create_Default_System()
    Call Create_Default_User
    Call Create_Default_Company
    Call Create_Default_Bank
    Call Create_Default_Shift
    Call Create_Default_ShiftStructure
    Call Create_Default_PayPeriods(Year(GetDateTime))
    Call Create_Default_PayStructure("COM0001")
    Call Create_Default_LeaveTypes
    Call Create_Default_Holidays(Year(GetDateTime))
    Call Create_Default_EDFTypes
End Sub
