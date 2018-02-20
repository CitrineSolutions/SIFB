VERSION 5.00
Begin VB.Form frm_ChangePwd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Password"
   ClientHeight    =   4770
   ClientLeft      =   2145
   ClientTop       =   2820
   ClientWidth     =   6570
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
   ScaleHeight     =   4770
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Txtc_NewPwd1 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2055
      Width           =   2655
   End
   Begin VB.TextBox Txtc_NewPwd2 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2505
      Width           =   2655
   End
   Begin VB.TextBox Txtc_UserName 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2880
      TabIndex        =   0
      Top             =   1200
      Width           =   2655
   End
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
      Left            =   3540
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3480
      Width           =   1155
   End
   Begin VB.CommandButton Btn_Ok 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Ok "
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
      Left            =   1920
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3480
      Width           =   1155
   End
   Begin VB.TextBox Txtc_OldPwd 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1620
      Width           =   2655
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Re-type New Password"
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
      Left            =   810
      TabIndex        =   10
      Top             =   2595
      Width           =   1920
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "New Password"
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
      Left            =   1485
      TabIndex        =   9
      Top             =   2130
      Width           =   1245
   End
   Begin VB.Label Label3 
      Caption         =   "Change Password is required either this is your first login or your password is expired"
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   915
      TabIndex        =   8
      Top             =   270
      Width           =   5010
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
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
      Left            =   2340
      TabIndex        =   7
      Top             =   1260
      Width           =   390
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Old Password"
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
      Left            =   1575
      TabIndex        =   6
      Top             =   1680
      Width           =   1155
   End
End
Attribute VB_Name = "frm_ChangePwd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mnuOption As String

Private Sub Form_Load()
    ' mnuOption   UC - User Changing
    
    Txtc_UserName = Proper(g_UserName)
    Txtc_UserName.Enabled = False
End Sub

Private Sub Form_Activate()
    Txtc_OldPwd.SetFocus
End Sub

Private Sub Btn_Ok_Click()
On Error GoTo ErrSave
  Dim rsChk As New ADODB.Recordset
  
  
     If ChkSave = False Then
        Exit Sub
     End If
  
     CON.BeginTrans
     
        g_Sql = "update pr_user_mst set c_user_pwd = '" & Trim(Txtc_NewPwd1) & "', n_pwd = n_pwd+1, " & _
                "d_pwd_on = '" & GetDateTime() & "' " & _
                "where c_user_id = '" & Trim(Txtc_UserName) & "'"
        CON.Execute g_Sql
        
     CON.CommitTrans

     If Not Check_User() Then
        Exit Sub
     End If
       
     Unload Me
  
  Exit Sub
  
ErrSave:
     CON.RollbackTrans
     g_SaveFlagNull = False
     Screen.MousePointer = vbDefault
     MsgBox "Error while Saving - " + Err.Description, vbCritical, "Critical"
End Sub

Private Function ChkSave() As Boolean
  Dim i As Integer
  
  If Trim(Txtc_UserName) = "" Then
     MsgBox "User should not be empty", vbInformation, "Information"
     Txtc_OldPwd.SetFocus
     Exit Function
  ElseIf Trim(Txtc_OldPwd) = "" Then
     MsgBox "Old password should not be empty", vbInformation, "Information"
     Txtc_OldPwd.SetFocus
     Exit Function
  ElseIf Trim(Txtc_NewPwd1) = "" Then
     MsgBox "New Password should not be empty", vbInformation, "Information"
     Txtc_NewPwd1.SetFocus
     Exit Function
  ElseIf Trim(Txtc_NewPwd2) = "" Then
     MsgBox "Re-type New Password should not be empty", vbInformation, "Information"
     Txtc_NewPwd2.SetFocus
     Exit Function
  ElseIf Trim(Txtc_NewPwd1) <> Trim(Txtc_NewPwd2) Then
     MsgBox "New password and re-type new password is not matched", vbInformation, "Information"
     Txtc_NewPwd2.SetFocus
     Exit Function
  ElseIf Trim(Txtc_OldPwd) = Trim(Txtc_NewPwd1) Then
     MsgBox "New password should not be same as old password.", vbInformation, "Information"
     Txtc_NewPwd2.SetFocus
     Exit Function
  End If
   
  ChkSave = True
End Function


Private Function Check_User() As Boolean
  Dim rsChk As New ADODB.Recordset

     Set rsChk = Nothing
     g_Sql = "select *  from pr_user_mst where c_user_id = '" & Trim(Txtc_UserName) & "' and " & _
             "c_user_pwd = '" & Trim(Txtc_NewPwd1.Text) & "'"
     rsChk.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
     If rsChk.RecordCount > 0 Then
        g_UserName = Proper(Trim(Txtc_UserName))
        
        If Is_Null(rsChk("c_user_level").Value, True) = 2 Then
           g_Admin = True
        Else
           g_Admin = False
        End If
        
        If UCase(g_UserName) = "NITHIN" Then
           g_Author = True
        End If

     Else
        MsgBox "Not a Valid User", vbInformation, "Information"
        Txtc_NewPwd1.SetFocus
        Check_User = False
        Exit Function
     End If
     
     Check_User = True
End Function

Private Sub Btn_Cancel_Click()
    If mnuOption = "UC" Then
       Unload Me
    Else
       End
    End If
End Sub

Private Sub Txtc_NewPwd1_Validate(Cancel As Boolean)
    If Trim(Txtc_NewPwd1) <> "" Then
       If Len(Trim(Txtc_NewPwd1)) < 7 Then
          MsgBox "Password should be minimum 7 charactors length", vbInformation, "Information"
          Txtc_NewPwd1.SetFocus
          Cancel = True
       End If
    End If
End Sub

Private Sub Txtc_OldPwd_Validate(Cancel As Boolean)
  Dim rsChk As New ADODB.Recordset

    If Trim(Txtc_OldPwd) <> "" Then
       Set rsChk = Nothing
       g_Sql = "select *  from pr_user_mst where c_user_id = '" & UCase(Trim(Txtc_UserName.Text)) & "' and " & _
               "c_user_pwd = '" & UCase(Trim(Txtc_OldPwd.Text)) & "'"
       rsChk.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
       If rsChk.RecordCount = 0 Then
          MsgBox "Incorrect old password.", vbInformation, "Information"
          Txtc_OldPwd.SetFocus
          Cancel = True
       End If
    End If
End Sub
