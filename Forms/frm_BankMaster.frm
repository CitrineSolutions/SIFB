VERSION 5.00
Begin VB.Form frm_BankMaster 
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   13695
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6825
   ScaleWidth      =   13695
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   900
      Left            =   120
      TabIndex        =   23
      Top             =   -60
      Width           =   10395
      Begin VB.CommandButton Btn_Print 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   3435
         Picture         =   "frm_BankMaster.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Delete 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   2715
         Picture         =   "frm_BankMaster.frx":35EE
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Clear 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   1725
         Picture         =   "frm_BankMaster.frx":6C98
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Exit 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   4470
         Picture         =   "frm_BankMaster.frx":A308
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Save 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   30
         Picture         =   "frm_BankMaster.frx":D968
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_View 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   1005
         Picture         =   "frm_BankMaster.frx":10FF1
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   700
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2190
      Left            =   120
      TabIndex        =   15
      Top             =   1200
      Width           =   10410
      Begin VB.TextBox Txtc_Remarks 
         Height          =   750
         Left            =   6375
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   1170
         Width           =   3720
      End
      Begin VB.TextBox Txtc_BankAdd 
         Height          =   750
         Left            =   1185
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   1170
         Width           =   3720
      End
      Begin VB.TextBox Txtc_Branch 
         Height          =   300
         Left            =   6375
         MaxLength       =   50
         TabIndex        =   11
         Top             =   855
         Width           =   3720
      End
      Begin VB.TextBox Txtc_ShortName 
         Height          =   300
         Left            =   1185
         MaxLength       =   10
         TabIndex        =   8
         Top             =   855
         Width           =   3720
      End
      Begin VB.TextBox Txtc_BankCode 
         Height          =   300
         Left            =   6375
         MaxLength       =   10
         TabIndex        =   10
         Top             =   555
         Width           =   3720
      End
      Begin VB.TextBox Txtc_BankName 
         Height          =   300
         Left            =   1185
         MaxLength       =   50
         TabIndex        =   7
         Top             =   540
         Width           =   3720
      End
      Begin VB.TextBox Txtc_Code 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   1185
         TabIndex        =   6
         Top             =   240
         Width           =   885
      End
      Begin VB.Label Label5 
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
         Left            =   5550
         TabIndex        =   22
         Top             =   1440
         Width           =   750
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   375
         TabIndex        =   21
         Top             =   1425
         Width           =   735
      End
      Begin VB.Label Label3 
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
         Left            =   5730
         TabIndex        =   20
         Top             =   900
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Short Name"
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
         Left            =   150
         TabIndex        =   19
         Top             =   900
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Bank Code"
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
         Left            =   5415
         TabIndex        =   18
         Top             =   585
         Width           =   885
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Bank Name"
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
         Left            =   195
         TabIndex        =   17
         Top             =   585
         Width           =   915
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
         Left            =   675
         TabIndex        =   16
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
      Left            =   9225
      TabIndex        =   14
      Top             =   915
      Width           =   360
   End
   Begin VB.Label lbl_scr_name 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Bank Master"
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
      Left            =   300
      TabIndex        =   13
      Top             =   930
      Width           =   1035
   End
   Begin VB.Shape shp_scr_name 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   120
      Top             =   885
      Width           =   10395
   End
End
Attribute VB_Name = "frm_BankMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs As New ADODB.Recordset

Private Sub Form_Load()
    lbl_date.Caption = Format(Date, "dd-mmm-yyyy")
    Call TGControlProperty(Me)
    Call Create_Default_Bank
End Sub

Private Sub Form_Activate()
    Txtc_BankName.SetFocus
End Sub

Private Sub Btn_Exit_Click()
    Unload Me
End Sub

Private Sub Btn_Clear_Click()
    Clear_Controls Me
    Txtc_Code.Enabled = False
End Sub

Private Sub Btn_Delete_Click()
On Error GoTo ErrDel
    If Trim(Txtc_Code) = "" Then
       Exit Sub
    End If
    
    If g_Admin Then
       If (MsgBox("Are you sure you want to delete ?", vbYesNo, "Confirmation") = vbYes) Then
          CON.BeginTrans
          CON.Execute "update pr_bankmast set " & GetDelFlag & " where c_code = '" & Trim(Txtc_Code) & "'"
          CON.CommitTrans
       End If
       Call Btn_Clear_Click
    Else
        MsgBox "No access to delete. Please check with Admin", vbInformation, "Information"
    End If
  Exit Sub

ErrDel:
    CON.RollbackTrans
    MsgBox "Error while Deleting - " + Err.Description, vbCritical, "Critical"
End Sub

Private Sub Btn_Print_Click()
On Error GoTo Err_Print
  Dim SelFor As String, RepTitle As String
  
  
   RepTitle = "Bank List"
   SelFor = ""
   Call Print_Rpt(SelFor, "Pr_Bank_List.rpt")
  
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

    Search.Query = "select c_bankname BankName, c_shortname ShortName, c_bankcode BankCode, c_branch Branch, c_code Code " & _
                   "from pr_bankmast where c_rec_sta='A' "
    Search.CheckFields = "ShiftName, Code"
    Search.ReturnField = "ShiftName, Code"
    SerVar = Search.Search(, , CON)
    If Len(Search.col2) <> 0 Then
        Txtc_Code = Search.col2
        Call Display_Records
        Txtc_Code.Enabled = False
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
        
        Save_PR_BankMast
        
     CON.CommitTrans
     
     g_SaveFlagNull = False
     Screen.MousePointer = vbDefault
    
     MsgBox "Record Saved Successfully", vbInformation, "Information"
     Txtc_Code.Enabled = False
  
  Exit Sub
     
ErrSave:
     CON.RollbackTrans
     g_SaveFlagNull = False
     Screen.MousePointer = vbDefault
     MsgBox "Error while Saving - " + Err.Description, vbCritical, "Critical"
     
End Sub

Private Function ChkSave() As Boolean
  Dim i As Integer
  
  If Trim(Txtc_BankName) = "" Then
     MsgBox "Bank Name should not be empty", vbInformation, "Information"
     Txtc_BankName.SetFocus
     Exit Function
  ElseIf Trim(Txtc_ShortName) = "" Then
     MsgBox "Short Name should not be empty", vbInformation, "Information"
     Txtc_ShortName.SetFocus
     Exit Function
  ElseIf Trim(Txtc_BankCode) = "" Then
     MsgBox "Bank Code should not be empty", vbInformation, "Information"
     Txtc_BankCode.SetFocus
     Exit Function
  End If
   
  ChkSave = True
End Function

Private Sub Save_PR_BankMast()
    Set rs = Nothing
    g_Sql = "Select * from pr_bankmast where c_code = '" & Trim(Txtc_Code) & "' and c_rec_sta='A'"
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
    rs("c_bankname").Value = Is_Null(Txtc_BankName, False)
    rs("c_shortname").Value = Is_Null(Txtc_ShortName, False)
    rs("c_bankcode").Value = Is_Null(Txtc_BankCode, False)
    rs("c_branch").Value = Is_Null(Txtc_Branch, False)
    rs("c_bankadd").Value = Is_Null(Txtc_BankAdd, False)
    rs("c_remarks").Value = Is_Null(Txtc_Remarks, False)
    
    rs("c_rec_sta").Value = "A"
    rs.Update
End Sub

Private Sub Start_Generate_New()
  Dim MaxNo As ADODB.Recordset
  g_Sql = "Select max(right(c_code,2)) from pr_bankmast "
  Set MaxNo = CON.Execute(g_Sql)
  Txtc_Code = "B" & Format(Is_Null(MaxNo(0).Value, True) + 1, "00")
End Sub

Private Sub Display_Records()
On Error GoTo Err_Display
  Dim DyDisp As New ADODB.Recordset
  Dim i, j As Long
  Dim vType As String
  
  Set DyDisp = Nothing
  g_Sql = "select * from pr_bankmast where c_rec_sta = 'A' and c_code = '" & Trim(Txtc_Code) & "'"
  DyDisp.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
  
  Txtc_Code = Is_Null(DyDisp("c_code").Value, False)
  Txtc_BankName = Is_Null(DyDisp("c_bankname").Value, False)
  Txtc_ShortName = Is_Null(DyDisp("c_shortname").Value, False)
  Txtc_BankCode = Is_Null(DyDisp("c_bankcode").Value, False)
  Txtc_Branch = Is_Null(DyDisp("c_branch").Value, False)
  Txtc_BankAdd = Is_Null(DyDisp("c_bankadd").Value, False)
  Txtc_Remarks = Is_Null(DyDisp("c_remarks").Value, False)

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


