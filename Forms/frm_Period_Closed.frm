VERSION 5.00
Begin VB.Form frm_Period_Closed 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   360
      TabIndex        =   5
      Top             =   600
      Width           =   5295
      Begin VB.CommandButton Btn_OK 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Ok"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1305
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1200
         Width           =   1080
      End
      Begin VB.CommandButton Btn_Exit 
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
         Height          =   375
         Left            =   2925
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1200
         Width           =   1080
      End
      Begin VB.TextBox Txtc_Month 
         Height          =   300
         Left            =   2325
         MaxLength       =   2
         TabIndex        =   0
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox Txtc_Year 
         Height          =   300
         Left            =   2865
         MaxLength       =   4
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         X1              =   0
         X2              =   6000
         Y1              =   915
         Y2              =   915
      End
      Begin VB.Label Lbl_Year 
         AutoSize        =   -1  'True
         Caption         =   "Period"
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
         Left            =   1605
         TabIndex        =   6
         Top             =   405
         Width           =   540
      End
   End
   Begin VB.Label lbl_scr_name 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Pay Period Closed"
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
      Left            =   2340
      TabIndex        =   4
      Top             =   285
      Width           =   1515
   End
   Begin VB.Shape shp_scr_name 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   360
      Top             =   240
      Width           =   5295
   End
End
Attribute VB_Name = "frm_Period_Closed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mnuOption As String
Private vPayPeriod As Long

Private Sub Form_Load()
   If mnuOption = "C" Then
      lbl_scr_name.Caption = "Pay Period Closing"
   ElseIf mnuOption = "O" Then
      lbl_scr_name.Caption = "Pay Period Re-Open"
   End If
   
   Call TGControlProperty(Me)
End Sub

Private Sub Btn_Ok_Click()
 On Error GoTo Err_Proc
   
      If Trim(Txtc_Month) = "" Then
         MsgBox "Month should not be empty", vbInformation, "Information"
         Txtc_Month.SetFocus
         Exit Sub
      End If
      If Trim(Txtc_Year) = "" Then
         MsgBox "Year should not be empty", vbInformation, "Information"
         Txtc_Year.SetFocus
         Exit Sub
      End If
     
      vPayPeriod = Format(Trim(Txtc_Year), "0000") & Format(Trim(Txtc_Month), "00")
 
      Screen.MousePointer = vbHourglass
        
      If mnuOption = "C" Or mnuOption = "O" Then
         Save_Pr_Period_Closed
      End If
       
      Screen.MousePointer = vbDefault
  
  Exit Sub

Err_Proc:
      Screen.MousePointer = vbDefault
      CON.RollbackTrans
      MsgBox Err.Description, vbCritical, "Error"
      Btn_Ok.Enabled = True
      Btn_Exit.Enabled = True
End Sub

Private Sub Btn_Exit_Click()
  Unload Me
End Sub

Private Sub Txtc_Month_KeyPress(KeyAscii As Integer)
    Call OnlyNumeric(Txtc_Month, KeyAscii, 2)
End Sub

Private Sub Txtc_Month_Validate(Cancel As Boolean)
    If Trim(Txtc_Month) <> "" Then
       If Val(Txtc_Month) <= 0 Or Val(Txtc_Month) > 13 Then
          MsgBox "Not a valid month", vbInformation, "Information"
          Txtc_Month.SetFocus
          Cancel = True
       End If
    End If
End Sub

Private Sub txtc_year_KeyPress(KeyAscii As Integer)
    Call OnlyNumeric(Txtc_Year, KeyAscii, 4)
End Sub

Private Sub txtc_year_Validate(Cancel As Boolean)
   If Trim(Txtc_Year) <> "" Then
      If Len(Txtc_Year) <> 4 Then
         MsgBox "Not a valid year", vbInformation, "Information"
         Txtc_Year.SetFocus
         Cancel = True
      End If
   End If
End Sub

Private Sub Save_Pr_Period_Closed()
  Dim rsChk As New ADODB.Recordset
     
     Set rsChk = Nothing
     g_Sql = "select * from pr_payperiod_dtl where n_period = " & vPayPeriod & " and c_type = 'W'"
     rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
     If rsChk.RecordCount = 0 Then
        MsgBox "This period is not found.", vbInformation, "Information"
        Exit Sub
     End If
     
     If mnuOption = "C" Then
        If rsChk("c_period_closed").Value = "Y" Then
           MsgBox "This period is already closed", vbInformation, "Information"
           Exit Sub
        End If
     Else
        If rsChk("c_period_closed").Value = "N" Then
           MsgBox "This period is already open", vbInformation, "Information"
           Exit Sub
        End If
     End If
     
     CON.BeginTrans
         If mnuOption = "C" Then
            g_Sql = "update pr_payperiod_dtl set c_period_closed = 'Y' where n_period = " & vPayPeriod & " and c_type = 'W'"
            CON.Execute g_Sql
         Else
            g_Sql = "update pr_payperiod_dtl set c_period_closed = 'N' where n_period = " & vPayPeriod & " and c_type = 'W'"
            CON.Execute g_Sql
         End If
        
         Set rsChk = Nothing
         g_Sql = "select * from pr_payperiod_dtl_log where 1=2"
         rsChk.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
        
         rsChk.AddNew
         rsChk("n_period").Value = Is_Null(vPayPeriod, True)
         rsChk("c_type").Value = "W"
         
         If mnuOption = "C" Then
            rsChk("c_remarks").Value = "Closed"
         Else
            rsChk("c_remarks").Value = "Re-opened"
         End If
            
         rsChk("c_usr_id").Value = Is_Null(g_UserName, False)
         rsChk("d_created").Value = GetDateTime
         rsChk.Update
     CON.CommitTrans
    
     If mnuOption = "C" Then
        MsgBox "Period Closed Successfully", vbInformation, "Information"
     Else
        MsgBox "Period Re-opened Successfully", vbInformation, "Information"
     End If
End Sub




