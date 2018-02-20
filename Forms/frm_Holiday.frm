VERSION 5.00
Object = "{C3A136DA-B937-492B-968D-A437638F7AAB}#1.0#0"; "CS_DateControl.ocx"
Begin VB.Form frm_Holiday 
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
      TabIndex        =   21
      Top             =   -45
      Width           =   8910
      Begin VB.CommandButton Btn_Print 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   3435
         Picture         =   "frm_Holiday.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Delete 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   2715
         Picture         =   "frm_Holiday.frx":35EE
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Clear 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   1725
         Picture         =   "frm_Holiday.frx":6C98
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Exit 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   4470
         Picture         =   "frm_Holiday.frx":A308
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Save 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   30
         Picture         =   "frm_Holiday.frx":D968
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_View 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   1005
         Picture         =   "frm_Holiday.frx":10FF1
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   700
      End
      Begin VB.Frame Frm_EmpStatus 
         Caption         =   "Year Option"
         ForeColor       =   &H00C00000&
         Height          =   720
         Left            =   5685
         TabIndex        =   22
         Top             =   120
         Width           =   2925
         Begin VB.OptionButton Opt_All 
            Caption         =   "All"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   2115
            TabIndex        =   7
            Top             =   345
            Width           =   705
         End
         Begin VB.OptionButton Opt_Current 
            Caption         =   "Current Year"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   240
            TabIndex        =   6
            Top             =   330
            Value           =   -1  'True
            Width           =   1560
         End
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
      Height          =   1650
      Left            =   120
      TabIndex        =   12
      Top             =   1245
      Width           =   5400
      Begin VB.TextBox Txtn_Uno 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   480
         MaxLength       =   20
         TabIndex        =   8
         Top             =   225
         Width           =   825
      End
      Begin VB.ComboBox Cmb_No 
         Height          =   330
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1650
         Width           =   2505
      End
      Begin VB.ComboBox Cmb_Type 
         Height          =   330
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1035
         Width           =   2500
      End
      Begin VB.TextBox Txtc_Desp 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   10
         Top             =   740
         Width           =   2500
      End
      Begin CS_DateControl.DateControl Dtp_Date 
         Height          =   345
         Left            =   2040
         TabIndex        =   9
         Top             =   225
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   609
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   0
         X2              =   5345
         Y1              =   1515
         Y2              =   1515
      End
      Begin VB.Label Lbl_Uno 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Uno"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   210
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Lbl_Flag 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Flag"
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
         Left            =   1590
         TabIndex        =   19
         Top             =   1710
         Width           =   330
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         X1              =   10
         X2              =   5360
         Y1              =   650
         Y2              =   650
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Holiday Description"
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
         Left            =   345
         TabIndex        =   18
         Top             =   800
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Date"
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
         Left            =   1560
         TabIndex        =   17
         Top             =   285
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Type"
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
         Left            =   1515
         TabIndex        =   16
         Top             =   1095
         Width           =   405
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
      Left            =   4320
      TabIndex        =   15
      Top             =   990
      Width           =   360
   End
   Begin VB.Label lbl_scr_name 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Holiday Details"
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
      Left            =   225
      TabIndex        =   14
      Top             =   990
      Width           =   1185
   End
   Begin VB.Shape shp_scr_name 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   120
      Top             =   930
      Width           =   5400
   End
End
Attribute VB_Name = "frm_Holiday"
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
    Call Create_Default_Holidays(Year(g_CurrentDate))
    
    Lbl_Uno.Visible = False
    Txtn_Uno.Visible = False
    Lbl_Flag.Visible = False
    Cmb_No.Visible = False
    Line2.Visible = False
End Sub

Private Sub Form_Activate()
    Dtp_Date.SetFocus
    Cmb_Type.ListIndex = 0
    Cmb_No.ListIndex = 0
End Sub

Private Sub Btn_Exit_Click()
    Unload Me
End Sub

Private Sub Btn_Clear_Click()
    Clear_Controls Me
    Cmb_Type.ListIndex = 0
    Cmb_No.ListIndex = 0
End Sub

Private Sub Btn_Delete_Click()
On Error GoTo ErrDel

   If IsDate(Dtp_Date.Text) Then
      If CDate(Dtp_Date.Text) < g_CurrentDate Then
         MsgBox "No access to delete", vbInformation, "Information"
         Exit Sub
      End If
   Else
      Exit Sub
   End If

   If g_Admin Then
      If (MsgBox("Are you sure you want to delete ?", vbExclamation + vbYesNo, lbl_scr_name.Caption) = vbYes) Then
         CON.BeginTrans
         CON.Execute "update pr_holiday_mst set " & GetDelFlag & " where n_uno = " & Is_Null(Txtn_Uno, True)
         CON.CommitTrans
         Call Btn_Clear_Click
      End If
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
  
   If IsDate(Dtp_Date.Text) Then
      SelFor = "YEAR({PR_HOLIDAY_MST.D_PHDATE})=" & Year(Dtp_Date.Text)
   Else
      SelFor = "YEAR({PR_HOLIDAY_MST.D_PHDATE})=" & Year(g_CurrentDate)
   End If
   
   RepTitle = "Holiday Master"
   Call Print_Rpt(SelFor, "Pr_Holiday_List.rpt")
  
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
      
    If Opt_Current.Value = True Then
       Search.Query = "Select c_desp Description, Convert(char(10), d_phdate,126) HolidayDate, n_uno No " & _
                      "from pr_holiday_mst where c_rec_sta='A' and " & _
                      "Year(d_phdate) = " & Year(g_CurrentDate)
    Else
       Search.Query = "Select c_desp Description, d_phdate HolidayDate " & _
                      "from pr_holiday_mst where c_rec_sta='A' "
    End If
    
    Search.CheckFields = "No"
    Search.ReturnField = "No"
    SerVar = Search.Search(, , CON)
    If Len(Search.col1) <> 0 Then
       Txtn_Uno = Search.col1
       Display_Records
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
    
        Set rs = Nothing
        g_Sql = "Select * from pr_holiday_mst where n_uno = '" & Is_Null(Txtn_Uno, True) & "'"
        rs.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
         
        If rs.RecordCount = 0 Then
           rs.AddNew
           Txtn_Uno = Start_Generate_New
           rs("c_usr_id").Value = g_UserName
           rs("d_created").Value = GetDateTime
        Else
           rs("c_musr_id").Value = g_UserName
           rs("d_modified").Value = GetDateTime
        End If
         
        rs("n_uno") = Is_Null(Txtn_Uno, True)
        rs("d_phdate") = Is_Date(Dtp_Date.Text, "S")
        rs("c_desp") = Is_Null(UCase(Txtc_Desp), False)
        rs("c_type").Value = Is_Null(Right(Trim(Cmb_Type), 2), False)
        rs("n_no").Value = Is_Null(Right(Trim(Cmb_No), 1), True)
        rs("c_rec_sta").Value = "A"
        rs.Update
        
    CON.CommitTrans
    
    g_SaveFlagNull = False
    Screen.MousePointer = vbDefault
    
    MsgBox "Record Saved Successfully", vbInformation, "Information"
    Call Btn_Clear_Click
    
Exit Sub
    
ErrSave:
   CON.RollbackTrans
   g_SaveFlagNull = False
   Screen.MousePointer = vbDefault
   MsgBox "Error while Saving - " + Err.Description, vbCritical, "Critical"
   
End Sub

Private Function ChkSave() As Boolean
Dim rsChk As New ADODB.Recordset
  If Not IsDate(Dtp_Date.Text) Then
     MsgBox "Date should not be empty", vbInformation, "Information"
     Dtp_Date.SetFocus
     Exit Function
  ElseIf Trim(Txtc_Desp) = "" Then
     MsgBox "Description should not be empty", vbInformation, "Information"
     Txtc_Desp.SetFocus
     Exit Function
  ElseIf Trim(Cmb_Type) = "" Then
     MsgBox "Type should not be empty", vbInformation, "Information"
     Cmb_Type.SetFocus
     Exit Function
  ElseIf Trim(Cmb_No) = "" Then
     MsgBox "Flag should not be empty", vbInformation, "Information"
     Cmb_No.SetFocus
     Exit Function
  End If
  ChkSave = True
End Function

Private Function Start_Generate_New() As Double
  Dim MaxNo As ADODB.Recordset
  
  g_Sql = "Select max(n_uno) from pr_holiday_mst "
  Set MaxNo = CON.Execute(g_Sql)
  Start_Generate_New = Is_Null(MaxNo(0).Value, True) + 1
  
End Function

Private Sub Display_Records()
  Dim rsDisp As New ADODB.Recordset
  Dim i As Integer
    
    Set rsDisp = Nothing
    g_Sql = "select * from pr_holiday_mst where n_uno = " & Is_Null(Txtn_Uno, True)
    rsDisp.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If rsDisp.RecordCount > 0 Then
       Txtn_Uno = Is_Null(rsDisp("n_uno").Value, True)
       Txtc_Desp = Is_Null(rsDisp("c_desp").Value, False)
       Dtp_Date.Text = Is_Date(rsDisp("d_phdate").Value, "D")
       
       For i = 0 To Cmb_Type.ListCount - 1
           If Right(Trim(Cmb_Type.List(i)), 2) = rsDisp("c_type").Value Then
               Cmb_Type.ListIndex = i
               Exit For
           End If
        Next i
    
       For i = 0 To Cmb_No.ListCount - 1
           If Val(Right(Trim(Cmb_No.List(i)), 1)) = rsDisp("n_no").Value Then
               Cmb_No.ListIndex = i
               Exit For
           End If
        Next i
        
    End If
End Sub

Private Sub Load_Combo()
  
   Cmb_Type.Clear
   Cmb_Type.AddItem "Public Holiday" & Space(100) & "PH"
   Cmb_Type.AddItem "Other Holiday" & Space(100) & "OH"

   Cmb_No.Clear
   Cmb_No.AddItem "First Holiday in a Week" & Space(100) & "1"
   Cmb_No.AddItem "Second Holiday in a Week" & Space(100) & "2"
   Cmb_No.AddItem "Third Holiday in a Week" & Space(100) & "3"
   Cmb_No.AddItem "Fourth Holiday in a Week" & Space(100) & "4"
   Cmb_No.AddItem "Fifth Holiday in a Week" & Space(100) & "5"
   Cmb_No.AddItem "Sixth Holiday in a Week" & Space(100) & "6"
   Cmb_No.AddItem "Seventh Holiday in a Week" & Space(100) & "7"

End Sub

