VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{B9411660-10E6-4A53-BE96-7FED334704FA}#7.0#0"; "fpSpru70.ocx"
Object = "{C3A136DA-B937-492B-968D-A437638F7AAB}#1.0#0"; "CS_DateControl.ocx"
Begin VB.Form frm_Clock_Emp 
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   6840
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3840
   ScaleWidth      =   6840
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   900
      Left            =   120
      TabIndex        =   31
      Top             =   -75
      Width           =   19320
      Begin VB.CommandButton Btn_GridDefault 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Set as Your Default Grid"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   17490
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   195
         Width           =   1440
      End
      Begin VB.CheckBox Chk_Hide 
         Caption         =   "Hide Employee Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   210
         Index           =   0
         Left            =   13215
         TabIndex        =   44
         Top             =   285
         Width           =   2385
      End
      Begin VB.CheckBox Chk_Hide 
         Caption         =   "Hide Clockings"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   210
         Index           =   1
         Left            =   13215
         TabIndex        =   43
         Top             =   540
         Width           =   2325
      End
      Begin VB.CommandButton Cmd_Filter 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   6135
         Picture         =   "frm_Clock_Emp.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Discp 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   6870
         Picture         =   "frm_Clock_Emp.frx":35B0
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Export 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   7860
         Picture         =   "frm_Clock_Emp.frx":6C38
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Hide 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   4410
         Picture         =   "frm_Clock_Emp.frx":A27F
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Restore 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   5145
         Picture         =   "frm_Clock_Emp.frx":D984
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_View 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   990
         Picture         =   "frm_Clock_Emp.frx":1104A
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Save 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   30
         Picture         =   "frm_Clock_Emp.frx":146FA
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Exit 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   8595
         Picture         =   "frm_Clock_Emp.frx":17D83
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Clear 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   1710
         Picture         =   "frm_Clock_Emp.frx":1B3E3
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Delete 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   2700
         Picture         =   "frm_Clock_Emp.frx":1EA53
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Print 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   3420
         Picture         =   "frm_Clock_Emp.frx":220FD
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   120
         Width           =   700
      End
   End
   Begin FPUSpreadADO.fpSpread Va_Details 
      Height          =   7290
      Left            =   120
      TabIndex        =   12
      Top             =   2205
      Width           =   19305
      _Version        =   458752
      _ExtentX        =   34052
      _ExtentY        =   12859
      _StockProps     =   64
      AutoClipboard   =   0   'False
      ColsFrozen      =   2
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
      MaxCols         =   48
      MaxRows         =   50
      ProcessTab      =   -1  'True
      SpreadDesigner  =   "frm_Clock_Emp.frx":256EB
   End
   Begin MSComDlg.CommonDialog ComDialog 
      Left            =   19500
      Top             =   1230
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filter Option"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   975
      Left            =   120
      TabIndex        =   15
      Top             =   1215
      Width           =   19320
      Begin VB.ComboBox Cmb_Shift 
         Height          =   315
         Left            =   13275
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   195
         Width           =   1545
      End
      Begin VB.TextBox Txtc_Month 
         Height          =   300
         Left            =   1125
         TabIndex        =   0
         Top             =   210
         Width           =   690
      End
      Begin VB.ComboBox Cmb_Desig 
         Height          =   315
         Left            =   9660
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   540
         Width           =   2370
      End
      Begin VB.TextBox Txtc_EmployeeName 
         Height          =   300
         Left            =   13275
         TabIndex        =   10
         Top             =   525
         Width           =   4245
      End
      Begin VB.ComboBox Cmb_EmpType 
         Height          =   315
         Left            =   16140
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   195
         Width           =   1380
      End
      Begin VB.TextBox Txtc_Year 
         Height          =   300
         Left            =   1125
         TabIndex        =   1
         Top             =   510
         Width           =   690
      End
      Begin VB.ComboBox Cmb_Company 
         Height          =   315
         Left            =   5895
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   210
         Width           =   2370
      End
      Begin VB.ComboBox Cmb_Dept 
         Height          =   315
         Left            =   9660
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   210
         Width           =   2370
      End
      Begin VB.ComboBox Cmb_Branch 
         Height          =   315
         Left            =   5895
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   525
         Width           =   2370
      End
      Begin VB.CommandButton Btn_Display 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   18015
         Picture         =   "frm_Clock_Emp.frx":26A23
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   165
         Width           =   700
      End
      Begin CS_DateControl.DateControl Dtp_FromDate 
         Height          =   345
         Left            =   2745
         TabIndex        =   2
         Top             =   210
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   609
      End
      Begin CS_DateControl.DateControl Dtp_ToDate 
         Height          =   345
         Left            =   2745
         TabIndex        =   3
         Top             =   540
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   609
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Shift"
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
         Left            =   12825
         TabIndex        =   47
         Top             =   255
         Width           =   375
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Month"
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
         Left            =   510
         TabIndex        =   46
         Top             =   255
         Width           =   525
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Desig"
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
         Left            =   9105
         TabIndex        =   30
         Top             =   585
         Width           =   465
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   12750
         TabIndex        =   29
         Top             =   570
         Width           =   465
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "To"
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
         Left            =   2415
         TabIndex        =   28
         Top             =   600
         Width           =   210
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Emp Type"
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
         Left            =   15240
         TabIndex        =   21
         Top             =   240
         Width           =   810
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
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
         Height          =   210
         Left            =   660
         TabIndex        =   20
         Top             =   555
         Width           =   375
      End
      Begin VB.Label Company 
         AutoSize        =   -1  'True
         Caption         =   "Company"
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
         Left            =   5025
         TabIndex        =   19
         Top             =   255
         Width           =   780
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Dept"
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
         Left            =   9195
         TabIndex        =   18
         Top             =   255
         Width           =   375
      End
      Begin VB.Label Label4 
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
         ForeColor       =   &H00404040&
         Height          =   210
         Left            =   5235
         TabIndex        =   17
         Top             =   570
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "From"
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
         Left            =   2190
         TabIndex        =   16
         Top             =   270
         Width           =   435
      End
   End
   Begin VB.Frame Frm_Filter 
      Caption         =   "Additional Filter Option"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   4605
      Left            =   13110
      TabIndex        =   22
      Top             =   6000
      Width           =   6840
      Begin FPUSpreadADO.fpSpread Va_Filter 
         Height          =   3825
         Left            =   90
         TabIndex        =   26
         Top             =   210
         Width           =   6630
         _Version        =   458752
         _ExtentX        =   11695
         _ExtentY        =   6747
         _StockProps     =   64
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
         SpreadDesigner  =   "frm_Clock_Emp.frx":2A161
      End
      Begin VB.CommandButton Btn_Ok 
         BackColor       =   &H00C0E0FF&
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   540
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   4155
         Width           =   1575
      End
      Begin VB.CommandButton Btn_Cancel 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2610
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   4155
         Width           =   1575
      End
      Begin VB.CommandButton Btn_ClearFilter 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Clear Filter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4680
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   4170
         Width           =   1575
      End
   End
   Begin VB.Label Lbl_Info 
      AutoSize        =   -1  'True
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   120
      TabIndex        =   27
      Top             =   9570
      Width           =   570
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
      Left            =   17535
      TabIndex        =   14
      Top             =   900
      Width           =   360
   End
   Begin VB.Label lbl_scr_name 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Attendance Details"
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
      Left            =   270
      TabIndex        =   13
      Top             =   900
      Width           =   1545
   End
   Begin VB.Shape shp_scr_name 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   120
      Top             =   855
      Width           =   19305
   End
End
Attribute VB_Name = "frm_Clock_Emp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mnuOption As String
Private rs As New ADODB.Recordset
Private tmpFrmActivateCtr As Integer
Private strFilter As String, vLeaveCodes As String
Private SelFor As String, SelFilFor As String, RepHead As String, RepHeadDate As String
Private vF1, vF2, vF3, vF4, vF5 As String
Private vPayPeriod As Long, vPayPeriodFrom As Date, vPayPeriodTo As Date
Private vInBool As Boolean, vOutBool As Boolean

Private Sub Form_Activate()
    Chk_Hide(0).Value = 1
    Chk_Hide(1).Value = 1
    Call Spread_Hide_Check
    tmpFrmActivateCtr = tmpFrmActivateCtr + 1
    
    If Val(Txtc_Month) = 0 Then
       Txtc_Month.SetFocus
    Else
       Txtc_EmployeeName.SetFocus
    End If
End Sub

Private Sub Form_Load()
    ' mnuOption
    ' A  -  Attendance Details
    ' S  -  Shift Change

    If mnuOption = "A" Then
       lbl_scr_name = "Attendance Details"
    Else
       lbl_scr_name = "Attendance Shift Change Details"
    End If
    lbl_date.Caption = Format(Date, "dd-mmm-yyyy")
    Enable_Controls Me, True

    Frm_Filter.Visible = False
    Lbl_Info.Caption = ""

    Clear_Spread Va_Details

    Call Combo_Load
    Call Spread_Lock
    Call TGControlProperty(Me)
    Call SpreadHeaderFont(Va_Details, "Arial", 7, False)
    Call Spread_Row_Height(Va_Filter, 12, 13)
    Call GetLeaveCodes
    Call Create_Default_Clock_Filter
    
   ' Btn_View.Visible = False
   ' Btn_Delete.Visible = False
    Cmb_Company.ListIndex = 1
End Sub

Private Sub Btn_Exit_Click()
    Unload Me
End Sub

Private Sub Btn_Clear_Click()
  Dim i As Integer
  
    Call Clear_Spread(Va_Details)
    Va_Details.MaxRows = 0
    Va_Details.MaxRows = 50
    
    Cmb_Branch.ListIndex = 0
    Cmb_Dept.ListIndex = 0
    Cmb_Desig.ListIndex = 0
    Cmb_Shift.ListIndex = 0
    Cmb_EmpType.ListIndex = 0
    Txtc_EmployeeName = ""

    Frm_Filter.Visible = False
    Lbl_Info.Caption = ""
    
    For i = 1 To Va_Details.MaxRows
        Va_Details.Row = i
        Va_Details.Col = 1
           Va_Details.Value = False
    Next i
    
End Sub


Private Sub Btn_Print_Click()
On Error GoTo Err_Print
  Dim i As Long
  Dim RepOpt As String, tmpStr As String
  
    SelFor = ""
    RepHead = "": RepHeadDate = ""
    vF1 = "": vF2 = "": vF3 = "": vF4 = "": vF5 = ""
    
    If Not IsDate(Dtp_FromDate.Text) And Not IsDate(Dtp_ToDate.Text) Then
       MsgBox "Please Enter Date. Should not be Empty", vbInformation, "Information"
       Dtp_FromDate.SetFocus
       Exit Sub
    End If
    
    
    tmpStr = "1. Attendance Details " & vbCrLf & _
             "2. Attendance Summary " & vbCrLf & _
             "3. Daily Attendance " & vbCrLf & vbCrLf & _
             "4. Over Time Details " & vbCrLf & _
             "5. First In Last Out Report " & vbCrLf & vbCrLf & _
             "6. Discrepancies Details "

    RepOpt = InputBox(tmpStr, "Select your option", "1")
            
    If Trim(Cmb_Company) <> "" Then
       RepHead = Proper(Trim(Left(Trim(Cmb_Company), 50))) & "  -  "
    End If
    If Trim(Cmb_Branch) <> "" Then
       RepHeadDate = "List for " & Proper(Trim(Left(Trim(Cmb_Branch), 50)))
    Else
       RepHeadDate = "List for all Branches "
    End If
        
    Call Pr_Clock_Emp_RepFilter
    
    Call Assign_Add_Filter_Rep
    
    SelFor = ReportFilterOption(SelFor, SelFilFor)
    
     If Val(RepOpt) = 1 Then
        tmpStr = "1. Attendance Details " & vbCrLf & _
                 "2. Attendance Details by Employee "
        RepOpt = InputBox(tmpStr, "Select your option", "1")
        
        If Val(RepOpt) = 1 Then
           RepHead = RepHead & "Attendance Details"
           Call Print_Rpt(SelFor, "Pr_Clock_Attn_Dtl.rpt")
        ElseIf Val(RepOpt) = 2 Then
           RepHead = RepHead & "Attendance Details"
           Call Print_Rpt(SelFor, "Pr_Clock_Daily_Dtl_F3.rpt")
        Else
           Exit Sub
        End If
     
     ElseIf Val(RepOpt) = 2 Then
        tmpStr = "1. Status by Date " & vbCrLf & _
                 "2. Status by Dept "
        RepOpt = InputBox(tmpStr, "Select your option", "1")
        
        If Val(RepOpt) = 1 Then
           RepHead = RepHead & "Clocking Status by Date"
           Call Pr_Clock_Status_RepFilter
           Call Print_Rpt(SelFor, "Pr_Clock_Attn_Br_Sum_F2.rpt")
        ElseIf Val(RepOpt) = 2 Then
           RepHead = RepHead & "Clocking Status by Department"
           Call Pr_Clock_Status_RepFilter
           Call Print_Rpt(SelFor, "Pr_Clock_Attn_Br_Sum.rpt")
        Else
           Exit Sub
        End If
        
     ElseIf Val(RepOpt) = 3 Then
        tmpStr = "1. Attendance Details " & vbCrLf & vbCrLf & _
                 "2. Absent Details " & vbCrLf & _
                 "3. Leave Details " & vbCrLf & vbCrLf & _
                 "4. Time Log Details"
        RepOpt = InputBox(tmpStr, "Select your option", "1")
        
        If Val(RepOpt) = 1 Then
           RepHead = RepHead & "Daily  Attendance  Details"
           Call Print_Rpt(SelFor, "Pr_Clock_Daily_Dtl_F2.rpt")
        ElseIf Val(RepOpt) = 2 Then
           RepHead = RepHead & "Daily  Absentees  Details"
           SelFor = SelFor & " AND {PR_CLOCK_EMP.C_PRESABS} = 'A' "
           Call Print_Rpt(SelFor, "Pr_Clock_Daily_Dtl.rpt")
        ElseIf Val(RepOpt) = 3 Then
           tmpStr = InputBox("Please input Leave Code", "Leave Code")
           RepHead = RepHead & "Daily  " & Proper(GetLeaveName(tmpStr)) & "  Details"
           SelFor = SelFor & " AND {PR_CLOCK_EMP.C_PRESABS} = '" & Trim(tmpStr) & "'"
           Call Print_Rpt(SelFor, "Pr_Clock_Daily_Dtl.rpt")
        ElseIf Val(RepOpt) = 4 Then
           RepHead = RepHead & "Daily  Time  Log  Details"
           Call Print_Rpt(SelFor, "Pr_Clock_Log_Dtl.rpt")
        Else
           Exit Sub
        End If
        
     ElseIf Val(RepOpt) = 4 Then
        tmpStr = "1. Over Time Details " & vbCrLf & _
                 "2. Approved Over Time Details "
        
        RepOpt = InputBox(tmpStr, "Select your option", "1")
        
        If Val(RepOpt) = 1 Then
           SelFor = SelFor & " AND {PR_CLOCK_EMP.N_OVERTIME} > 0 "
        ElseIf Val(RepOpt) = 2 Then
           SelFor = SelFor & " AND ({PR_CLOCK_EMP.N_OTIN_STR} = '$' OR {PR_CLOCK_EMP.N_OTOUT_STR} = '$') AND ({PR_CLOCK_EMP.N_OT15}+{PR_CLOCK_EMP.N_OT20}+{PR_CLOCK_EMP.N_OT30}) > 0 "
        Else
           Exit Sub
        End If
        
        RepHead = RepHead & "Daily Over Time Details"
        Call Print_Rpt(SelFor, "Pr_Clock_Attn_Dtl_F2.rpt")
        Mdi_Ta_HrPay.CRY1.Formulas(3) = "PeriodFrom = '" & Is_Date(Dtp_FromDate.Text, "D") & "'"
        Mdi_Ta_HrPay.CRY1.Formulas(4) = "PeriodTo = '" & Is_Date(Dtp_ToDate.Text, "D") & "'"
        
     ElseIf Val(RepOpt) = 5 Then
        RepHead = "FIRST IN LAST OUT REPORT"
        Call Pr_Clock_FILO_RepFilter
        Call Print_Rpt(SelFor, "Pr_Clock_Attn_FILO.rpt")
        
     ElseIf Val(RepOpt) = 6 Then
        Call Save_Pr_Clock_Tran_Discp
        
        RepHead = "Unit : " & Trim(Left(Cmb_Branch, 50)) & Space(15) & "Period : " & Format(Txtc_Month, "00") & "-" & Trim(Format(Txtc_Year, "0000"))
        Call Print_Rpt("", "Pr_Clock_Emp_Discp.rpt")
        
     Else
        Exit Sub
     End If

    If Trim(RepHead) <> "" Then
       Mdi_Ta_HrPay.CRY1.Formulas(1) = "ReportHead='" & UCase(RepHead) & "'"
    End If

    If Trim(RepHeadDate) <> "" Then
       Mdi_Ta_HrPay.CRY1.Formulas(2) = "RepHeadDate='" & RepHeadDate & "'"
    End If
    Mdi_Ta_HrPay.CRY1.Action = 1
  
 Exit Sub

Err_Print:
    Screen.MousePointer = vbDefault
    MsgBox "Error while Generating - " + Err.Description, vbInformation, "Information"
End Sub

Private Sub Btn_Save_Click()
On Error GoTo Err_Flag
  
   If Not Chk_Mand Then Exit Sub
   
   If ChkSave = False Then
      Exit Sub
   End If
   
   If mnuOption = "A" Then
      Call Re_Calculate_WorkHrs
      Call Assign_OTSign
   End If
   
   If MsgBox("Do you want to Save?", vbYesNo, "Confirmation") = vbYes Then
   
      Screen.MousePointer = vbHourglass
      g_SaveFlagNull = True
      CON.BeginTrans
      
      If mnuOption = "A" Then
         Call Update_PR_Clock_Emp
      Else
         ' Shift update
         Call Update_PR_Clock_Emp_User
      End If
      
      CON.CommitTrans
      Screen.MousePointer = vbDefault
      g_SaveFlagNull = False
   Else
      Exit Sub
   End If
   
   MsgBox "Saved Successfully", vbInformation, "Information"
   
 Exit Sub

Err_Flag:
   CON.RollbackTrans
   Screen.MousePointer = vbDefault
   g_SaveFlagNull = False
   MsgBox "Error While Save - " & Err.Description
End Sub

Private Function ChkSave() As Boolean
  Dim i As Long
  Dim tmpEmpNo As String
    
    If Trim(Txtc_Month) <> "" And Trim(Txtc_Year) <> "" Then
       vPayPeriod = Is_Null(Format(Trim(Txtc_Year), "0000") & Trim(Txtc_Month), True)
       If Not ChkPeriodOpen(vPayPeriod, "W") Then
          Txtc_Month.SetFocus
          Exit Function
       End If
    End If
    
    For i = 1 To Va_Details.DataRowCnt
        Va_Details.Row = i
        Va_Details.Col = 2
           tmpEmpNo = UCase(Trim(Va_Details.Text))
        Va_Details.Col = 1
        If Trim(tmpEmpNo) <> "" And Va_Details.Value = 1 And mnuOption = "A" Then
           Va_Details.Col = 20
              If Trim(Va_Details.Text) = "" Then
                 MsgBox "Present Status should not be Empty", vbInformation, "Information"
                 Va_Details.SetFocus
                 Exit Function
              End If
        End If
    Next

   ChkSave = True
End Function

Private Sub Update_PR_Clock_Emp()
  Dim i As Integer
  Dim tmpEmpNo As String, tmpDate As String
  Dim tmpArrTime As Double, tmpDepTime As Double, tmpWorkHrs As Double, tmpOtHrs As Double
  Dim tmpPermHrs As Double, tmpLateHrs As Double, tmpEarlyHrs As Double, tmpPresent As Double
  Dim tmpOtIn As Double, tmpOtOut As Double, tmpOt15 As Double, tmpOt20 As Double, tmpOt30 As Double
  Dim tmpPresAbs As String, tmpShift As String, tmpFlag As String, tmpOtInStr As String, tmpOtOutStr As String

  
    For i = 1 To Va_Details.DataRowCnt
        tmpEmpNo = "": tmpDate = ""
        tmpArrTime = 0: tmpDepTime = 0: tmpWorkHrs = 0: tmpOtHrs = 0: tmpPermHrs = 0: tmpLateHrs = 0: tmpEarlyHrs = 0: tmpPresent = 0
        tmpOtIn = 0: tmpOtOut = 0: tmpOt15 = 0: tmpOt20 = 0: tmpOt30 = 0
    
        Va_Details.Row = i
        Va_Details.Col = 2
           tmpEmpNo = UCase(Trim(Va_Details.Text))
        Va_Details.Col = 10
           If IsDate(Va_Details.Text) Then
              tmpDate = Trim(Va_Details.Text)
           End If
           
        Va_Details.Col = 1
        If Trim(tmpEmpNo) <> "" And tmpDate <> "" And Va_Details.Value = 1 Then
           Va_Details.Col = 2
              tmpEmpNo = UCase(Trim(Va_Details.Text))
           Va_Details.Col = 10
              tmpDate = Format(CDate(Va_Details.Text), "yyyy-mm-dd")
           Va_Details.Col = 12
              tmpShift = UCase(Trim(Va_Details.Text))
            
           Va_Details.Col = 13
              tmpArrTime = Val(Va_Details.Text)
           Va_Details.Col = 14
              tmpLateHrs = Val(Va_Details.Text)
           Va_Details.Col = 15
              tmpDepTime = Val(Va_Details.Text)
           Va_Details.Col = 16
              tmpEarlyHrs = Val(Va_Details.Text)
           Va_Details.Col = 17
              tmpWorkHrs = Val(Va_Details.Text)
           Va_Details.Col = 18
              tmpOtHrs = Val(Va_Details.Text)
           Va_Details.Col = 19
              tmpPermHrs = Val(Va_Details.Text)
           
           Va_Details.Col = 20
              tmpPresAbs = Trim(Va_Details.Text)
           Va_Details.Col = 21
              tmpPresent = Val(Va_Details.Text)
           
           Va_Details.Col = 23
              tmpOtIn = Val(Va_Details.Text)
              tmpOtInStr = Trim(Va_Details.Text)
              If Trim(Right(tmpOtInStr, 1)) = "$" Then
                 tmpOtInStr = "$"
              Else
                 tmpOtInStr = ""
              End If
              
           Va_Details.Col = 24
              tmpOtOut = Val(Va_Details.Text)
              tmpOtOutStr = Trim(Va_Details.Text)
              If Trim(Right(tmpOtOutStr, 1)) = "$" Then
                 tmpOtOutStr = "$"
              Else
                 tmpOtOutStr = ""
              End If
              
           Va_Details.Col = 25
              tmpOt15 = Val(Va_Details.Text)
           Va_Details.Col = 26
              tmpOt20 = Val(Va_Details.Text)
           Va_Details.Col = 27
              tmpOt30 = Val(Va_Details.Text)
                
           tmpFlag = "U"
           
           g_Sql = "update pr_clock_emp set " & _
                           "n_period = " & vPayPeriod & ", " & _
                           "n_arrtime = " & tmpArrTime & ", n_arrtime_dec = " & MinsToDecimal(tmpArrTime) & ", n_arrtime_min = " & TimeToMins(tmpArrTime) & ", " & _
                           "n_latehrs = " & tmpLateHrs & ", n_latehrs_dec = " & MinsToDecimal(tmpLateHrs) & ", n_latehrs_min = " & TimeToMins(tmpLateHrs) & ", " & _
                           "n_deptime = " & tmpDepTime & ", n_deptime_dec = " & MinsToDecimal(tmpDepTime) & ", n_deptime_min = " & TimeToMins(tmpDepTime) & ", " & _
                           "n_earlhrs = " & tmpEarlyHrs & ", n_earlhrs_dec = " & MinsToDecimal(tmpEarlyHrs) & ", n_earlhrs_min = " & TimeToMins(tmpEarlyHrs) & ", " & _
                           "n_permhrs = " & tmpPermHrs & ", n_permhrs_dec = " & MinsToDecimal(tmpPermHrs) & ", n_permhrs_min = " & TimeToMins(tmpPermHrs) & ", " & _
                           "n_workhrs = " & tmpWorkHrs & ", n_workhrs_dec = " & MinsToDecimal(tmpWorkHrs) & ", n_workhrs_min = " & TimeToMins(tmpWorkHrs) & ", " & _
                           "n_overtime = " & tmpOtHrs & ", n_overtime_dec = " & MinsToDecimal(tmpOtHrs) & ", n_overtime_min = " & TimeToMins(tmpOtHrs) & ", " & _
                           "n_otin = " & tmpOtIn & ", n_otin_dec = " & MinsToDecimal(tmpOtIn) & ", n_otin_min = " & TimeToMins(tmpOtIn) & ", " & _
                           "n_otout = " & tmpOtOut & ", n_otout_dec = " & MinsToDecimal(tmpOtOut) & ", n_otout_min = " & TimeToMins(tmpOtOut) & ", " & _
                           "n_ot15 = " & tmpOt15 & ", n_ot15_dec = " & MinsToDecimal(tmpOt15) & ", n_ot15_min = " & TimeToMins(tmpOt15) & ", " & _
                           "n_ot20 = " & tmpOt20 & ", n_ot20_dec = " & MinsToDecimal(tmpOt20) & ", n_ot20_min = " & TimeToMins(tmpOt20) & ", " & _
                           "n_ot30 = " & tmpOt30 & ", n_ot30_dec = " & MinsToDecimal(tmpOt30) & ", n_ot30_min = " & TimeToMins(tmpOt30) & ", " & _
                           "n_otin_str = '" & Is_Null(tmpOtInStr, False) & "', n_otout_str = '" & Is_Null(tmpOtOutStr, False) & "', " & _
                           "n_present = " & tmpPresent & ", c_presabs = '" & tmpPresAbs & "', " & _
                           "c_chq = '', c_shift = '" & tmpShift & "', c_flag = '" & tmpFlag & "', " & _
                           "c_usr_id = '" & g_UserName & "', d_modified = '" & GetDateTime & "' " & _
                   "where c_empno = '" & tmpEmpNo & "' and d_date = '" & tmpDate & "'"
                   
            
           CON.Execute g_Sql
        End If
    Next i
End Sub

Private Sub Update_PR_Clock_Emp_User()
  Dim i As Integer
  Dim tmpEmpNo As String, tmpDate As String
  Dim tmpArrTime As Double, tmpDepTime As Double, tmpWorkHrs As Double, tmpPresent As Double
  Dim tmpPresAbs As String, tmpShift As String, tmpFlag As String
  
    For i = 1 To Va_Details.DataRowCnt
        Va_Details.Row = i
        Va_Details.Col = 2
           tmpEmpNo = UCase(Trim(Va_Details.Text))
        Va_Details.Col = 1
        
        If Trim(tmpEmpNo) <> "" And Va_Details.Value = 1 Then
           Va_Details.Col = 10
              tmpDate = Format(CDate(Va_Details.Text), "yyyy-mm-dd")
           Va_Details.Col = 12
              tmpShift = UCase(Trim(Va_Details.Text))
           
           g_Sql = "update pr_clock_emp set " & _
                           "n_period = " & vPayPeriod & ", " & _
                           "c_shift = '" & tmpShift & "', c_sh_flag = 'U' " & _
                   "where c_empno = '" & tmpEmpNo & "' and d_date = '" & tmpDate & "'"
             
           CON.Execute g_Sql
        End If
    Next i
End Sub

Private Sub Btn_Display_Click()
  
  If Trim(Txtc_EmployeeName) = "" Then
     If Not IsDate(Dtp_FromDate.Text) Then
        MsgBox "Date From should not be Empty", vbInformation, "Information"
        Dtp_FromDate.SetFocus
        Exit Sub
     ElseIf Not IsDate(Dtp_ToDate.Text) Then
        MsgBox "Date From should not be Empty", vbInformation, "Information"
        Dtp_ToDate.SetFocus
        Exit Sub
     End If
     
     If CDate(Dtp_ToDate.Text) - CDate(Dtp_FromDate.Text) > 40 Then
        MsgBox "You choosen big date range. Please choose one month period", vbInformation, "Information"
        Dtp_ToDate.SetFocus
        Exit Sub
     End If
  End If
  
  If Va_Details.DataRowCnt > 0 Then
     If MsgBox("The Data will be refesh?. Do you want to display?", vbYesNo + vbDefaultButton2, "Confirmation") = vbNo Then
        Exit Sub
     Else
        Va_Details.MaxRows = 0
        Va_Details.MaxRows = 50
     End If
  End If
  
  If Trim(Txtc_EmployeeName) = "" Then
     Lbl_Info.Caption = ""
  End If
  
  Call Display_Records
End Sub

Private Sub Display_Records(Optional vRow As Long)
On Error GoTo Err_Display
  Dim DyDisp As New ADODB.Recordset
  Dim i As Long, j As Long
  Dim tmpPeriod As String, tmpEmpNo As String
  
  Screen.MousePointer = vbHourglass
  
  g_Sql = "select a.c_empno, c.c_companyname, b.c_branch, b.c_dept, b.c_desig, b.c_emptype, " & _
          "a.d_date, a.n_wkday, a.n_arrtime, a.n_latehrs, a.n_deptime, a.n_earlhrs, a.n_permhrs, " & _
          "a.n_workhrs, a.n_overtime, a.n_present, a.c_presabs, a.c_shift, a.n_workhrs_fp, a.n_canteen, " & _
          "a.c_chq, a.c_flag, a.c_usr_id, a.d_modified, b.c_name, b.c_othername, a.c_remarks, " & _
          "a.n_otin, a.n_otout, a.n_ot15, a.n_ot20, a.n_ot30, a.n_otin_str, a.n_otout_str, " & _
          "a.n_time1, a.n_time2, a.n_time3, a.n_time4, a.n_time5, a.n_time6, " & _
          "a.n_time7, a.n_time8, a.n_time9, a.n_time10, a.n_time11, a.n_time12 " & _
          "from pr_clock_emp a,  pr_emp_mst b, pr_company_mst c " & _
          "where a.c_empno = b.c_empno and b.c_company = c.c_company and b.c_rec_sta = 'A' and (a.d_date <= b.d_dol or b.d_dol is null)  "
  
  If vRow > 0 Then
     Va_Details.Row = vRow
     Va_Details.Col = 2
        If Trim(Va_Details.Text) <> "" Then
           g_Sql = g_Sql & " and a.c_empno = '" & Trim(Va_Details.Text) & "'"
        Else
           Exit Sub
        End If
     Va_Details.Col = 4
        If IsDate(Va_Details.Text) Then
           g_Sql = g_Sql & " and a.d_date = '" & Is_Date(Va_Details.Text, "S") & "'"
        Else
           Exit Sub
        End If
  Else
     If IsDate(Dtp_FromDate.Text) And IsDate(Dtp_ToDate.Text) Then
        g_Sql = g_Sql & " and a.d_date >= '" & Is_Date(Dtp_FromDate.Text, "S") & "' and a.d_date <= '" & Is_Date(Dtp_ToDate.Text, "S") & "' "
     End If
     
     If Right(Trim(Txtc_EmployeeName), 7) <> "" Then
        g_Sql = g_Sql & " and a.c_empno = '" & Trim(Right(Trim(Txtc_EmployeeName), 7)) & "'"
     Else
        If Trim(Cmb_Company) <> "" Then
           g_Sql = g_Sql & " and b.c_company = '" & Trim(Right(Trim(Cmb_Company), 7)) & "'"
        End If
        
        If Trim(Cmb_Branch) <> "" Then
           g_Sql = g_Sql & " and b.c_branch = '" & Trim(Cmb_Branch) & "'"
        End If
                
        If Trim(Cmb_Dept) <> "" Then
           g_Sql = g_Sql & " and b.c_dept = '" & Trim(Cmb_Dept) & "'"
        End If
        
        If Trim(Cmb_Desig) <> "" Then
           g_Sql = g_Sql & " and b.c_desig = '" & Trim(Cmb_Desig) & "'"
        End If
        
        If Trim(Cmb_Shift) <> "" Then
           g_Sql = g_Sql & " and b.c_shiftcode = '" & Trim(Right(Trim(Cmb_Shift), 7)) & "'"
        End If
        
        If Trim(Cmb_EmpType) <> "" Then
           g_Sql = g_Sql & " and b.c_emptype = '" & Trim(Cmb_EmpType) & "'"
        End If
     End If
  End If
  
  If vRow = 0 Then
     g_Sql = g_Sql & strFilter
  End If
  
  g_Sql = g_Sql & " order by a.c_empno, a.d_date"
  
  If vRow = 0 Then
     Clear_Spread Va_Details
  End If
  
  Set DyDisp = Nothing
  DyDisp.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
  If vRow = 0 Then
     Va_Details.MaxRows = DyDisp.RecordCount + 10000
  End If
  j = 0
  If DyDisp.RecordCount > 0 Then
     DyDisp.MoveFirst
     For i = 1 To DyDisp.RecordCount
         If vRow > 0 Then
            j = vRow
         Else
            j = j + 1
            If Is_Null(DyDisp("n_wkday").Value, True) = 1 Or tmpEmpNo <> Is_Null(DyDisp("c_empno").Value, False) Then
               If j > 1 Then
                  j = j + 1
               End If
            End If
         End If
         
         If tmpEmpNo <> Is_Null(DyDisp("c_empno").Value, False) Then
            Call GetLeaveOpenBalance(Is_Null(DyDisp("c_empno").Value, False), j)
         End If
         
         Va_Details.Row = j
         Va_Details.Col = 2
            Va_Details.Text = Is_Null(DyDisp("c_empno").Value, False)
         Va_Details.Col = 3
            Va_Details.Text = Proper(Is_Null(DyDisp("c_name").Value, False)) & " " & Proper(Is_Null(DyDisp("c_othername").Value, False))
         Va_Details.Col = 4
            Va_Details.Text = Proper(Is_Null(DyDisp("c_companyname").Value, False))
         Va_Details.Col = 5
            Va_Details.Text = Proper(Is_Null(DyDisp("c_branch").Value, False))
         Va_Details.Col = 6
            Va_Details.Text = Proper(Is_Null(DyDisp("c_dept").Value, False))
         Va_Details.Col = 7
            Va_Details.Text = Proper(Is_Null(DyDisp("c_desig").Value, False))
         Va_Details.Col = 8
            Va_Details.Text = Proper(Is_Null(DyDisp("c_emptype").Value, False))
         Va_Details.Col = 9
            Va_Details.Text = ""
         
         Va_Details.Col = 10
            Va_Details.Text = Is_DateSpread(DyDisp("d_date").Value, True)
         Va_Details.Col = 11
            Va_Details.Text = WeekdayName(Is_Null(DyDisp("n_wkday").Value, True), True, vbMonday)

         Va_Details.Col = 12
            Va_Details.Text = Is_Null(DyDisp("c_shift").Value, False)
         
         Va_Details.Col = 13
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_arrtime").Value, True), True)
         Va_Details.Col = 14
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_latehrs").Value, True), True)
         
         Va_Details.Col = 15
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_deptime").Value, True), True)
         Va_Details.Col = 16
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_earlhrs").Value, True), True)
         
         Va_Details.Col = 17
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_workhrs").Value, True), True)
         Va_Details.Col = 18
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_overtime").Value, True), True)
         Va_Details.Col = 19
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_permhrs").Value, True), True)
         
         Va_Details.Col = 20
            Va_Details.Text = Is_Null(DyDisp("c_presabs").Value, False)
         Va_Details.Col = 21
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_present").Value, True), True)
         Va_Details.Col = 22
            Va_Details.Text = ""
         
         Va_Details.Col = 23
            Va_Details.Text = OTSignAddOnDisplay(DecimalToString(Is_Null(DyDisp("n_otin").Value, True)), Is_Null(DyDisp("n_otin_str").Value, False))
            If Trim(Right(Trim(Va_Details.Text), 1)) = "$" Then
               Va_Details.FontBold = True
            Else
               Va_Details.FontBold = False
            End If
         Va_Details.Col = 24
            Va_Details.Text = OTSignAddOnDisplay(DecimalToString(Is_Null(DyDisp("n_otout").Value, True)), Is_Null(DyDisp("n_otout_str").Value, False))
            If Trim(Right(Trim(Va_Details.Text), 1)) = "$" Then
               Va_Details.FontBold = True
            Else
               Va_Details.FontBold = False
            End If
         Va_Details.Col = 25
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_ot15").Value, True), True)
         Va_Details.Col = 26
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_ot20").Value, True), True)
         Va_Details.Col = 27
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_ot30").Value, True), True)

         
         Va_Details.Col = 29
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_time1").Value, True), True)
         Va_Details.Col = 30
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_time2").Value, True), True)
         Va_Details.Col = 31
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_time3").Value, True), True)
         Va_Details.Col = 32
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_time4").Value, True), True)
         Va_Details.Col = 33
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_time5").Value, True), True)
         Va_Details.Col = 34
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_time6").Value, True), True)
         
         Va_Details.Col = 35
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_time7").Value, True), True)
         Va_Details.Col = 36
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_time8").Value, True), True)
         Va_Details.Col = 37
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_time9").Value, True), True)
         Va_Details.Col = 38
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_time10").Value, True), True)
         Va_Details.Col = 39
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_time11").Value, True), True)
         Va_Details.Col = 40
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_time12").Value, True), True)
         Va_Details.Col = 41
            Va_Details.Text = ""
         
         Va_Details.Col = 42
            If Is_Null(DyDisp("c_chq").Value, False) = "*" Then
               Va_Details.Text = "Single Clocking"
            ElseIf Is_Null(DyDisp("c_flag").Value, False) = "D" Then  'discrepancy import
               Va_Details.Text = Is_Null(DyDisp("c_remarks").Value, False)
            Else
             '  Va_Details.Text = Is_Null(DyDisp("c_chq").Value, False)
            End If
         
         Va_Details.Col = 43
            If Is_Null(DyDisp("c_flag").Value, False) = "U" Then
               Va_Details.Text = Proper(Is_Null(DyDisp("c_usr_id").Value, False))
            End If
         Va_Details.Col = 44
            If Is_Null(DyDisp("c_flag").Value, False) = "U" Then
               Va_Details.Text = Format(DyDisp("d_modified").Value, "dd/mm/yy hh:mm:ss")
            End If
         
         tmpEmpNo = Is_Null(DyDisp("c_empno").Value, False)
        DyDisp.MoveNext
     Next i
     
     If vRow = 0 Then
        Va_Details.MaxRows = Va_Details.DataRowCnt
     End If
     DoEvents
  Else
    ' MsgBox "No details found", vbInformation, "Information"
  End If
  
 Screen.MousePointer = vbDefault

 Exit Sub

Err_Display:
    Screen.MousePointer = vbDefault
    MsgBox "Error While Display " & Err.Description
End Sub


Private Sub Txtc_EmployeeName_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim Search As New Search.MyClass, SerVar
    
   If KeyCode = vbKeyDelete Then
      Txtc_EmployeeName.Text = ""
   End If
 
   If KeyCode = vbKeyF2 Then
      Search.Query = "select c_empno EmpNo, c_name Name, c_othername OtherName, c_dept Dept, " & _
                   "c_desig Desig, c_branch Branch, c_emptype Type " & _
                   "from pr_emp_mst where c_rec_sta = 'A'"
      Search.CheckFields = "EmpNo, Name"
      Search.ReturnField = "EmpNo, Name"
      SerVar = Search.Search(, , CON)
      If Len(Search.col1) <> 0 Then
         Txtc_EmployeeName = Search.col2 & Space(100) & Search.col1
      End If
   End If
End Sub

Private Sub Txtc_EmployeeName_Validate(Cancel As Boolean)
 Dim rsChk As New ADODB.Recordset
 Dim i As Integer
  If Trim(Txtc_EmployeeName) <> "" Then
     Set rsChk = Nothing
     g_Sql = "select c_empno, c_name, c_othername, c_company, c_branch, c_dept, c_desig, c_emptype, " & _
             "d_doj, d_dol, c_stafftype, c_daywork from pr_emp_mst " & _
             "where c_rec_sta = 'A' and c_empno = '" & Trim(Right(Trim(Txtc_EmployeeName), 7)) & "'"
     rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
     If rsChk.RecordCount > 0 Then
        Txtc_EmployeeName = Is_Null(rsChk("c_name").Value, False) & " " & Is_Null(rsChk("c_othername").Value, False) & Space(100) & Is_Null(rsChk("c_empno").Value, False)
'        Call DisplayComboCompany(Me, Is_Null(rsChk("c_company").Value, False))
'        Call DisplayComboBranch(Me, Is_Null(rsChk("c_branch").Value, False))
'        Call DisplayComboDept(Me, Is_Null(rsChk("c_dept").Value, False))
'        Call DisplayComboDesig(Me, Is_Null(rsChk("c_desig").Value, False))
'
'        For i = 0 To Cmb_EmpType.ListCount - 1
'          If Trim(Cmb_EmpType.List(i)) = Is_Null(rsChk("c_emptype").Value, False) Then
'             Cmb_EmpType.ListIndex = i
'             Exit For
'          End If
'        Next i
        
        Lbl_Info.Caption = "Date Join : " & Format(rsChk("d_doj"), "dd/mm/yy")
        Lbl_Info.Caption = Lbl_Info.Caption & Space(10)
        Lbl_Info.Caption = Lbl_Info.Caption & "Date left : " & Format(rsChk("d_dol"), "dd/mm/yy")
        Lbl_Info.Caption = Lbl_Info.Caption & Space(10)
        If Is_Null(rsChk("c_stafftype").Value, False) = "F" Then
           Lbl_Info.Caption = Lbl_Info.Caption & "Flat "
        ElseIf Is_Null(rsChk("c_stafftype").Value, False) = "O" Then
           Lbl_Info.Caption = Lbl_Info.Caption & "Over Time "
        Else
           Lbl_Info.Caption = Lbl_Info.Caption & "Not Defined"
        End If
        Lbl_Info.Caption = Lbl_Info.Caption & Space(10)
        Lbl_Info.Caption = Lbl_Info.Caption & rsChk("c_daywork")
     Else
        MsgBox "Employee not found. Press <F2> to select.", vbInformation, "Information"
        Cancel = True
     End If
  End If
End Sub

Private Sub Spread_Lock()
 Dim i As Long
   
   If mnuOption = "A" Then
      For i = 1 To Va_Details.MaxCols
          Va_Details.Row = -1
          Va_Details.Col = i
             If g_Admin Or g_FrmSupUser Then
                If i = 1 Or i = 12 Or i = 13 Or i = 15 Or i = 17 Or i = 19 Or i = 20 Or i = 21 Or i = 22 Then
                   Va_Details.Lock = False
                Else
                   Va_Details.Lock = True
                End If
             Else
                Va_Details.Lock = True
             End If
      Next i
   Else
      For i = 1 To Va_Details.MaxCols
          Va_Details.Row = -1
          Va_Details.Col = i
             If i = 12 Then
                Va_Details.Lock = False
             Else
                Va_Details.Lock = True
             End If
      Next i
   End If
End Sub

Private Sub Txtc_Month_KeyPress(KeyAscii As Integer)
    Call OnlyNumeric(Txtc_Month, KeyAscii, 2)
End Sub

Private Sub Txtc_Month_Validate(Cancel As Boolean)
    If Trim(Txtc_Month) <> "" Then
       Call MakeMonthTwoDigits(Me)
       If (Val(Txtc_Month) <= 0 Or Val(Txtc_Month) > 12) Then
          MsgBox "Not a valid month", vbInformation, "Information"
          Txtc_Month.SetFocus
          Cancel = True
          Exit Sub
       End If
    End If
    
    If Trim(Txtc_Month) <> "" And Trim(Txtc_Year) <> "" Then
       vPayPeriod = Is_Null(Format(Trim(Txtc_Year), "0000") & Trim(Txtc_Month), True)
       If Not ChkPeriodExists(vPayPeriod, "W") Then
          Txtc_Month.SetFocus
          Cancel = True
          Exit Sub
       Else
          Call Assign_PayPeriodDate
          Dtp_FromDate.Text = Is_Date(vPayPeriodFrom, "D")
          Dtp_ToDate.Text = Is_Date(vPayPeriodTo, "D")
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
  
  If Trim(Txtc_Month) <> "" And Trim(Txtc_Year) <> "" Then
     vPayPeriod = Format(Trim(Txtc_Year), "0000") & Format(Trim(Txtc_Month), "00")
     If Not ChkPeriodExists(Is_Null(vPayPeriod, True), "W") Then
        Txtc_Year.SetFocus
        Cancel = True
     Else
        Call Assign_PayPeriodDate
        Dtp_FromDate.Text = Is_Date(vPayPeriodFrom, "D")
        Dtp_ToDate.Text = Is_Date(vPayPeriodTo, "D")
     End If
  End If
End Sub

Private Sub Assign_PayPeriodDate()
  Dim rsChk As New ADODB.Recordset
  
    vPayPeriod = Is_Null(Format(Trim(Txtc_Year), "0000") & Trim(Txtc_Month), True)
     
    Set rsChk = Nothing
    g_Sql = "select d_fromdate, d_todate from pr_payperiod_dtl " & _
            "where n_period = " & vPayPeriod & " and c_type = 'W' "
    rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If rsChk.RecordCount > 0 Then
       vPayPeriodFrom = rsChk("d_fromdate").Value
       vPayPeriodTo = rsChk("d_todate").Value
    End If
End Sub


Private Sub Va_Details_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
  Dim i As Integer
  Dim tmpStr As String
  
    If BlockRow = -1 Then
       Exit Sub
    End If
     
    If BlockCol = 12 And mnuOption = "S" Then
       If MsgBox("Do you want to change", vbOKCancel + vbDefaultButton2, "Information") = vbOK Then
          Va_Details.Row = BlockRow
          Va_Details.Col = BlockCol
             tmpStr = Trim(Va_Details.Text)
             
          For i = BlockRow To BlockRow2
              Va_Details.Col = BlockCol
              Va_Details.Row = i
                 If Va_Details.Lock = False Then
                    Va_Details.Text = tmpStr
                    Va_Details.Col = 1
                       Va_Details.Value = 1
                 End If
          Next i
       End If
    End If
    
    If (BlockCol = 12 Or BlockCol = 13 Or BlockCol = 15 Or BlockCol = 17 Or BlockCol = 20) And mnuOption = "A" Then
       If MsgBox("Do you want to change", vbOKCancel + vbDefaultButton2, "Information") = vbOK Then
          Va_Details.Row = BlockRow
          Va_Details.Col = BlockCol
             tmpStr = Trim(Va_Details.Text)
             
          For i = BlockRow To BlockRow2
              Va_Details.Col = BlockCol
              Va_Details.Row = i
                 Va_Details.Text = tmpStr
                 Va_Details.Col = 1
                    Va_Details.Value = 1
          Next i
       End If
    End If
End Sub

Private Sub Va_Details_Change(ByVal Col As Long, ByVal Row As Long)
  Dim rsChk As New ADODB.Recordset
  Dim tmpPresAbs As String, tmpEmpNo As String
  Dim vLeaveHrs As String
  
    If (Col = 2 Or Col = 10) Then
       Va_Details.Row = Row
       Va_Details.Col = 10
          If IsDate(Va_Details.Text) Then
             If CDate(Va_Details.Text) >= vPayPeriodFrom And CDate(Va_Details.Text) <= vPayPeriodTo Then
             Else
                Va_Details.Text = ""
                MsgBox "Date should be between Pay Periods", vbInformation, "Information"
                Exit Sub
             End If
          End If
       Va_Details.Col = 2
          If Trim(Va_Details.Text) <> "" Then
             Va_Details.Col = 10
                If IsDate(Va_Details.Text) Then
                   Call Display_Records(Row)
                End If
          End If
    End If
  
    If Col = 12 Then
       Va_Details.Row = Row
       Va_Details.Col = Col
          If Trim(Va_Details.Text) <> "" Then
             Set rsChk = Nothing
             g_Sql = "select c_shiftcode from pr_clock_shift where c_shiftcode = '" & Trim(Va_Details.Text) & "'"
             rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
             If rsChk.RecordCount <= 0 Then
                Va_Details.Text = ""
                MsgBox "Shift code not found in the master. Press <F2> to View ", vbInformation, "Information"
             Else
                Va_Details.Text = UCase(Va_Details.Text)
             End If
          End If
    End If
    
    If Col >= 12 And Col <= 24 Then
       Va_Details.Row = Row
       Va_Details.Col = 1
          Va_Details.Value = 1
    End If
    
End Sub


Private Sub Va_Details_Click(ByVal Col As Long, ByVal Row As Long)
  Dim tmpStr As String, tmpOld As String
  
    If Col = 23 Then
       vInBool = Not vInBool
       Va_Details.Row = Row
       Va_Details.Col = Col
          tmpOld = Trim(Va_Details.Text)
          tmpStr = OTSignRemove(Trim(Va_Details.Text))
          If vInBool Then
             Va_Details.Text = OTSignAdd(tmpStr)
             Va_Details.FontBold = True
          Else
             Va_Details.Text = tmpStr
             Va_Details.FontBold = False
          End If
          
          If tmpOld <> Trim(Va_Details.Text) Then
             Va_Details.Col = 1
                Va_Details.Value = 1
             Call Calculate_WorkHrs(Row)
          End If
    
    ElseIf Col = 24 Then
       vOutBool = Not vOutBool
       Va_Details.Row = Row
       Va_Details.Col = Col
          tmpOld = Trim(Va_Details.Text)
          tmpStr = OTSignRemove(Trim(Va_Details.Text))
          If vOutBool Then
             Va_Details.Text = OTSignAdd(tmpStr)
             Va_Details.FontBold = True
             Va_Details.Col = 1
                Va_Details.Value = 1
          Else
             Va_Details.Text = tmpStr
             Va_Details.FontBold = False
          End If
          
          If tmpOld <> Trim(Va_Details.Text) Then
             Va_Details.Col = 1
                Va_Details.Value = 1
             Call Calculate_WorkHrs(Row)
          End If
    End If
End Sub

Private Sub Va_Details_DblClick(ByVal Col As Long, ByVal Row As Long)
   If Row = 0 Then
      Call SpreadColSort(Va_Details, Col, Row)
   End If
End Sub

Private Sub Va_Details_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim Search As New Search.MyClass, SerArray, SerVar
  
  If KeyCode = vbKeyDelete Then
     Call SpreadCellDataClear(Va_Details, Va_Details.ActiveRow, Va_Details.ActiveCol)
     Va_Details.Row = Va_Details.ActiveRow
     Va_Details.Col = Va_Details.ActiveCol
        If Va_Details.Lock = False And Va_Details.Col <> 1 Then
           Va_Details.Row = Va_Details.ActiveRow
           Va_Details.Col = 1
              Va_Details.Value = 1
        End If
  
  ElseIf Va_Details.ActiveCol = 12 And KeyCode = vbKeyF2 Then
     Search.Query = "select c_shiftcode Code, n_shifthrs ShiftHrs, n_starthrs StartTime, " & _
                    "n_endhrs EndTime from pr_clock_shift "
     Search.CheckFields = "Code"
     Search.ReturnField = "Code"
     SerVar = Search.Search(, , CON)
     If Len(Search.col1) <> 0 Then
        Va_Details.Row = Va_Details.ActiveRow
        Va_Details.Col = 12
           Va_Details.Text = Search.col1
        Va_Details.Col = 1
           Va_Details.Value = 1
     End If
  End If
End Sub

Private Sub Va_Details_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
  Dim i As Integer
  Dim rsChk As New ADODB.Recordset
  Dim tmpStr As String, tmpLeaveCodes As String
  Dim tmpOT As Double

    If Col = 20 Or Col = 21 Then
       Va_Details.Row = Row
       Va_Details.Col = 20
          tmpStr = UCase(Va_Details.Text)
          Va_Details.Text = tmpStr
          
          If Trim(tmpStr) <> "" Then
             tmpLeaveCodes = "'" & vLeaveCodes & "'"
             tmpStr = "'" & tmpStr & "'"
             If InStr(1, tmpLeaveCodes, tmpStr, vbTextCompare) = 0 Then
                Va_Details.Text = ""
                MsgBox "Not a valid code.", vbInformation, "Information"
             Else
                Va_Details.Col = 20
                   tmpStr = Trim(UCase(Va_Details.Text))
                   If tmpStr = "CL" Or tmpStr = "SL" Or tmpStr = "VL" Then
                      Call DisplayLeaveBalance(Row)
                   End If
             End If
          End If
    End If
    
    If Col = 13 Or Col = 15 Or Col = 19 Then
       Va_Details.Row = Row
       Va_Details.Col = 13  'in
          Va_Details.Text = Spread_NumFormat(ChkValidTime(Val(Va_Details.Text)), True)
       Va_Details.Col = 15  'out
          Va_Details.Text = Spread_NumFormat(ChkValidTime(Val(Va_Details.Text)), True)
       Va_Details.Col = 19  'perm
          Va_Details.Text = Spread_NumFormat(ChkValidTime(Val(Va_Details.Text)), True)
    End If
    
    If Col = 12 Or Col = 13 Or Col = 15 Or Col = 17 Or Col = 19 Or Col = 20 Or Col = 21 Then
       Va_Details.Row = Row
       Va_Details.Col = 13   'in time
          If Val(Va_Details.Text) = 0 Then
             Va_Details.Col = 14
                Va_Details.Text = ""
          End If
       Va_Details.Col = 15   'out time
          If Val(Va_Details.Text) = 0 Then
             Va_Details.Col = 16
                Va_Details.Text = ""
          End If
       Va_Details.Col = 17
         ' If Val(Va_Details.Text) = 0 Then
             Va_Details.Col = 1
                If Va_Details.Value = 1 Then
                   Call Calculate_WorkHrs(Row)
                   
                   Va_Details.Row = Row
                   Va_Details.Col = 17
                      If Val(Va_Details.Text) > 0 Then
                         Va_Details.Col = 20
                            If Va_Details.Text = "A" Or Va_Details.Text = "WO" Then
                               Va_Details.Text = "P"
                            End If
                      Else
                         Va_Details.Col = 17
                            Va_Details.Text = ""
                         Va_Details.Col = 18
                            Va_Details.Text = ""
                      End If
                End If
         ' End If
    End If
    
    If Col = 21 Then
       Va_Details.Row = Row
       Va_Details.Col = 21
          If Val(Va_Details.Text) > 1 Then
             Va_Details.Text = 1
             MsgBox "Should not be more than 1 day", vbInformation, "Information"
          End If
    End If
    
End Sub

Private Sub Re_Calculate_WorkHrs()
    Dim i As Integer
    
    For i = 1 To Va_Details.DataRowCnt
        Va_Details.Row = i
        Va_Details.Col = 1
           If Va_Details.Value = 1 Then
              Call Calculate_WorkHrs(i)
           End If
    Next i
End Sub

Private Sub Calculate_WorkHrs(ByVal vRow As Long)
  Dim rsChk As New ADODB.Recordset
  Dim i As Integer, vCanteen As Integer, vWeekDay As Integer
  Dim vShift As String, vEmpNo As String, vDesig As String, vStaffType As String, vPresAbs As String, vDayWork As String, vDept As String
  Dim vArrTime, vDepTime, vWorkHrs, vPermHrs, vLateHrs, vEarlyHrs As Double
  Dim vShStartHrs, vShEndHrs, vShWorkHrs, vShLateHrs, vShBreakHrs As Double
  Dim vShBr1, vShBrMin1, vShBr2, vShBrMin2, vShBr3, vShBrMin3, vLeaveHrs As Double
  Dim vActArrTime, vActDepTime, vActWorkhrs As Double, vOThrs As Double, vAddBrTime As Double, tmpVal As Double, vPresent As Double
  Dim vOtIn As Double, vOtOut As Double, vOt15 As Double, vOt20 As Double, vOt30 As Double
  Dim vDate As Date
  Dim vOtInApp As Boolean, vOtOutApp As Boolean
  
    vShift = "": vEmpNo = "": vDesig = "": vStaffType = ""
    vArrTime = 0: vDepTime = 0: vWorkHrs = 0: vPermHrs = 0: vLateHrs = 0: vEarlyHrs = 0: vCanteen = 0
    vShStartHrs = 0: vShEndHrs = 0: vShWorkHrs = 0: vShLateHrs = 0: vShBreakHrs = 0
    vShBr1 = 0: vShBrMin1 = 0: vShBr2 = 0: vShBrMin2 = 0: vShBr3 = 0: vShBrMin3 = 0: vLeaveHrs = 0
    vActArrTime = 0: vActDepTime = 0: vActWorkhrs = 0: vOThrs = 0: vOtIn = 0: vOtOut = 0: vOt15 = 0: vOt20 = 0: vOt30 = 0
    vAddBrTime = 0: vWeekDay = 0: tmpVal = 0: vPresent = 0
    vOtInApp = False: vOtOutApp = False
    
    Va_Details.Row = vRow
    Va_Details.Col = 2
       vEmpNo = Trim(Va_Details.Text)
    Va_Details.Col = 10
       If Trim(Va_Details.Text) <> "" Then
          vDate = Trim(Va_Details.Text)
          vWeekDay = Weekday(vDate, vbMonday)
       End If
       
       
    Set rsChk = Nothing
    g_Sql = "select c_desig, c_stafftype, c_daywork, c_dept from pr_emp_mst where c_rec_sta = 'A' and c_empno = '" & Trim(vEmpNo) & "'"
    rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If rsChk.RecordCount > 0 Then
       vDesig = Is_Null(rsChk("c_desig").Value, False)
       vStaffType = Is_Null(rsChk("c_stafftype").Value, False)
       vDayWork = Left(Is_Null(rsChk("c_daywork").Value, False), 1)
       vDept = Is_Null(rsChk("c_dept").Value, False)
    End If
       
    Va_Details.Row = vRow
    Va_Details.Col = 12
       vShift = Trim(Va_Details.Text)
    Va_Details.Col = 13
       vArrTime = Val(Va_Details.Text)
    Va_Details.Col = 15
       vDepTime = Val(Va_Details.Text)
    Va_Details.Col = 19
       vPermHrs = Val(Va_Details.Text)
    Va_Details.Col = 20
       vPresAbs = Trim(Va_Details.Text)
    Va_Details.Col = 21
       vPresent = Val(Va_Details.Text)
       
    If Trim(vShift) <> "" And vArrTime > 0 And vDepTime > 0 Then
       Set rsChk = Nothing
       g_Sql = "Select starthrs, endhrs, shifthrs, breakhrs, latehrs, " & _
               "break1, mins1, break2, mins2, break3, mins3 " & _
               "from pr_clock_shift where c_shiftcode = '" & Trim(vShift) & "' "
       rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
       If rsChk.RecordCount > 0 Then
          vShStartHrs = rsChk("starthrs").Value
          vShEndHrs = rsChk("endhrs").Value
          vShWorkHrs = rsChk("shifthrs").Value
          vShLateHrs = rsChk("latehrs").Value
          vShBreakHrs = rsChk("breakhrs").Value
            
          vShBr1 = rsChk("break1").Value
          vShBrMin1 = rsChk("mins1").Value
          vShBr2 = rsChk("break2").Value
          vShBrMin2 = rsChk("mins2").Value
          vShBr3 = rsChk("break3").Value
          vShBrMin3 = rsChk("mins3").Value
        
          ' Arr Time
          vArrTime = TimeToMins(vArrTime)
          If vArrTime = 0 Then
             vActArrTime = 0
          ElseIf vArrTime < (vShStartHrs - vShLateHrs) Then
             vActArrTime = vArrTime
          ElseIf vArrTime >= (vShStartHrs - vShLateHrs) And vArrTime <= (vShStartHrs + vShLateHrs) Then
             vActArrTime = vShStartHrs
          ElseIf vArrTime > (vShStartHrs + vShLateHrs) And vArrTime <= (vShBr1 + vShBrMin1 + vShLateHrs) Then
             If vArrTime < (vShBr1 - vShLateHrs) Then
                vActArrTime = vArrTime
             ElseIf vArrTime >= (vShBr1 - vShLateHrs) Then
                vActArrTime = vShBr1 + vShBrMin1: vShBreakHrs = vShBreakHrs - vShBrMin1
             End If
          ElseIf vArrTime > (vShBr1 + vShBrMin1 + vShLateHrs) And vArrTime <= (vShBr2 + vShBrMin2 + vShLateHrs) Then
             If vArrTime < (vShBr2 - vShLateHrs) Then
                vActArrTime = vArrTime: vShBreakHrs = vShBreakHrs - vShBrMin1
             ElseIf vArrTime >= (vShBr2 - vShLateHrs) Then
                vActArrTime = vShBr2 + vShBrMin2: vShBreakHrs = vShBreakHrs - (vShBrMin1 + vShBrMin2)
             End If
          ElseIf vArrTime > (vShBr2 + vShBrMin2 + vShLateHrs) And vArrTime <= (vShBr3 + vShBrMin3 + vShLateHrs) Then
             If vArrTime < (vShBr3 - vShLateHrs) Then
                vActArrTime = vArrTime: vShBreakHrs = vShBreakHrs - (vShBrMin1 + vShBrMin2)
             ElseIf vArrTime >= (vShBr3 - vShLateHrs) Then
                vActArrTime = vShBr3 + vShBrMin3: vShBreakHrs = 0
             End If
          Else
             vActArrTime = vArrTime: vShBreakHrs = 0
          End If
          
          ' OT In time
          If vActArrTime > 0 And vActArrTime < vShStartHrs Then
             vOtIn = vShStartHrs - vActArrTime
          End If
          
 
          ' Dep Time
          vDepTime = TimeToMins(vDepTime)
          If vDepTime <= (vShStartHrs + vShLateHrs) Then
             vActDepTime = 0: vShBreakHrs = 0
          ElseIf vDepTime > (vShStartHrs + vShLateHrs) And vDepTime <= (vShBr1 + vShBrMin1 + vShLateHrs) Then
             If vDepTime < vShBr1 Then
                vActDepTime = vDepTime: vShBreakHrs = 0
             ElseIf vDepTime >= vShBr1 Then
                vActDepTime = vShBr1: vShBreakHrs = 0
             End If
          ElseIf vDepTime > (vShBr1 + vShBrMin1 + vShLateHrs) And vDepTime <= (vShBr2 + vShBrMin2 + vShLateHrs) Then
             If vDepTime < vShBr2 Then
                vActDepTime = vDepTime: vShBreakHrs = vShBrMin1
             ElseIf vDepTime >= vShBr2 Then
                vActDepTime = vShBr2: vShBreakHrs = vShBrMin1
             End If
          ElseIf vDepTime > (vShBr2 + vShBrMin2 + vShLateHrs) And vDepTime <= (vShBr3 + vShBrMin3 + vShLateHrs) Then
             If vDepTime < vShBr3 Then
                vActDepTime = vDepTime: vShBreakHrs = vShBreakHrs - vShBrMin3
             ElseIf vDepTime >= vShBr3 Then
                vActDepTime = vShBr3: vShBreakHrs = vShBreakHrs - vShBrMin3
             End If
          ElseIf vDepTime > (vShBr3 + vShBrMin3 + vShLateHrs) And vDepTime <= vShEndHrs + vShLateHrs Then
             If vDepTime < (vShEndHrs - vShLateHrs) Then
                vActDepTime = vDepTime
             Else
                vActDepTime = vShEndHrs
             End If
          Else
             vActDepTime = vDepTime
          End If
          
          ' OT Out time
          If vActDepTime > 0 And vActDepTime > vShEndHrs Then
             vOtOut = vActDepTime - vShEndHrs
          End If
          
          
          ' Late Hrs
          If vActArrTime > vShStartHrs And vActArrTime < vShEndHrs Then
             If vActArrTime < vShBr1 Then
                vLateHrs = vActArrTime - vShStartHrs
             ElseIf vActArrTime >= vShBr1 And vActArrTime < vShBr2 Then
                vLateHrs = (vActArrTime - vShBrMin1) - vShStartHrs
             ElseIf vActArrTime >= vShBr2 And vActArrTime < vShBr3 Then
                vLateHrs = (vActArrTime - (vShBrMin1 + vShBrMin2)) - vShStartHrs
             Else
                vLateHrs = (vActArrTime - (vShBrMin1 + vShBrMin2 + vShBrMin3)) - vShStartHrs
             End If
          End If
          If vWeekDay = 6 Or vWeekDay = 7 Or vPresAbs = "PH" Then
             vLateHrs = 0
          End If
          
          
          ' Early Hrs
          If vActDepTime > vShStartHrs And vActDepTime < vShEndHrs Then
             If vActDepTime <= vShBr1 Then
                vEarlyHrs = (vShEndHrs - vActDepTime) - (vShBrMin1 + vShBrMin2 + vShBrMin3)
             ElseIf vActDepTime > vShBr1 And vActDepTime <= vShBr2 Then
                vEarlyHrs = (vShEndHrs - vActDepTime) - (vShBrMin2 + vShBrMin3)
             ElseIf vActDepTime > vShBr2 And vActDepTime <= vShBr3 Then
                vEarlyHrs = (vShEndHrs - vActDepTime) - vShBrMin3
             Else
                vEarlyHrs = vShEndHrs - vActDepTime
             End If
          End If
          If vWeekDay = 6 Or vWeekDay = 7 Or vPresAbs = "PH" Then
             vEarlyHrs = 0
          End If
          
          If vArrTime = 0 Then vLateHrs = 0
          If vDepTime = 0 Then vEarlyHrs = 0
          If vActArrTime >= vActDepTime Then vLateHrs = 0
          If vActDepTime <= vActArrTime Then vEarlyHrs = 0
          
                    
          ' Work hrs
          vPermHrs = TimeToMins(vPermHrs)
          'If vWeekDay = 6 Or vWeekDay = 7 Or vPresAbs = "PH" Then
          '   vWorkHrs = (vActDepTime - vActArrTime)
          'Else
             vWorkHrs = (vActDepTime - vActArrTime) - (vShBreakHrs + vPermHrs)
          'End If
          If vWorkHrs < 0 Then
             vWorkHrs = 0
          ElseIf vWorkHrs > 1800 Then '30 hrs
             vWorkHrs = 0
          End If
          
          If vPresAbs = "P" Or vPresAbs = "WO" Or vPresAbs = "A" Or vPresAbs = "PH" Then
          Else
             If vWorkHrs > 0 Then
                If vLateHrs > 0 Then vLateHrs = 0
                If vEarlyHrs > 0 Then vEarlyHrs = 0
                vPresent = 0.5
             End If
          End If
          
          
          ' OT break-up condition
          Va_Details.Row = vRow
          Va_Details.Col = 23
             If Va_Details.FontBold = True Then
                vOtInApp = True
             End If
          Va_Details.Col = 24
             If Va_Details.FontBold = True Then
                vOtOutApp = True
             End If
             
          ' OT break-up
          If vWeekDay = 7 Or vPresAbs = "PH" Then   'sun or ph
             vOtIn = vWorkHrs
             vOtOut = 0
             If vOtIn > 0 And vOtInApp Then
                vOt20 = FloorTo15Mins(vOtIn)
                If vOt20 > 480 Then
                   vOt30 = vOt20 - 480
                   vOt20 = 480
                End If
             End If
             
          ElseIf vWeekDay = 6 Then   'sat
            vOtIn = vWorkHrs
            vOtOut = 0
             If vOtIn > 0 And vOtInApp Then
                If vActArrTime >= 360 Then  '6.00
                   vOt15 = FloorTo15Mins(vOtIn)
                Else
                   vOt20 = FloorTo15Mins(360 - vActArrTime)
                   vOt15 = FloorTo15Mins(vOtIn) - vOt20
                End If
                
                If vActDepTime <= 1380 Then  '23.00
                   vOt15 = vOt15 + FloorTo15Mins(vOtOut)
                Else
                   If vOt20 > 0 Then
                      tmpVal = vOt20
                   End If
                   vOt20 = vOt20 + FloorTo15Mins(vActDepTime - 1380)
                   vOt15 = vOt15 + (FloorTo15Mins(vOtOut) - (vOt20 - tmpVal))
                End If
             End If
            
          Else
             If vOtIn > 0 And vOtInApp Then
                If vActArrTime >= 360 Then  '6.00
                   vOt15 = FloorTo15Mins(vOtIn)
                Else
                   vOt20 = FloorTo15Mins(360 - vActArrTime)
                   vOt15 = FloorTo15Mins(vOtIn) - vOt20
                End If
             End If
            
             If vOtOut > 0 And vOtOutApp Then
                If vActDepTime <= 1380 Then  '23.00
                   vOt15 = vOt15 + FloorTo15Mins(vOtOut)
                Else
                   If vOt20 > 0 Then
                      tmpVal = vOt20
                   End If
                   vOt20 = vOt20 + FloorTo15Mins(vActDepTime - 1380)
                   vOt15 = vOt15 + (FloorTo15Mins(vOtOut) - (vOt20 - tmpVal))
                End If
             End If
          End If
          
          ' OT hrs
          If vWeekDay = 6 Or vWeekDay = 7 Or vPresAbs = "PH" Then
             vOThrs = vWorkHrs
          Else
             vOThrs = vOtIn + vOtOut
          End If

          If vWorkHrs = 0 Then vOThrs = 0
          If vOThrs < 0 Then vOThrs = 0
          
          If vOThrs = 0 Then
             vOtIn = 0: vOtOut = 0: vOt15 = 0: vOt20 = 0: vOt30 = 0
          End If
          
       End If
       
       Va_Details.Row = vRow
       Va_Details.Col = 14
          If vLateHrs > 0 Then
             Va_Details.Text = MinsToTime(vLateHrs)
          Else
             Va_Details.Text = ""
          End If
       Va_Details.Col = 16
          If vEarlyHrs > 0 Then
             Va_Details.Text = MinsToTime(vEarlyHrs)
          Else
             Va_Details.Text = ""
          End If
       Va_Details.Col = 17
          If vWorkHrs > 0 Then
             Va_Details.Text = MinsToTime(vWorkHrs)
          Else
             Va_Details.Text = ""
          End If
       Va_Details.Col = 18
          If vOThrs > 0 Then
             Va_Details.Text = MinsToTime(vOThrs)
          Else
             Va_Details.Text = ""
          End If
       Va_Details.Col = 21
         If vWorkHrs > 0 Then
            If vPresAbs = "P" Or vPresAbs = "WO" Or vPresAbs = "A" Or vPresAbs = "PH" Then
            Else
               Va_Details.Text = vPresent
            End If
         End If
       Va_Details.Col = 23
          If vOtIn > 0 Then
             Va_Details.Text = MinsToTime(vOtIn)
             Va_Details.Text = DecimalToString(Va_Details.Text)
             If Va_Details.FontBold = True Then
                Va_Details.Text = OTSignAdd(Trim(Va_Details.Text))
             End If
          Else
             Va_Details.Text = ""
             Va_Details.FontBold = False
          End If
       Va_Details.Col = 24
          If vOtOut > 0 Then
             Va_Details.Text = MinsToTime(vOtOut)
             Va_Details.Text = DecimalToString(Va_Details.Text)
             If Va_Details.FontBold = True Then
                Va_Details.Text = OTSignAdd(Trim(Va_Details.Text))
             End If
          Else
             Va_Details.Text = ""
             Va_Details.FontBold = False
          End If
       Va_Details.Col = 25
          If vOt15 > 0 Then
             Va_Details.Text = MinsToTime(vOt15)
          Else
             Va_Details.Text = ""
          End If
       Va_Details.Col = 26
          If vOt20 > 0 Then
             Va_Details.Text = MinsToTime(vOt20)
          Else
             Va_Details.Text = ""
          End If
       Va_Details.Col = 27
          If vOt30 > 0 Then
             Va_Details.Text = MinsToTime(vOt30)
          Else
             Va_Details.Text = ""
          End If
    Else
        If vArrTime = 0 Or vDepTime = 0 Then
           Va_Details.Row = vRow
           Va_Details.Col = 17
              Va_Details.Text = ""
           Va_Details.Col = 18
              Va_Details.Text = ""
           Va_Details.Col = 19
              Va_Details.Text = ""
           Va_Details.Col = 23
              Va_Details.Text = ""
           Va_Details.Col = 24
              Va_Details.Text = ""
           Va_Details.Col = 25
              Va_Details.Text = ""
           Va_Details.Col = 26
              Va_Details.Text = ""
           Va_Details.Col = 27
              Va_Details.Text = ""
        End If
        If vPresAbs = "P" Or vPresAbs = "WO" Or vPresAbs = "A" Or vPresAbs = "PH" Then
           If vArrTime = 0 And vDepTime = 0 Then
              Va_Details.Col = 42
                 Va_Details.Text = ""
           End If
        End If
    End If
End Sub

Private Sub Combo_Load()
    Call LoadComboCompany(Me)
    Call LoadComboBranch(Me)
    Call LoadComboDept(Me)
    Call LoadComboDesig(Me)
    Call LoadComboEmpType(Me)
    Call LoadComboShift(Me)
End Sub

Private Sub Chk_Hide_Click(Index As Integer)
   Call Spread_ColHide
End Sub

Private Sub Btn_GridDefault_Click()
On Error GoTo ErrSave
  Dim vGridView As String
  Dim i As Integer
  
    For i = 1 To Va_Details.MaxCols
        Va_Details.Row = -1
        Va_Details.Col = i
           If Va_Details.ColHidden = True Then
              vGridView = vGridView & Trim(Str(i)) & ","
           End If
    Next i
    
    CON.BeginTrans
    
    Set rs = Nothing
    g_Sql = "select * from pr_user_grid_set where c_user_id = '" & g_UserName & "' and c_screen_id = '" & g_AttnScr & "'"
    rs.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
    If rs.RecordCount = 0 Then
       rs.AddNew
       rs("c_user_id").Value = g_UserName
       rs("c_screen_id").Value = g_AttnScr
       rs("c_gridview").Value = vGridView
       rs.Update
    Else
       rs("c_gridview").Value = vGridView
       rs.Update
    End If
    
    CON.CommitTrans
    
    MsgBox "Grid column display is set successfully. The grid will be displaied as you set for next time when you open this screen.", vbInformation, "Information"
    
  Exit Sub
     
ErrSave:
     CON.RollbackTrans
     Screen.MousePointer = vbDefault
     MsgBox "Error while Saving - " + Err.Description, vbCritical, "Critical"
End Sub

Private Sub Spread_Hide_Check()
  Dim rsChk As New ADODB.Recordset
  Dim vGridView As String
    
    Set rsChk = Nothing
    g_Sql = "select c_gridview from pr_user_grid_set where c_user_id = '" & g_UserName & "' and c_screen_id = '" & g_AttnScr & "'"
    rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If rsChk.RecordCount = 0 Then
       Call Spread_ColHide
    Else
       vGridView = Is_Null(rsChk("c_gridview").Value, False)
       If Trim(vGridView) = "" Then
          Call Spread_ColHide
       Else
          Call Spread_ColUser_Hide(vGridView)
       End If
    End If
    
End Sub

Private Sub Spread_ColUser_Hide(ByVal vGridView As String)
  Dim i As Integer
    
    For i = 1 To Va_Details.MaxCols
        Va_Details.Row = -1
        Va_Details.Col = i
           Va_Details.ColHidden = IsColHide(vGridView, i)
    Next i
  
End Sub

Private Function IsColHide(ByVal vGridView As String, ByVal vCol As Integer) As Boolean
  Dim i As Integer
  Dim vColHide As Variant
  
    vColHide = Split(vGridView, ",")
    If UBound(vColHide) < 0 Then Exit Function
       
    For i = 0 To UBound(vColHide)
        If vColHide(i) = vCol Then
           IsColHide = True
           Exit For
        End If
    Next i
End Function

Private Sub Spread_ColHide()
  Dim i As Integer, j As Integer
  
    Va_Details.Row = -1
    Va_Details.Col = 9
       Va_Details.ColHidden = True
    Va_Details.Col = 22
       Va_Details.ColHidden = True
    Va_Details.Col = 28
       Va_Details.ColHidden = True
    Va_Details.Col = 41
       Va_Details.ColHidden = True
    Va_Details.Col = 45
       Va_Details.ColHidden = True
    Va_Details.Col = 46
       Va_Details.ColHidden = True
    Va_Details.Col = 47
       Va_Details.ColHidden = True
    Va_Details.Col = 48
       Va_Details.ColHidden = True
       
    For j = 1 To 5
        Va_Details.Row = -1
        Va_Details.Col = 3 + j
           If Chk_Hide(0).Value = 1 Then
              Va_Details.ColHidden = True
           Else
              Va_Details.ColHidden = False
           End If
    Next j
    
    For j = 1 To 12
        Va_Details.Row = -1
        Va_Details.Col = 28 + j
           If Chk_Hide(1).Value = 1 Then
              Va_Details.ColHidden = True
           Else
              Va_Details.ColHidden = False
           End If
    Next j
       
End Sub

Private Sub Btn_Close_Click()
   Va_Details.SetFocus
End Sub

Private Sub Cmd_Filter_Click()
  Dim rsChk As New ADODB.Recordset
  Dim i As Integer
  Dim tmpStr As String
              
   Set rsChk = Nothing
   g_Sql = "select * from Pr_Clock_Filter where desp is not null "
   If g_Admin Then
      g_Sql = g_Sql & " order by desp "
   Else
      g_Sql = g_Sql & " and flag = 'U' order by desp "
   End If
   rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
   For i = 1 To rsChk.RecordCount
       tmpStr = tmpStr & rsChk("desp").Value & Space(100) & "~" & rsChk("fieldname").Value & "~" & rsChk("type").Value & Chr$(9)
       rsChk.MoveNext
   Next i
    
   Va_Filter.Row = -1
   Va_Filter.Col = 1
      Va_Filter.TypeComboBoxList = tmpStr
      
   tmpStr = "is equal to" & Space(100) & "1" & Chr$(9) & _
            "is less than " & Space(100) & "2" & Chr$(9) & _
            "is greater than " & Space(100) & "3" & Chr$(9) & _
            "is not equal to " & Space(100) & "8"

   Va_Filter.Row = -1
   Va_Filter.Col = 2
      Va_Filter.TypeComboBoxList = tmpStr

   Frm_Filter.Left = 5000
   Frm_Filter.Top = 2800
   Frm_Filter.Visible = True
   Frm_Filter.ZOrder 0
   
   Va_Filter.Row = -1
   Va_Filter.Col = -1
      Va_Filter.Lock = False
   
End Sub

Private Sub Btn_Cancel_Click()
   strFilter = ""
   Frm_Filter.Visible = False
End Sub

Private Sub Btn_ClearFilter_Click()
   strFilter = ""
   Clear_Spread Va_Filter
End Sub

Private Sub Btn_Ok_Click()
  Dim i As Integer
  Dim tmpStr As String
  
    For i = 1 To Va_Filter.DataRowCnt
        Va_Filter.Row = i
        Va_Filter.Col = 4
           If Trim(Va_Filter.Text) <> "" Then
              tmpStr = Trim(Va_Filter.Text)
           End If
        Va_Filter.Col = 3
           If Trim(Va_Filter.Text) <> "" Then
              strFilter = Make_Filter_Option(strFilter, Trim(Va_Filter.Text), i, tmpStr)
           End If
    Next i
   
   Call Btn_Display_Click
   
   strFilter = ""
   Frm_Filter.Visible = False
End Sub

Private Function Make_Filter_Option(vTotStr As String, vOptStr As String, vRow As Integer, Optional vOptStr2 As String)
   Dim tmpStr As String, tmpOption As String
   Dim tmpField
    
    Va_Filter.Row = vRow
    Va_Filter.Col = 1
       Va_Filter.TypeComboBoxIndex = Va_Filter.TypeComboBoxCurSel
       If Trim(Va_Filter.TypeComboBoxString) <> "" Then
          tmpField = Split(Trim(Va_Filter.TypeComboBoxString), "~")
       Else
          Make_Filter_Option = vTotStr
          Exit Function
       End If
    Va_Filter.Col = 2
       Va_Filter.TypeComboBoxIndex = Va_Filter.TypeComboBoxCurSel
       tmpOption = Right(Trim(Va_Filter.TypeComboBoxString), 1)
       If Trim(tmpOption) = "" Then
          Make_Filter_Option = vTotStr
          Exit Function
       End If
       
    If Val(tmpOption) = 7 Then
       If vOptStr = "" Or vOptStr2 = "" Then
          Make_Filter_Option = vTotStr
          Exit Function
       End If
    End If
       
    If Val(tmpOption) = 7 Then
       tmpStr = " and " & tmpField(1) & " between '"
    Else
       tmpStr = " and " & tmpField(1)
    End If
    
    If Val(tmpOption) = 7 Then
       ' to avoid
    ElseIf Val(tmpOption) >= 4 And Val(tmpOption) <= 6 Then
       If tmpOption = "4" Then
          tmpStr = tmpStr & " like ('%"
       ElseIf tmpOption = "5" Then
          tmpStr = tmpStr & " like ('"
       ElseIf tmpOption = "6" Then
          tmpStr = tmpStr & " like ('%"
       End If
    Else
       If tmpOption = "1" Then
          tmpStr = tmpStr & " = "
       ElseIf tmpOption = "2" Then
          tmpStr = tmpStr & " < "
       ElseIf tmpOption = "3" Then
          tmpStr = tmpStr & " > "
       ElseIf tmpOption = "8" Then
          tmpStr = tmpStr & " <> "
       End If
        
       If tmpField(2) <> "N" Then
          If Trim(vOptStr) = "''" Then
             tmpStr = Replace(tmpStr, "=", "") & " is null"
          Else
             tmpStr = tmpStr & "'"
          End If
       End If
    End If
    
    If tmpField(2) = "D" Then
       vOptStr = Format(vOptStr, "yyyy-mm-dd")
       vOptStr2 = Format(vOptStr2, "yyyy-mm-dd")
    End If
    
    If Val(tmpOption) = 7 Then
       tmpStr = tmpStr & vOptStr & "' and '" & vOptStr2 & "'"
    Else
       tmpStr = tmpStr & vOptStr
    End If
    
    If Val(tmpOption) = 7 Then
       ' to avoid
    ElseIf Val(tmpOption) >= 4 And Val(tmpOption) <= 6 Then
       If tmpOption = "4" Then
          tmpStr = tmpStr & "%')"
       ElseIf tmpOption = "5" Then
          tmpStr = tmpStr & "%')"
       ElseIf tmpOption = "6" Then
          tmpStr = tmpStr & "')"
       End If
    Else
       If tmpField(2) <> "N" Then
          If Trim(vOptStr) = "''" Then
             tmpStr = Replace(tmpStr, "''", "")
          Else
             tmpStr = tmpStr & "'"
          End If
       End If
    End If
    
    Make_Filter_Option = vTotStr & tmpStr
       
End Function

Private Sub Va_Filter_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim Search As New Search.MyClass, SerVar, SerArray
  Dim tmpFilter As String
    
   If KeyCode = vbKeyDelete Then
      Call SpreadCellDataClear(Va_Filter, Va_Filter.ActiveRow, Va_Filter.ActiveCol)
   End If
End Sub

Private Sub Va_Filter_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
  Dim tmpArr
  
    If Col = 1 Then
       Va_Filter.Row = Row
       Va_Filter.Col = Col
          Va_Filter.TypeComboBoxIndex = Va_Filter.TypeComboBoxCurSel
          If Trim(Va_Filter.TypeComboBoxString) <> "" Then
             tmpArr = Split(Trim(Va_Filter.TypeComboBoxString), "~")
             If tmpArr(1) = "n_otin_str" Or tmpArr(1) = "n_otout_str" Then
                Va_Filter.Col = 3
                   Va_Filter.Text = "$"
             End If
          End If
    End If
End Sub

Private Sub Pr_Clock_Emp_RepFilter()

    If IsDate(Dtp_FromDate.Text) And IsDate(Dtp_ToDate.Text) Then
       If CDate(Dtp_FromDate.Text) = CDate(Dtp_ToDate.Text) Then
          RepHeadDate = "Date : " & Format(Dtp_FromDate.Text, "dd/mm/yyyy")
       Else
          RepHeadDate = "Date From " & Format(Dtp_FromDate.Text, "dd/mm/yyyy") & "  To  " & Format(Dtp_ToDate.Text, "dd/mm/yyyy")
       End If
    End If
    
    vF1 = "{PR_CLOCK_EMP.D_DATE}>=DATETIME(" & Year(Dtp_FromDate.Text) & "," & Format(Month(Dtp_FromDate.Text), "00") & "," & Format(Day(Dtp_FromDate.Text), "00") & ")"
    vF1 = Trim(vF1) & " AND {PR_CLOCK_EMP.D_DATE}<=DATETIME(" & Year(Dtp_ToDate.Text) & "," & Format(Month(Dtp_ToDate.Text), "00") & "," & Format(Day(Dtp_ToDate.Text), "00") & ")"
    
    If Trim(Cmb_Company) <> "" Then
       vF2 = "{V_PR_EMP_MST.C_COMPANY}='" & Trim(Right(Trim(Cmb_Company), 7)) & "'"
    End If
    
    If Trim(Cmb_Branch) <> "" Then
       vF3 = "{V_PR_EMP_MST.C_BRANCH}='" & Trim(Cmb_Branch) & "'"
    End If
    
    If Trim(Cmb_Dept) <> "" Then
       vF4 = "{V_PR_EMP_MST.C_DEPT}='" & Trim(Cmb_Dept) & "'"
    End If
      
    If Trim(Cmb_Desig) <> "" Then
       vF5 = "{V_PR_EMP_MST.C_DESIG}='" & Trim(Cmb_Desig) & "'"
    End If
      
      
    SelFor = ReportFilterOption(vF1, vF2, vF3, vF4, vF5)
      
    vF1 = "": vF2 = "": vF3 = "": vF4 = "": vF5 = ""
    
    If Trim(Cmb_Shift) <> "" Then
       vF1 = "{V_PR_EMP_MST.C_SHIFTCODE}='" & Trim(Right(Trim(Cmb_Shift), 7)) & "'"
    End If
   
    If Trim(Cmb_EmpType) <> "" Then
       vF2 = "{V_PR_EMP_MST.C_EMPTYPE}='" & Trim(Cmb_EmpType) & "'"
    End If
    
    If Right(Trim(Txtc_EmployeeName), 6) <> "" Then
       vF3 = "{V_PR_EMP_MST.C_EMPNO}='" & Trim(Right(Trim(Txtc_EmployeeName), 7)) & "'"
    End If
    
    SelFor = ReportFilterOption(SelFor, vF1, vF2, vF3)
    
End Sub

Private Sub Pr_Clock_FILO_RepFilter()
    
    vF1 = "": vF2 = "": vF3 = "": vF4 = "": vF5 = ""

    If IsDate(Dtp_FromDate.Text) And IsDate(Dtp_ToDate.Text) Then
       RepHeadDate = "FOR THE MONTH OF " & UCase(MonthName(Month(Dtp_FromDate.Text))) & "  -  " & Trim(Str(Year(Dtp_FromDate.Text)))
    End If
    
    vF1 = "{V_PR_CLOCK_DAY.N_MONTH} = " & Month(Dtp_FromDate.Text) & " AND {V_PR_CLOCK_DAY.N_YEAR} = " & Year(Dtp_FromDate.Text)
    
    If Trim(Cmb_Company) <> "" Then
       vF2 = "{V_PR_EMP_MST.C_COMPANY}='" & Trim(Right(Trim(Cmb_Company), 7)) & "'"
    End If
    
    If Trim(Cmb_Branch) <> "" Then
       vF3 = "{V_PR_EMP_MST.C_BRANCH}='" & Trim(Cmb_Branch) & "'"
    End If
    
    If Trim(Cmb_Dept) <> "" Then
       vF4 = "{V_PR_EMP_MST.C_DEPT}='" & Trim(Cmb_Dept) & "'"
    End If
      
    If Trim(Cmb_Desig) <> "" Then
       vF5 = "{V_PR_EMP_MST.C_DESIG}='" & Trim(Cmb_Desig) & "'"
    End If
      
      
    SelFor = ReportFilterOption(vF1, vF2, vF3, vF4, vF5)
      
    vF1 = "": vF2 = "": vF3 = "": vF4 = "": vF5 = ""
    
    If Trim(Cmb_Shift) <> "" Then
       vF1 = "{V_PR_EMP_MST.C_SHIFTCODE}='" & Trim(Right(Trim(Cmb_Shift), 7)) & "'"
    End If
    
    If Trim(Cmb_EmpType) <> "" Then
       vF2 = "{V_PR_EMP_MST.C_EMPTYPE}='" & Trim(Cmb_EmpType) & "'"
    End If
    
    If Right(Trim(Txtc_EmployeeName), 6) <> "" Then
       vF3 = "{V_PR_EMP_MST.C_EMPNO}='" & Trim(Right(Trim(Txtc_EmployeeName), 6)) & "'"
    End If
    
    SelFor = ReportFilterOption(SelFor, vF1, vF2, vF3)
    
End Sub

Private Sub Pr_Clock_Status_RepFilter()

    vF1 = "": vF2 = "": vF3 = "": vF4 = "": vF5 = ""

    If IsDate(Dtp_FromDate.Text) And IsDate(Dtp_ToDate.Text) Then
       If CDate(Dtp_FromDate.Text) = CDate(Dtp_ToDate.Text) Then
          RepHeadDate = "Date : " & Format(Dtp_FromDate.Text, "dd/mm/yyyy")
       Else
          RepHeadDate = "Date From " & Format(Dtp_FromDate.Text, "dd/mm/yyyy") & "  To  " & Format(Dtp_ToDate.Text, "dd/mm/yyyy")
       End If
    End If
    
    vF1 = "{V_PR_CLOCK_STATUS.D_DATE}>=DATETIME(" & Year(Dtp_FromDate.Text) & "," & Format(Month(Dtp_FromDate.Text), "00") & "," & Format(Day(Dtp_FromDate.Text), "00") & ")"
    vF1 = Trim(vF1) & " AND {V_PR_CLOCK_STATUS.D_DATE}<=DATETIME(" & Year(Dtp_ToDate.Text) & "," & Format(Month(Dtp_ToDate.Text), "00") & "," & Format(Day(Dtp_ToDate.Text), "00") & ")"
    
    If Trim(Cmb_Company) <> "" Then
       vF2 = "{V_PR_CLOCK_STATUS.C_COMPANYNAME}='" & Trim(Left(Trim(Cmb_Company), 100)) & "'"
    End If
    
    If Trim(Cmb_Branch) <> "" Then
       vF3 = "{V_PR_CLOCK_STATUS.C_BRANCH}='" & Trim(Cmb_Branch) & "'"
    End If
    
    If Trim(Cmb_Dept) <> "" Then
       vF4 = "{V_PR_CLOCK_STATUS.C_DEPT}='" & Trim(Cmb_Dept) & "'"
    End If
      
    If Trim(Cmb_Desig) <> "" Then
       vF5 = "{V_PR_CLOCK_STATUS.C_DESIG}='" & Trim(Cmb_Desig) & "'"
    End If
      
      
    SelFor = ReportFilterOption(vF1, vF2, vF3, vF4, vF5)
   
End Sub

Private Sub Assign_Add_Filter_Rep()
  Dim i As Integer
  Dim tmpStr As String
  
    SelFilFor = ""
    
    For i = 1 To Va_Filter.DataRowCnt
        Va_Filter.Row = i
        Va_Filter.Col = 4
           If Trim(Va_Filter.Text) <> "" Then
              tmpStr = Trim(Va_Filter.Text)
           End If
        Va_Filter.Col = 3
           If Trim(Va_Filter.Text) <> "" Then
              SelFilFor = Make_Rep_Filter_Option(SelFilFor, Trim(Va_Filter.Text), i, tmpStr)
           End If
    Next i
   
End Sub

Private Function Make_Rep_Filter_Option(vTotStr As String, vOptStr As String, vRow As Integer, Optional vOptStr2 As String)
   Dim tmpStr As String, tmpOption As String
   Dim tmpField
    
    Va_Filter.Row = vRow
    Va_Filter.Col = 1
       Va_Filter.TypeComboBoxIndex = Va_Filter.TypeComboBoxCurSel
       If Trim(Va_Filter.TypeComboBoxString) <> "" Then
          tmpField = Split(Trim(Va_Filter.TypeComboBoxString), "~")
       Else
          Make_Rep_Filter_Option = vTotStr
          Exit Function
       End If
    Va_Filter.Col = 2
       Va_Filter.TypeComboBoxIndex = Va_Filter.TypeComboBoxCurSel
       tmpOption = Right(Trim(Va_Filter.TypeComboBoxString), 1)
       If Trim(tmpOption) = "" Then
          Make_Rep_Filter_Option = vTotStr
          Exit Function
       End If
       
    If Val(tmpOption) = 7 Then
       If vOptStr = "" Or vOptStr2 = "" Then
          Make_Rep_Filter_Option = vTotStr
          Exit Function
       End If
    End If
       
    If Val(tmpOption) = 7 Then
       tmpStr = " and " & tmpField(1) & " between '"
    Else
       If InStr(1, UCase(tmpField(1)), "A.", vbTextCompare) > 0 Then
          tmpStr = "{PR_CLOCK_EMP." & Replace(UCase(tmpField(1)), "A.", "") & "}"
       ElseIf InStr(1, UCase(tmpField(1)), "B.", vbTextCompare) > 0 Then
          tmpStr = "{PR_EMP_MST." & Replace(UCase(tmpField(1)), "B.", "") & "}"
       Else
          tmpStr = "{PR_CLOCK_EMP." & UCase(tmpField(1)) & "}"
       End If
    End If
    
    If Val(tmpOption) = 7 Then
       ' to avoid
    ElseIf Val(tmpOption) >= 4 And Val(tmpOption) <= 6 Then
       If tmpOption = "4" Then
          tmpStr = tmpStr & " like ['%"
       ElseIf tmpOption = "5" Then
          tmpStr = tmpStr & " like ['"
       ElseIf tmpOption = "6" Then
          tmpStr = tmpStr & " like ['%"
       End If
    Else
       If tmpOption = "1" Then
          tmpStr = tmpStr & " = "
       ElseIf tmpOption = "2" Then
          tmpStr = tmpStr & " < "
       ElseIf tmpOption = "3" Then
          tmpStr = tmpStr & " > "
       ElseIf tmpOption = "8" Then
          tmpStr = tmpStr & " <> "
       End If
        
       If tmpField(2) <> "N" Then
          tmpStr = tmpStr & "'"
       End If
    End If
    
    If tmpField(2) = "D" Then
       vOptStr = Format(vOptStr, "yyyy-mm-dd")
       vOptStr2 = Format(vOptStr2, "yyyy-mm-dd")
    End If
    
    If Val(tmpOption) = 7 Then
       tmpStr = tmpStr & vOptStr & "' and '" & vOptStr2 & "'"
    Else
       tmpStr = tmpStr & vOptStr
    End If
    
    If Val(tmpOption) = 7 Then
       ' to avoid
    ElseIf Val(tmpOption) >= 4 And Val(tmpOption) <= 6 Then
       If tmpOption = "4" Then
          tmpStr = tmpStr & "%']"
       ElseIf tmpOption = "5" Then
          tmpStr = tmpStr & "%']"
       ElseIf tmpOption = "6" Then
          tmpStr = tmpStr & "']"
       End If
    Else
       If tmpField(2) <> "N" Then
          tmpStr = tmpStr & "'"
       End If
    End If
    
    Make_Rep_Filter_Option = ReportFilterOption(vTotStr, tmpStr)
       
End Function

Private Sub Btn_Hide_Click()
  Dim i As Integer
  
    For i = Va_Details.SelBlockCol To Va_Details.SelBlockCol2
        If i > 2 Then
           Va_Details.Row = -1
           Va_Details.Col = i
             Va_Details.ColHidden = True
        End If
    Next i
End Sub

Private Sub Btn_Restore_Click()
  Dim i As Integer
    For i = 1 To Va_Details.MaxCols
        Va_Details.Row = -1
        Va_Details.Col = i
           Va_Details.ColHidden = False
    Next i
    Call Spread_ColHide
End Sub


Private Sub Btn_Discp_Click()
  
    If Not IsDate(Dtp_FromDate.Text) Then
       MsgBox "Date From should not be Empty", vbInformation, "Information"
       Dtp_FromDate.SetFocus
       Exit Sub
    ElseIf Not IsDate(Dtp_ToDate.Text) Then
       MsgBox "Date From should not be Empty", vbInformation, "Information"
       Dtp_ToDate.SetFocus
       Exit Sub
    End If
    
    If CDate(Dtp_ToDate.Text) - CDate(Dtp_FromDate.Text) > 40 Then
       MsgBox "You choosen big date range. Please choose one month period", vbInformation, "Information"
       Dtp_ToDate.SetFocus
       Exit Sub
    End If
  
    If Va_Details.DataRowCnt > 0 Then
       If MsgBox("The Data will be refesh?. Do you want to display?", vbYesNo + vbDefaultButton2, "Confirmation") = vbNo Then
          Exit Sub
       Else
          Va_Details.MaxRows = 0
          Va_Details.MaxRows = 50
       End If
    End If
  
    If Trim(Txtc_EmployeeName) = "" Then
       Lbl_Info.Caption = ""
    End If
    
    Clear_Spread Va_Details
    
    Call Display_Discp_Records("SC")  ' Single Clocking
    Call Display_Discp_Records("WC")  ' Wrong Code. System not Recoginsed the code.
    Call Display_Discp_Records("WP")  ' Wrong Period. Period should be 0.5 or 1
    
    Call Display_Discp_Records("AJ")  ' Attn exits, but not join
    Call Display_Discp_Records("AL")  ' Attn exits but left
    
    Call Display_Discp_Records("PL")  ' Present but assign Leave
    Call Display_Discp_Records("NH")  ' Negative work hrs
    Call Display_Discp_Records("BH")  ' No Work hrs
    Call Display_Discp_Records("HD")  ' Half day working, but other half day no leave assigned / less than 5 hrs working
    Call Display_Discp_Records("HP")  ' Present half day, but other half day no leave assigned
    Call Display_Discp_Records("HL")  ' Half day leave, but no in/out
    
    
    Call Display_Discp_Records("LE")  ' Leave assign before doj / not eligible
    Call Display_Discp_Records("LH")  ' Leave assigned on Public Holiday
    Call Display_Discp_Records("LT")  ' Leave assigned on Saturday
    
    Call Display_Discp_Records("CL")  ' CL assigned. but not eligible
    Call Display_Discp_Records("SL")  ' SL assigned. but not eligible
    Call Display_Discp_Records("VL")  ' VL assigned. but not eligible
    
End Sub

Private Sub Display_Discp_Records(ByVal vDispType As String)
On Error GoTo Err_Display
  Dim DyDisp As New ADODB.Recordset
  Dim i As Long, j As Long
  Dim tmpPeriod As String, tmpEmpNo As String
  
  i = 0: j = 0
  
  g_Sql = "select a.c_empno, c.c_companyname, b.c_branch, b.c_dept, b.c_desig, b.c_emptype, " & _
          "a.d_date, a.n_wkday, a.n_arrtime, a.n_latehrs, a.n_deptime, a.n_earlhrs, a.n_permhrs, " & _
          "a.n_workhrs, a.n_overtime, a.n_present, a.c_presabs, a.c_shift, a.n_workhrs_fp, a.c_chq, a.c_flag, a.c_usr_id, a.d_modified, " & _
          "b.c_name, b.c_othername, b.d_doj, b.d_dol, " & _
          "a.n_otin, a.n_otout, a.n_ot15, a.n_ot20, a.n_ot30, a.n_otin_str, a.n_otout_str, " & _
          "a.n_time1, a.n_time2, a.n_time3, a.n_time4, a.n_time5, a.n_time6, " & _
          "a.n_time7, a.n_time8, a.n_time9, a.n_time10, a.n_time11, a.n_time12 " & _
          "from pr_clock_emp a,  pr_emp_mst b, pr_company_mst c " & _
          "Where a.c_empno = b.c_empno and b.c_company = c.c_company and b.c_rec_sta = 'A' "
          
  If Trim(vDispType) = "SC" Then   ' Single clocking
     g_Sql = g_Sql & " and (a.c_chq = '*' or a.n_workhrs = 0) and " & _
                     " a.n_time1 > 0 and c_flag = 'B' "
  
  ElseIf Trim(vDispType) = "AJ" Then ' Attendance exits, but not join
     g_Sql = g_Sql & " and a.c_presabs not in ('WO') and " & _
                     " a.d_date < b.d_doj "
  
  ElseIf Trim(vDispType) = "AL" Then ' Attendance exits, but left
     g_Sql = g_Sql & " and a.c_presabs not in ('WO') and " & _
                     " a.d_date > b.d_dol "
  
  ElseIf Trim(vDispType) = "PL" Then ' Present but assign Leave
     g_Sql = g_Sql & " and a.n_workhrs > 0 and n_present = 1 and " & _
                     "c_presabs not in ('P','PH') "
  
  ElseIf Trim(vDispType) = "NH" Then ' Negative Work hrs
     g_Sql = g_Sql & " and a.n_workhrs < 0 "
  
  ElseIf Trim(vDispType) = "BH" Then ' No Work hrs
     g_Sql = g_Sql & " and a.c_presabs in ('P') and " & _
                     " a.n_workhrs = 0 and a.n_time1 = 0 "
  
  ElseIf Trim(vDispType) = "HD" Then ' half day working (less then 5 hrs working)
     g_Sql = g_Sql & " and a.c_presabs = 'P' and a.n_present = 1 and " & _
                     " a.n_workhrs > 0 and a.n_workhrs < 5 and a.n_wkday >= 1 and a.n_wkday <= 5 "
  
  ElseIf Trim(vDispType) = "HP" Then ' half day working, other half no leave
     g_Sql = g_Sql & " and a.c_presabs = 'P' and a.n_present = 0.5 and " & _
                     " a.n_workhrs > 0 "
  
  ElseIf Trim(vDispType) = "HL" Then ' half day working, other half no leave
     g_Sql = g_Sql & " and a.c_presabs not in ('P','PH') and a.n_present = 0.5 and " & _
                     " a.n_workhrs = 0 "
                     
  ElseIf Trim(vDispType) = "WC" Then ' Wrong Code. System not Recoginsed the code.
     g_Sql = g_Sql & " and c_presabs not in ('" & Trim(vLeaveCodes) & "') "
  
  ElseIf Trim(vDispType) = "WP" Then ' Wrong Period. Period should be 0.5 or 1
     g_Sql = g_Sql & " and (n_present <> 0.5 and n_present <> 1)"
     
  ElseIf Trim(vDispType) = "LE" Then ' Leave is not Eligble
     g_Sql = g_Sql & " and a.n_present > 0 and a.c_presabs in ('CL','SL','VL','ML') and " & _
             "a.d_date < dateadd(dd,0,b.d_doj) "
 
  ElseIf Trim(vDispType) = "LH" Then ' Leave assigned on Public Holiday / Sunday
     g_Sql = g_Sql & " and a.c_presabs not in ('P','PH','WO','ML','VL') and (a.n_wkday = 7 or a.c_presabs = 'PH') "
  
  ElseIf Trim(vDispType) = "LT" Then ' Leave assigned on Saturday
     g_Sql = g_Sql & " and a.c_presabs not in ('P','PH','WO','ML','VL') and a.n_wkday = 6 " & _
                     " and left(b.c_daywork,2) = '5D' "
  
  ElseIf Trim(vDispType) = "CL" Then ' Local Leave assigned. but not eligible
     g_Sql = "select a.c_empno, d.c_companyname, b.c_branch, b.c_dept, b.c_desig, b.c_emptype, " & _
          "a.d_date, a.n_wkday, a.n_arrtime, a.n_latehrs, a.n_deptime, a.n_earlhrs, a.n_permhrs, " & _
          "a.n_workhrs, a.n_overtime, a.n_present, a.c_presabs, a.c_shift, a.n_workhrs_fp, a.c_chq, a.c_flag, a.c_usr_id, a.d_modified,  " & _
          "b.c_name, b.c_othername, b.d_doj, b.d_dol, " & _
          "a.n_otin, a.n_otout, a.n_ot15, a.n_ot20, a.n_ot30, a.n_otin_str, a.n_otout_str, " & _
          "a.n_time1, a.n_time2, a.n_time3, a.n_time4, a.n_time5, a.n_time6, " & _
          "a.n_time7, a.n_time8, a.n_time9, a.n_time10, a.n_time11, a.n_time12, IsNull(c.n_clbal,0) n_balance " & _
          "from pr_clock_emp a, pr_company_mst d,  pr_emp_mst b left outer join pr_emp_leave_dtl c " & _
          "    on b.c_empno = c.c_empno and c.c_leave = 'CL' and c.c_yrstatus = 'C' " & _
          "Where a.c_empno = b.c_empno and b.c_company = d.c_company and a.c_presabs = 'CL' and " & _
          "(c.c_leave is null or c.n_clbal < 0) and b.c_rec_sta = 'A' "

  ElseIf Trim(vDispType) = "VL" Then ' Family Leave assigned. but not eligible
     g_Sql = "select a.c_empno, d.c_companyname, b.c_branch, b.c_dept, b.c_desig, b.c_emptype, " & _
          "a.d_date, a.n_wkday, a.n_arrtime, a.n_latehrs, a.n_deptime, a.n_earlhrs, a.n_permhrs, " & _
          "a.n_workhrs, a.n_overtime, a.n_present, a.c_presabs, a.c_shift, a.n_workhrs_fp, a.c_chq, a.c_flag, a.c_usr_id, a.d_modified, " & _
          "b.c_name, b.c_othername, b.d_doj, b.d_dol, " & _
          "a.n_otin, a.n_otout, a.n_ot15, a.n_ot20, a.n_ot30, a.n_otin_str, a.n_otout_str, " & _
          "a.n_time1, a.n_time2, a.n_time3, a.n_time4, a.n_time5, a.n_time6, " & _
          "a.n_time7, a.n_time8, a.n_time9, a.n_time10, a.n_time11, a.n_time12, IsNull(c.n_clbal,0) n_balance " & _
          "from pr_clock_emp a,  pr_company_mst d, pr_emp_mst b left outer join pr_emp_leave_dtl c " & _
          "    on b.c_empno = c.c_empno and c.c_leave = 'VL' and c.c_yrstatus = 'C' " & _
          "Where a.c_empno = b.c_empno and b.c_company = d.c_company and a.c_presabs = 'VL' and " & _
          "(c.c_leave is null or c.n_clbal < 0) and b.c_rec_sta = 'A' "

  ElseIf Trim(vDispType) = "SL" Then ' Sick Leave assigned. but not eligible
     g_Sql = "select a.c_empno, d.c_companyname, b.c_branch, b.c_dept, b.c_desig, b.c_emptype, " & _
          "a.d_date, a.n_wkday, a.n_arrtime, a.n_latehrs, a.n_deptime, a.n_earlhrs, a.n_permhrs, " & _
          "a.n_workhrs, a.n_overtime, a.n_present, a.c_presabs, a.c_shift, a.n_workhrs_fp,  a.c_chq, a.c_flag, a.c_usr_id, a.d_modified,  " & _
          "b.c_name, b.c_othername, b.d_doj, b.d_dol, " & _
          "a.n_otin, a.n_otout, a.n_ot15, a.n_ot20, a.n_ot30, a.n_otin_str, a.n_otout_str, " & _
          "a.n_time1, a.n_time2, a.n_time3, a.n_time4, a.n_time5, a.n_time6, " & _
          "a.n_time7, a.n_time8, a.n_time9, a.n_time10, a.n_time11, a.n_time12, IsNull(c.n_clbal,0) n_balance " & _
          "from pr_clock_emp a, pr_company_mst d, pr_emp_mst b left outer join pr_emp_leave_dtl c " & _
          "    on b.c_empno = c.c_empno and c.c_leave = 'SL' and c.c_yrstatus = 'C' " & _
          "Where a.c_empno = b.c_empno and b.c_company = d.c_company and a.c_presabs = 'SL' and " & _
          "(c.c_leave is null or c.n_clbal < 0) and b.c_rec_sta = 'A' "
  End If

  
  If IsDate(Dtp_FromDate.Text) And IsDate(Dtp_ToDate.Text) Then
     g_Sql = g_Sql & " and a.d_date >= '" & Is_Date(Dtp_FromDate.Text, "S") & "' and a.d_date <= '" & Is_Date(Dtp_ToDate.Text, "S") & "' "
  End If
          
  If Right(Trim(Txtc_EmployeeName), 7) <> "" Then
     g_Sql = g_Sql & " and a.c_empno = '" & Trim(Right(Trim(Txtc_EmployeeName), 7)) & "'"
  Else
     If Trim(Cmb_Company) <> "" Then
        g_Sql = g_Sql & " and b.c_company = '" & Trim(Right(Trim(Cmb_Company), 7)) & "'"
     End If
    
     If Trim(Cmb_Branch) <> "" Then
        g_Sql = g_Sql & " and b.c_branch = '" & Trim(Cmb_Branch) & "'"
     End If
            
     If Trim(Cmb_Dept) <> "" Then
        g_Sql = g_Sql & " and b.c_dept = '" & Trim(Cmb_Dept) & "'"
     End If
     
     If Trim(Cmb_Desig) <> "" Then
        g_Sql = g_Sql & " and b.c_desig = '" & Trim(Cmb_Desig) & "'"
     End If
     
     If Trim(Cmb_Shift) <> "" Then
        g_Sql = g_Sql & " and b.c_shiftcode = '" & Trim(Right(Trim(Cmb_Shift), 7)) & "'"
     End If
    
     If Trim(Cmb_EmpType) <> "" Then
        g_Sql = g_Sql & " and b.c_emptype = '" & Trim(Cmb_EmpType) & "'"
     End If
  End If
  
  g_Sql = g_Sql & strFilter
  
  g_Sql = g_Sql & " order by a.c_empno, a.d_date"
  
  Set DyDisp = Nothing
  DyDisp.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
  Va_Details.MaxRows = Va_Details.DataRowCnt + DyDisp.RecordCount + 1000
  
  If Va_Details.DataRowCnt > 0 Then
     j = Va_Details.DataRowCnt + 2
  End If
  If DyDisp.RecordCount > 0 Then
     DyDisp.MoveFirst
     For i = 1 To DyDisp.RecordCount
         Va_Details.Row = j + i
         Va_Details.Col = 2
            Va_Details.Text = Is_Null(DyDisp("c_empno").Value, False)
         Va_Details.Col = 3
            Va_Details.Text = Proper(Is_Null(DyDisp("c_name").Value, False)) & " " & Proper(Is_Null(DyDisp("c_othername").Value, False))
         Va_Details.Col = 4
            Va_Details.Text = Proper(Is_Null(DyDisp("c_companyname").Value, False))
         Va_Details.Col = 5
            Va_Details.Text = Proper(Is_Null(DyDisp("c_branch").Value, False))
         Va_Details.Col = 6
            Va_Details.Text = Proper(Is_Null(DyDisp("c_dept").Value, False))
         Va_Details.Col = 7
            Va_Details.Text = Proper(Is_Null(DyDisp("c_desig").Value, False))
         Va_Details.Col = 8
            Va_Details.Text = Proper(Is_Null(DyDisp("c_emptype").Value, False))
         Va_Details.Col = 9
            Va_Details.Text = ""
         
         Va_Details.Col = 10
            Va_Details.Text = Is_DateSpread(DyDisp("d_date").Value, True)
         Va_Details.Col = 11
            Va_Details.Text = WeekdayName(Is_Null(DyDisp("n_wkday").Value, True), True, vbMonday)
         Va_Details.Col = 12
            Va_Details.Text = Is_Null(DyDisp("c_shift").Value, False)
         
         Va_Details.Col = 13
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_arrtime").Value, True), True)
         Va_Details.Col = 14
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_latehrs").Value, True), True)
         
         Va_Details.Col = 15
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_deptime").Value, True), True)
         Va_Details.Col = 16
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_earlhrs").Value, True), True)
         
         Va_Details.Col = 17
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_workhrs").Value, True), True)
         Va_Details.Col = 18
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_overtime").Value, True), True)
         Va_Details.Col = 19
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_permhrs").Value, True), True)
         
         Va_Details.Col = 20
            Va_Details.Text = Is_Null(DyDisp("c_presabs").Value, False)
         Va_Details.Col = 21
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_present").Value, True), True)
         Va_Details.Col = 22
            Va_Details.Text = ""
         
         Va_Details.Col = 23
            Va_Details.Text = OTSignAddOnDisplay(DecimalToString(Is_Null(DyDisp("n_otin").Value, True)), Is_Null(DyDisp("n_otin_str").Value, False))
            If Trim(Right(Trim(Va_Details.Text), 1)) = "$" Then
               Va_Details.FontBold = True
            Else
               Va_Details.FontBold = False
            End If
         Va_Details.Col = 24
            Va_Details.Text = OTSignAddOnDisplay(DecimalToString(Is_Null(DyDisp("n_otout").Value, True)), Is_Null(DyDisp("n_otout_str").Value, False))
            If Trim(Right(Trim(Va_Details.Text), 1)) = "$" Then
               Va_Details.FontBold = True
            Else
               Va_Details.FontBold = False
            End If
         Va_Details.Col = 25
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_ot15").Value, True), True)
         Va_Details.Col = 26
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_ot20").Value, True), True)
         Va_Details.Col = 27
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_ot30").Value, True), True)
         
         Va_Details.Col = 29
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_time1").Value, True), True)
         Va_Details.Col = 30
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_time2").Value, True), True)
         Va_Details.Col = 31
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_time3").Value, True), True)
         Va_Details.Col = 32
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_time4").Value, True), True)
         Va_Details.Col = 33
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_time5").Value, True), True)
         Va_Details.Col = 34
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_time6").Value, True), True)
         
         Va_Details.Col = 35
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_time7").Value, True), True)
         Va_Details.Col = 36
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_time8").Value, True), True)
         Va_Details.Col = 37
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_time9").Value, True), True)
         Va_Details.Col = 38
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_time10").Value, True), True)
         Va_Details.Col = 39
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_time11").Value, True), True)
         Va_Details.Col = 40
            Va_Details.Text = Spread_NumFormat(Is_Null(DyDisp("n_time12").Value, True), True)
         Va_Details.Col = 41
            Va_Details.Text = ""
         
         Va_Details.Col = 42
            If Trim(vDispType) = "SC" Then
               Va_Details.Text = "Single Clocking / No Work Hrs Found."
            ElseIf Trim(vDispType) = "WC" Then
               Va_Details.Text = "Wrong Code. System not Recognized"
            ElseIf Trim(vDispType) = "WP" Then
               Va_Details.Text = "Wrong Period. Period should be 0.5 or 1"
            
            ElseIf Trim(vDispType) = "AJ" Then
               Va_Details.Text = "Attn exits but not Joint. DOJ " & Is_DateSpread(DyDisp("d_doj").Value, True)
            ElseIf Trim(vDispType) = "AL" Then
               Va_Details.Text = "Attn exits but Left. Date Left " & Is_DateSpread(DyDisp("d_dol").Value, True)
               
            ElseIf Trim(vDispType) = "PL" Then
               Va_Details.Text = "Present but Assigned Leave"
            ElseIf Trim(vDispType) = "NH" Then
               Va_Details.Text = "Negative Work Hrs Found"
            ElseIf Trim(vDispType) = "BH" Then
               Va_Details.Text = "No Work Hrs Found"
            
            ElseIf Trim(vDispType) = "HD" Then
               Va_Details.Text = "Less than 5.00 Work Hrs Found"
            ElseIf Trim(vDispType) = "HP" Then
               Va_Details.Text = "Present half day, other half day no leave assigned"
            ElseIf Trim(vDispType) = "HL" Then
               Va_Details.Text = "Leave half day, other half day no hours"
            
            ElseIf Trim(vDispType) = "LE" Then
               Va_Details.Text = "Leave is not eligible. DOJ " & Is_DateSpread(DyDisp("d_doj").Value, True)
            
            ElseIf Trim(vDispType) = "LH" Then
               Va_Details.Text = "Leave Assigned on Sunday / Public Holiday "
            ElseIf Trim(vDispType) = "LT" Then
               Va_Details.Text = "Leave Assigned on Saturday "
            
            ElseIf Trim(vDispType) = "CL" Then
               Va_Details.Text = "CL no balance. CL balance : " & Spread_NumFormat(DyDisp("n_balance").Value, True, 2)
            ElseIf Trim(vDispType) = "SL" Then
               Va_Details.Text = "SL no balance. SL balance : " & Spread_NumFormat(DyDisp("n_balance").Value, True, 2)
            ElseIf Trim(vDispType) = "VL" Then
               Va_Details.Text = "VL no balance. VL balance : " & Spread_NumFormat(DyDisp("n_balance").Value, True, 2)
            End If

         Va_Details.Col = 43
            Va_Details.Text = Proper(Is_Null(DyDisp("c_usr_id").Value, False))
         Va_Details.Col = 44
            Va_Details.Text = Format(DyDisp("d_modified").Value, "dd/mm/yy hh:mm:ss")
         
         tmpEmpNo = Is_Null(DyDisp("c_empno").Value, False)
        DyDisp.MoveNext
     Next i
  Else
    ' MsgBox "No details found", vbInformation, "Information"
  End If
 
  Va_Details.MaxRows = Va_Details.DataRowCnt
  DoEvents
 
 Exit Sub

Err_Display:
    Screen.MousePointer = vbDefault
    MsgBox "Error While Display " & Err.Description
End Sub


Private Sub Save_Pr_Clock_Tran_Discp()
On Error GoTo Err_Save
  Dim rs As New ADODB.Recordset
  Dim RepTitle As String
  Dim i As Integer
  
    Screen.MousePointer = vbHourglass
    g_SaveFlagNull = True
   
    g_Sql = "truncate table pr_clock_tran_discp"
    CON.Execute g_Sql
    
    CON.BeginTrans
    Set rs = Nothing
    g_Sql = "select * from pr_clock_tran_discp where 1=2"
    rs.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
    
    For i = 1 To Va_Details.DataRowCnt
        Va_Details.Row = i
        Va_Details.Col = 2
           If Trim(Va_Details.Text) <> "" Then
              rs.AddNew
                 
              Va_Details.Col = 2
                 rs("c_empno").Value = Is_Null(Va_Details.Text, False)
              Va_Details.Col = 10
                 rs("d_date").Value = Is_Date(Va_Details.Text, "S")
              Va_Details.Col = 42
                 rs("c_remarks").Value = Is_Null(Va_Details.Text, False)
              rs.Update
           End If
    Next i
    CON.CommitTrans
          
    Screen.MousePointer = vbDefault
    g_SaveFlagNull = False

 Exit Sub
 
Err_Save:
   Screen.MousePointer = vbDefault
   g_SaveFlagNull = False
   MsgBox "Error while prepare Discrepancy " & Err.Description, vbInformation, "Information"
End Sub

Public Function Chk_Mand() As Boolean
     If Not IsDate(Dtp_FromDate.Text) Then
        MsgBox "Pls enter the From date", vbInformation, "Information"
        Dtp_FromDate.SetFocus
        Chk_Mand = False
        Exit Function
     ElseIf Not IsDate(Dtp_ToDate.Text) Then
        MsgBox "Pls enter the to date", vbInformation, "Information"
        Dtp_ToDate.SetFocus
        Chk_Mand = False
        Exit Function
     ElseIf Va_Details.DataRowCnt <= 0 Then
        MsgBox "There is no data to save", vbInformation, "Information"
        Va_Details.SetFocus
        Chk_Mand = False
        Exit Function
     ElseIf CDate(Dtp_ToDate.Text) - CDate(Dtp_FromDate.Text) > 40 Then
        MsgBox "You have choosen big date range. Please choose one month period", vbInformation, "Information"
        Dtp_ToDate.SetFocus
        Chk_Mand = False
        Exit Function
     End If
     
     Chk_Mand = True
End Function

Private Sub Btn_Export_Click()
On Error GoTo Err_Export
  Dim SelFor As String
  Dim vFileName As String
  
     If Va_Details.DataRowCnt = 0 Then
        Exit Sub
     End If
  
     If MsgBox("Do you want to Export to CSV file?", vbYesNo + vbDefaultButton2, "Confirmation") = vbNo Then
        Exit Sub
     End If
  
     Screen.MousePointer = vbHourglass
    
     CON.BeginTrans
        Save_Pr_Export_Csv_Rep
     CON.CommitTrans
     
     
     comDialog.FileName = "AttnDtl.CSV"
     comDialog.ShowSave
     vFileName = Trim(comDialog.FileName)
     
     If vFileName = "AttnDtl.CSV" Then  'user cancel the export to csv file in save dialog box.
        Screen.MousePointer = vbDefault
        Exit Sub
     ElseIf Trim(Right(Trim(vFileName), 4)) <> ".CSV" Then
        vFileName = vFileName & ".CSV"
     End If

     Call CsvFileExport_Process(vFileName)

     Screen.MousePointer = vbDefault
     
     MsgBox "Transfered Successfully", vbInformation, "Information"
     
  Exit Sub

Err_Export:
    Screen.MousePointer = vbDefault
    MsgBox "Error while Generating - " + Err.Description, vbInformation, "Information"
    CON.RollbackTrans
End Sub


Private Sub Save_Pr_Export_Csv_Rep()
  Dim i As Long, j As Long, Ctr As Long
  Dim vStr As String, vEmpNo As String
  
    Ctr = 0
    g_Sql = "truncate table pr_export_csv "
    CON.Execute (g_Sql)
    
    Set rs = Nothing
    g_Sql = "select * from pr_export_csv where 1=2"
    rs.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
    
    For i = 0 To Va_Details.DataRowCnt
        Va_Details.Row = i
        Va_Details.Col = 2
           If Trim(Va_Details.Text) <> "" Then
              vStr = "": vEmpNo = ""
              Ctr = Ctr + 1
              Va_Details.Row = i
              Va_Details.Col = 2
                 vEmpNo = Trim(Va_Details.Text)
                 vStr = Trim(Va_Details.Text)
              
              For j = 3 To Va_Details.MaxCols
                  Va_Details.Row = i
                  Va_Details.Col = j
                     If Va_Details.ColHidden = False Then
                        vStr = vStr & "," & Trim(Va_Details.Text)
                     End If
              Next j
              
              rs.AddNew
              rs("n_seq").Value = Ctr
              rs("c_refno").Value = vEmpNo
              rs("c_csv").Value = vStr
              rs("c_usr_id").Value = g_UserName
              rs("d_created").Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
              rs.Update
           End If
    Next i
End Sub

Private Sub GetLeaveCodes()
  Dim rsChk As New ADODB.Recordset
  Dim i As Integer
  
    vLeaveCodes = ""
    Set rsChk = Nothing
    g_Sql = "select c_leave from pr_leave_mst where c_rec_sta = 'A'"
    rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    
    rsChk.MoveFirst
    For i = 1 To rsChk.RecordCount
       If vLeaveCodes = "" Then
          vLeaveCodes = Is_Null(rsChk("c_leave").Value, False)
       Else
          vLeaveCodes = vLeaveCodes & "', '" & Is_Null(rsChk("c_leave").Value, False)
       End If
       rsChk.MoveNext
    Next i
   
End Sub

Private Sub GetLeaveOpenBalance(ByVal vEmpNo As String, ByVal vRow As Long)
  Dim rsChk As New ADODB.Recordset
  Dim i As Integer, vLeaveBal As Double
    
    Set rsChk = Nothing
    g_Sql = "Select a.c_empno, a.c_leave, a.n_opbal, a.n_alloted, a.n_adjusted, sum(b.n_present) n_utilised " & _
            "From pr_emp_leave_dtl a left outer join pr_clock_emp b on a.c_empno = b.c_empno and year(a.d_prfrom) = year(b.d_date) and " & _
            "                        b.c_presabs in ('CL','SL','VL') and b.d_date < '" & Is_Date(Dtp_FromDate.Text, "S") & "' " & _
            "Where a.c_yrstatus = 'C' and a.c_empno = '" & Trim(vEmpNo) & "' " & _
            "Group by a.c_empno, a.c_leave, a.n_opbal, a.n_alloted, a.n_adjusted "
    rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    
    For i = 1 To rsChk.RecordCount
        vLeaveBal = (Is_Null(rsChk("n_opbal").Value, True) + Is_Null(rsChk("n_alloted").Value, True)) - (Is_Null(rsChk("n_utilised").Value, True) + Is_Null(rsChk("n_adjusted").Value, True))
        Va_Details.Row = vRow
        Va_Details.Col = 46
           If Is_Null(rsChk("c_leave").Value, False) = "CL" Then
              Va_Details.Text = vLeaveBal
           End If
        Va_Details.Col = 47
           If Is_Null(rsChk("c_leave").Value, False) = "SL" Then
              Va_Details.Text = vLeaveBal
           End If
        Va_Details.Col = 48
           If Is_Null(rsChk("c_leave").Value, False) = "VL" Then
              Va_Details.Text = vLeaveBal
           End If
        rsChk.MoveNext
    Next i
End Sub

Private Sub DisplayLeaveBalance(vRow As Long)
  Dim i As Long, vEmpStartRow As Long, vEmpEndRow As Long
  Dim vCLBal As Double, vSLBal As Double, vVLBal As Double, vPresent As Double
  Dim vStatus As String, vEmpNo As String
  
    vEmpStartRow = 0: vEmpEndRow = 0: vCLBal = 0: vSLBal = 0: vVLBal = 0
    
    Va_Details.Row = vRow
    Va_Details.Col = 2
       vEmpNo = Trim(Va_Details.Text)
    
    For i = vRow To 1 Step -1
        Va_Details.Row = i
        Va_Details.Col = 2
           If Trim(Va_Details.Text) <> "" Then
              If Trim(vEmpNo) <> Trim(Va_Details.Text) Then
                 vEmpStartRow = i + 2
                 Exit For
              End If
           End If
    Next i
    If vEmpStartRow = 0 Then
       vEmpStartRow = 1
    End If
    
    For i = vRow To Va_Details.DataRowCnt
        Va_Details.Row = i
        Va_Details.Col = 2
           If Trim(Va_Details.Text) <> "" Then
              If Trim(vEmpNo) <> Trim(Va_Details.Text) Then
                 vEmpEndRow = i
                 Exit For
              End If
           End If
    Next i
    If vEmpEndRow = 0 Then
       vEmpEndRow = Va_Details.DataRowCnt
    End If
    
    Va_Details.Row = vEmpStartRow
    Va_Details.Col = 46
       vCLBal = Is_Null(Va_Details.Text, True)
    Va_Details.Col = 47
       vSLBal = Is_Null(Va_Details.Text, True)
    Va_Details.Col = 48
       vVLBal = Is_Null(Va_Details.Text, True)
       
    
    For i = vEmpStartRow To vEmpEndRow
        Va_Details.Row = i
        Va_Details.Col = 20
           vStatus = Trim(Va_Details.Text)
        Va_Details.Col = 21
           vPresent = Is_Null(Va_Details.Text, True)
        
        If vStatus = "CL" Then
           vCLBal = vCLBal - vPresent
           Va_Details.Col = 42
              Va_Details.Text = "CL Balance : " & Str(vCLBal)
        End If
        If vStatus = "SL" Then
           vSLBal = vSLBal - vPresent
           Va_Details.Col = 42
              Va_Details.Text = "SL Balance : " & Str(vSLBal)
        End If
        If vStatus = "VL" Then
           vVLBal = vVLBal - vPresent
           Va_Details.Col = 42
              Va_Details.Text = "VL Balance : " & Str(vVLBal)
        End If
    Next i
    
End Sub

Private Function DecimalToString(ByVal vDec As Double) As String
  Dim tmpStr As String
  
    tmpStr = Spread_NumFormat(vDec, True, 2)
    If Len(tmpStr) > 0 Then
       If Val(tmpStr) < 1 Then
          tmpStr = "0" & tmpStr
       End If
    End If
    DecimalToString = Trim(tmpStr)
End Function

Private Function OTSignAddOnDisplay(ByVal vStr As String, vOtApp As String) As String
   If Len(vStr) > 0 Then
      If Trim(Right(vOtApp, 1)) = "$" Then
         vStr = OTSignRemove(vStr)
         vStr = vStr & " $"
      End If
   End If
   OTSignAddOnDisplay = Trim(vStr)
End Function

Private Function OTSignAdd(ByVal vStr As String) As String
   If Len(Trim(vStr)) > 0 Then
      If Trim(Right(vStr, 1)) = "$" Then
         vStr = OTSignRemove(vStr)
      End If
      vStr = Trim(vStr) & " $"
   End If
   OTSignAdd = Trim(vStr)
End Function

Private Function OTSignRemove(ByVal vStr As String) As String
    If Len(vStr) > 0 Then
       If Trim(Right(vStr, 1)) = "$" Then
          vStr = Trim(Left(vStr, Len(vStr) - 1))
       End If
    End If
    OTSignRemove = Trim(vStr)
End Function

Private Sub Assign_OTSign()
 Dim i As Integer
 
    For i = 1 To Va_Details.DataRowCnt
        Va_Details.Row = i
        Va_Details.Col = 1
           If Va_Details.Value = 1 Then
              Va_Details.Col = 23
                 If Va_Details.FontBold = True Then
                    Va_Details.Text = OTSignAdd(Va_Details.Text)
                 End If
              Va_Details.Col = 24
                 If Va_Details.FontBold = True Then
                    Va_Details.Text = OTSignAdd(Va_Details.Text)
                 End If
           End If
    Next i
End Sub

Private Function ChkValidTime(ByVal vTime As Double) As Double
  Dim tmpVal As Long
  
    If vTime > 0 Then
       tmpVal = (vTime * 100) Mod 100
       If tmpVal >= 60 Then
          vTime = 0
       End If
    End If
    ChkValidTime = vTime
End Function

Private Function GetLeaveName(ByVal vLeave As String) As String
  Dim rsChk As New ADODB.Recordset
  
    If Trim(vLeave) = "" Then
       Exit Function
    End If
  
    Set rsChk = Nothing
    g_Sql = "select c_leavename from pr_leave_mst where c_rec_sta = 'A' and c_leave = '" & Trim(vLeave) & "'"
    rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If rsChk.RecordCount > 0 Then
       GetLeaveName = Is_Null(rsChk("c_leavename").Value, False)
    End If
End Function
