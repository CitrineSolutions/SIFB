VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_Company 
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
      TabIndex        =   81
      Top             =   -75
      Width           =   12285
      Begin VB.CommandButton Btn_View 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   1005
         Picture         =   "frm_Company.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Save 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   30
         Picture         =   "frm_Company.frx":36B0
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Exit 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   4470
         Picture         =   "frm_Company.frx":6D39
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Clear 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   1725
         Picture         =   "frm_Company.frx":A399
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Delete 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   2715
         Picture         =   "frm_Company.frx":DA09
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   700
      End
      Begin VB.CommandButton Btn_Print 
         BackColor       =   &H00FFFFFF&
         Height          =   720
         Left            =   3435
         Picture         =   "frm_Company.frx":110B3
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   700
      End
   End
   Begin VB.TextBox Txtc_CompanyName 
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
      Left            =   9585
      MaxLength       =   25
      TabIndex        =   8
      Top             =   1245
      Width           =   2100
   End
   Begin VB.TextBox Txtc_DisplayName 
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
      Left            =   3435
      MaxLength       =   100
      TabIndex        =   7
      Top             =   1230
      Width           =   4290
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3975
      Left            =   120
      TabIndex        =   52
      Top             =   1635
      Width           =   12285
      _ExtentX        =   21669
      _ExtentY        =   7011
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Company Details"
      TabPicture(0)   =   "frm_Company.frx":146A1
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra_address"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Social Security"
      TabPicture(1)   =   "frm_Company.frx":146BD
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      Begin VB.Frame fra_address 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3570
         Left            =   90
         TabIndex        =   66
         Top             =   315
         Width           =   12105
         Begin VB.TextBox Txtc_Desig 
            Height          =   315
            Left            =   7395
            MaxLength       =   100
            TabIndex        =   18
            Top             =   1890
            Width           =   3645
         End
         Begin VB.TextBox Txtc_Contact 
            Height          =   315
            Left            =   1860
            MaxLength       =   100
            TabIndex        =   17
            Top             =   1920
            Width           =   3645
         End
         Begin VB.TextBox Txtc_BrnNo 
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   9030
            MaxLength       =   25
            TabIndex        =   21
            Top             =   2640
            Width           =   2000
         End
         Begin VB.TextBox Txtc_VatNo 
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1860
            MaxLength       =   25
            TabIndex        =   19
            Top             =   2640
            Width           =   2000
         End
         Begin VB.TextBox Txtc_TanNo 
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   5595
            MaxLength       =   25
            TabIndex        =   20
            Top             =   2640
            Width           =   2000
         End
         Begin VB.TextBox Txtc_Ad3 
            Height          =   315
            Left            =   1860
            MaxLength       =   100
            TabIndex        =   11
            Top             =   960
            Width           =   3645
         End
         Begin VB.TextBox Txtc_Ad2 
            Height          =   315
            Left            =   1860
            MaxLength       =   100
            TabIndex        =   10
            Top             =   645
            Width           =   3645
         End
         Begin VB.TextBox Txtc_Ad1 
            Height          =   315
            Left            =   1860
            MaxLength       =   100
            TabIndex        =   9
            Top             =   315
            Width           =   3645
         End
         Begin VB.TextBox Txtc_Mobile 
            Height          =   315
            Left            =   7365
            MaxLength       =   100
            TabIndex        =   14
            Top             =   630
            Width           =   3645
         End
         Begin VB.TextBox Txtc_Tel 
            Height          =   315
            Left            =   7365
            MaxLength       =   100
            TabIndex        =   13
            Top             =   315
            Width           =   3645
         End
         Begin VB.TextBox Txtc_Email 
            Height          =   315
            Left            =   7365
            MaxLength       =   100
            TabIndex        =   15
            Top             =   945
            Width           =   3645
         End
         Begin VB.TextBox Txtc_Country 
            Height          =   315
            Left            =   1860
            MaxLength       =   100
            TabIndex        =   12
            Top             =   1275
            Width           =   3645
         End
         Begin VB.TextBox Txtc_Web 
            Height          =   315
            Left            =   7365
            MaxLength       =   100
            TabIndex        =   16
            Top             =   1260
            Width           =   3645
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000005&
            X1              =   -1455
            X2              =   12030
            Y1              =   2415
            Y2              =   2385
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Designation"
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
            Left            =   6270
            TabIndex        =   83
            Top             =   1935
            Width           =   975
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Contact Person"
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
            Left            =   465
            TabIndex        =   82
            Top             =   1965
            Width           =   1275
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "BRN No."
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
            Left            =   8355
            TabIndex        =   75
            Top             =   2685
            Width           =   615
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "VAT No."
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
            Left            =   1065
            TabIndex        =   74
            Top             =   2670
            Width           =   645
         End
         Begin VB.Line Line5 
            BorderColor     =   &H80000005&
            X1              =   -1455
            X2              =   12030
            Y1              =   1725
            Y2              =   1695
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TAN No."
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
            Left            =   4710
            TabIndex        =   73
            Top             =   2685
            Width           =   630
         End
         Begin VB.Label lbl_add 
            Alignment       =   1  'Right Justify
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
            Height          =   210
            Left            =   975
            TabIndex        =   72
            Top             =   360
            Width           =   720
         End
         Begin VB.Label lbl_contact 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Phone"
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
            Left            =   6720
            TabIndex        =   71
            Top             =   367
            Width           =   525
         End
         Begin VB.Label lbl_ema 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "E-mail"
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
            Left            =   6750
            TabIndex        =   70
            Top             =   975
            Width           =   495
         End
         Begin VB.Label lbl_tel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Mobile"
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
            Left            =   6690
            TabIndex        =   69
            Top             =   682
            Width           =   555
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Country"
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
            Left            =   1035
            TabIndex        =   68
            Top             =   1327
            Width           =   660
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Web"
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
            Left            =   6885
            TabIndex        =   67
            Top             =   1305
            Width           =   360
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
         Height          =   3090
         Left            =   -74940
         TabIndex        =   53
         Top             =   315
         Width           =   12150
         Begin VB.TextBox Txtn_ComEPZMin 
            Alignment       =   1  'Right Justify
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   9075
            TabIndex        =   44
            Top             =   1598
            Width           =   1400
         End
         Begin VB.TextBox Txtn_EmpEPZMin 
            Alignment       =   1  'Right Justify
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   4230
            TabIndex        =   32
            Top             =   1598
            Width           =   1400
         End
         Begin VB.TextBox Txtn_ComEPZMax 
            Alignment       =   1  'Right Justify
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   10485
            TabIndex        =   45
            Top             =   1605
            Width           =   1400
         End
         Begin VB.TextBox Txtn_EmpEPZMax 
            Alignment       =   1  'Right Justify
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   5640
            TabIndex        =   33
            Top             =   1598
            Width           =   1400
         End
         Begin VB.TextBox Txtn_EmpEPZ 
            Alignment       =   1  'Right Justify
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   2820
            TabIndex        =   31
            Top             =   1598
            Width           =   1400
         End
         Begin VB.TextBox Txtn_ComEPZ 
            Alignment       =   1  'Right Justify
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   7680
            TabIndex        =   43
            Top             =   1598
            Width           =   1400
         End
         Begin VB.TextBox Txtn_ComLevyMin 
            Alignment       =   1  'Right Justify
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   9075
            TabIndex        =   41
            Top             =   1305
            Width           =   1400
         End
         Begin VB.TextBox Txtn_EmpLevyMin 
            Alignment       =   1  'Right Justify
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   4230
            TabIndex        =   29
            Top             =   1290
            Width           =   1400
         End
         Begin VB.TextBox Txtn_ComLevyMax 
            Alignment       =   1  'Right Justify
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   10485
            TabIndex        =   42
            Top             =   1305
            Width           =   1400
         End
         Begin VB.TextBox Txtn_EmpLevyMax 
            Alignment       =   1  'Right Justify
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   5640
            TabIndex        =   30
            Top             =   1290
            Width           =   1400
         End
         Begin VB.TextBox Txtn_EmpLevy 
            Alignment       =   1  'Right Justify
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   2820
            TabIndex        =   28
            Top             =   1305
            Width           =   1400
         End
         Begin VB.TextBox Txtn_ComLevy 
            Alignment       =   1  'Right Justify
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   7680
            TabIndex        =   40
            Top             =   1320
            Width           =   1400
         End
         Begin VB.TextBox Txtn_DayMeal 
            Alignment       =   1  'Right Justify
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   2820
            TabIndex        =   46
            Top             =   2265
            Width           =   1400
         End
         Begin VB.TextBox Txtn_EmpNPF 
            Alignment       =   1  'Right Justify
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   2820
            TabIndex        =   22
            Top             =   690
            Width           =   1400
         End
         Begin VB.TextBox Txtn_ComNPF 
            Alignment       =   1  'Right Justify
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   7680
            TabIndex        =   34
            Top             =   705
            Width           =   1400
         End
         Begin VB.TextBox Txtc_Remarks 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   7680
            MaxLength       =   50
            TabIndex        =   48
            Top             =   2265
            Width           =   4230
         End
         Begin VB.TextBox Txtn_ComMed 
            Alignment       =   1  'Right Justify
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   7680
            TabIndex        =   37
            Top             =   1020
            Width           =   1400
         End
         Begin VB.TextBox Txtn_EmpMed 
            Alignment       =   1  'Right Justify
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   2820
            TabIndex        =   25
            Top             =   1005
            Width           =   1400
         End
         Begin VB.TextBox Txtn_EmpNPFMax 
            Alignment       =   1  'Right Justify
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   5640
            TabIndex        =   24
            Top             =   690
            Width           =   1400
         End
         Begin VB.TextBox Txtn_EmpMedMax 
            Alignment       =   1  'Right Justify
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   5640
            TabIndex        =   27
            Top             =   990
            Width           =   1400
         End
         Begin VB.TextBox Txtn_ComNPFMax 
            Alignment       =   1  'Right Justify
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   10485
            TabIndex        =   36
            Top             =   705
            Width           =   1400
         End
         Begin VB.TextBox Txtn_ComMedMax 
            Alignment       =   1  'Right Justify
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   10485
            TabIndex        =   39
            Top             =   1005
            Width           =   1400
         End
         Begin VB.TextBox Txtn_NightMeal 
            Alignment       =   1  'Right Justify
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   2820
            TabIndex        =   47
            Top             =   2595
            Width           =   1400
         End
         Begin VB.TextBox Txtn_EmpMedMin 
            Alignment       =   1  'Right Justify
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   4230
            TabIndex        =   26
            Top             =   990
            Width           =   1400
         End
         Begin VB.TextBox Txtn_EmpNPFMin 
            Alignment       =   1  'Right Justify
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   4230
            TabIndex        =   23
            Top             =   690
            Width           =   1400
         End
         Begin VB.TextBox Txtn_ComMedMin 
            Alignment       =   1  'Right Justify
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   9075
            TabIndex        =   38
            Top             =   1005
            Width           =   1400
         End
         Begin VB.TextBox Txtn_ComNPFMin 
            Alignment       =   1  'Right Justify
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   9075
            TabIndex        =   35
            Top             =   705
            Width           =   1400
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "EPZ Welfare Fund"
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
            Left            =   1260
            TabIndex        =   79
            Top             =   1620
            Width           =   1425
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Levy"
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
            Left            =   2280
            TabIndex        =   78
            Top             =   1335
            Width           =   390
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Day Meal Allowance"
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
            Left            =   1050
            TabIndex        =   76
            Top             =   2310
            Width           =   1620
         End
         Begin VB.Label Label3 
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
            Left            =   6810
            TabIndex        =   65
            Top             =   2310
            Width           =   765
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "National Pension Fund (NPF)"
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
            Left            =   390
            TabIndex        =   64
            Top             =   735
            Width           =   2280
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Medical Fund"
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
            TabIndex        =   63
            Top             =   1035
            Width           =   1080
         End
         Begin VB.Line Line2 
            BorderColor     =   &H80000005&
            X1              =   15
            X2              =   12105
            Y1              =   2100
            Y2              =   2085
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "E m p l o y e e"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   3690
            TabIndex        =   62
            Top             =   165
            Width           =   1125
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "E m p l o y e r"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   9225
            TabIndex        =   61
            Top             =   165
            Width           =   1095
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "(%)"
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
            Left            =   3360
            TabIndex        =   60
            Top             =   405
            Width           =   255
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "(%)"
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
            Left            =   8175
            TabIndex        =   59
            Top             =   405
            Width           =   255
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "(Max)"
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
            Left            =   6090
            TabIndex        =   58
            Top             =   405
            Width           =   450
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "(Max)"
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
            Left            =   10905
            TabIndex        =   57
            Top             =   405
            Width           =   450
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Night Allowance"
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
            TabIndex        =   56
            Top             =   2640
            Width           =   1320
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "(Min)"
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
            Left            =   4695
            TabIndex        =   55
            Top             =   405
            Width           =   420
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "(Min)"
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
            Left            =   9435
            TabIndex        =   54
            Top             =   405
            Width           =   420
         End
      End
   End
   Begin VB.TextBox Txtc_Company 
      BackColor       =   &H00E0E0E0&
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
      Left            =   705
      MaxLength       =   7
      TabIndex        =   6
      Top             =   1245
      Width           =   1050
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   16020
      Top             =   105
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label7 
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
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   8520
      TabIndex        =   80
      Top             =   1290
      Width           =   960
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   2565
      TabIndex        =   77
      Top             =   1275
      Width           =   780
   End
   Begin VB.Label Label5 
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
      Left            =   135
      TabIndex        =   51
      Top             =   1290
      Width           =   435
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
      Left            =   10890
      TabIndex        =   50
      Top             =   900
      Width           =   540
   End
   Begin VB.Label lbl_scr_name 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      Caption         =   "Company Details"
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
      TabIndex        =   49
      Top             =   915
      Width           =   1380
   End
   Begin VB.Shape shp_scr_name 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   120
      Top             =   855
      Width           =   12270
   End
End
Attribute VB_Name = "frm_Company"
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
    Call Create_Default_Company

    Txtc_Company.Enabled = False
End Sub

Private Sub Form_Activate()
    SSTab1.Tab = 0
    Txtc_DisplayName.SetFocus
End Sub

Private Sub Btn_Exit_Click()
    Unload Me
End Sub

Private Sub Btn_Clear_Click()
    Clear_Controls Me
End Sub

Private Sub Btn_View_Click()
 Dim Search As New Search.MyClass, SerVar
 
    Search.Query = "Select c_displayname Company, c_companyname ShortName, c_company Code from pr_company_mst where c_rec_sta='A'"
    Search.CheckFields = "Company, Code"
    Search.ReturnField = "Company, Code"
    SerVar = Search.Search(, , CON)
    If Len(Search.col1) <> 0 Then
       Txtc_DisplayName = Search.col1
       Txtc_Company = Search.col2
       Call Display_Records
    End If
End Sub

Private Sub Btn_Delete_Click()
On Error GoTo ErrDel
    If Trim(Txtc_Company) = "" Then
       Exit Sub
    End If
    
    If g_Admin Then
       If (MsgBox("Are you sure you want to delete ?", vbYesNo, "Confirmation") = vbYes) Then
          CON.BeginTrans
          CON.Execute "update pr_company_mst set " & GetDelFlag & " where c_company = '" & Trim(Txtc_Company) & "'"
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

Private Sub Btn_Save_Click()
On Error GoTo ErrSave
     
      If ChkSave = False Then
         Exit Sub
      End If
     
      Screen.MousePointer = vbHourglass
      g_SaveFlagNull = True
      
      CON.BeginTrans
        
        Save_Pr_Company_Mst
      
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

Private Sub Btn_Print_Click()
On Error GoTo Err_Print
  Dim SelFor As String, RepTitle As String
  
   If Trim(Txtc_Company) = "" Then
      Exit Sub
   End If
   
   RepTitle = "Company Information"
   SelFor = "{PR_COMPANY_MST.C_COMPANY}='" & Trim(Txtc_Company) & "'"
   Call Print_Rpt(SelFor, "Pr_Company_List.rpt")
  
   If Trim(RepTitle) <> "" Then
      Mdi_Ta_HrPay.CRY1.Formulas(1) = "ReportHead='" & UCase(Trim(RepTitle)) & "'"
   End If

   Mdi_Ta_HrPay.CRY1.Action = 1
  
  Exit Sub

Err_Print:
    MsgBox "Error while Generating - " + Err.Description, vbInformation, "Information"
End Sub

Private Function ChkSave() As Boolean
  If Trim(Txtc_DisplayName) = "" Then
     MsgBox "Company should not be empty", vbInformation, "Information"
     Txtc_DisplayName.SetFocus
     Exit Function
  ElseIf Trim(Txtc_CompanyName) = "" Then
     MsgBox "Short Name should not be empty", vbInformation, "Information"
     Txtc_CompanyName.SetFocus
     Exit Function
  End If
  ChkSave = True
End Function

Private Sub Save_Pr_Company_Mst()
    Set rs = Nothing
    g_Sql = "Select * from pr_company_mst where c_company = '" & Is_Null(Txtc_Company, False) & "'"
    rs.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
    
    If rs.RecordCount = 0 Then
       rs.AddNew
       Txtc_Company = Start_Generate_New
       rs("d_created").Value = GetDateTime
       rs("c_usr_id").Value = g_UserName
    Else
       rs("d_modified").Value = GetDateTime
       rs("c_musr_id").Value = g_UserName
    End If
    
    rs("c_company") = Is_Null(Txtc_Company, False)
    rs("c_displayname").Value = Is_Null(Txtc_DisplayName, False)
    rs("c_companyname").Value = Is_Null(Txtc_CompanyName, False)
    
    rs("c_add1").Value = Is_Null(Txtc_Ad1, False)
    rs("c_add2").Value = Is_Null(Txtc_Ad2, False)
    rs("c_add3").Value = Is_Null(Txtc_Ad3, False)
    rs("c_country").Value = Is_Null(Txtc_Country, False)
    rs("c_tel").Value = Is_Null(Txtc_Tel, False)
    rs("c_mobile").Value = Is_Null(Txtc_Mobile, False)
    rs("c_email").Value = Is_Null(Txtc_EMail, False)
    rs("c_web").Value = Is_Null(Txtc_Web, False)
    
    rs("c_contact").Value = Is_Null(Txtc_Contact, False)
    rs("c_desig").Value = Is_Null(Txtc_Desig, False)
    
    rs("c_vat").Value = Is_Null(Txtc_VatNo, False)
    rs("c_tan").Value = Is_Null(Txtc_TanNo, False)
    rs("c_brn").Value = Is_Null(Txtc_BrnNo, False)
       
    rs("n_empnpf").Value = Is_Null(Txtn_EmpNPF, True)
    rs("n_empnpfmin").Value = Is_Null(Txtn_EmpNPFMin, True)
    rs("n_empnpfmax").Value = Is_Null(Txtn_EmpNPFMax, True)
    
    rs("n_empmed").Value = Is_Null(Txtn_EmpMed, True)
    rs("n_empmedmin").Value = Is_Null(Txtn_EmpMedMin, True)
    rs("n_empmedmax").Value = Is_Null(Txtn_EmpMedMax, True)
    
    rs("n_emplevy").Value = Is_Null(Txtn_EmpLevy, True)
    rs("n_emplevymin").Value = Is_Null(Txtn_EmpLevyMin, True)
    rs("n_emplevymax").Value = Is_Null(Txtn_EmpLevyMax, True)
    
    rs("n_empepz").Value = Is_Null(Txtn_EmpEPZ, True)
    rs("n_empepzmin").Value = Is_Null(Txtn_EmpEPZMin, True)
    rs("n_empepzmax").Value = Is_Null(Txtn_EmpEPZMax, True)
    
    rs("n_comnpf").Value = Is_Null(Txtn_ComNPF, True)
    rs("n_comnpfmin").Value = Is_Null(Txtn_ComNPFMin, True)
    rs("n_comnpfmax").Value = Is_Null(Txtn_ComNPFMax, True)
    
    rs("n_commed").Value = Is_Null(Txtn_ComMed, True)
    rs("n_commedmin").Value = Is_Null(Txtn_ComMedMin, True)
    rs("n_commedmax").Value = Is_Null(Txtn_ComMedMax, True)
    
    rs("n_comlevy").Value = Is_Null(Txtn_ComLevy, True)
    rs("n_comlevymin").Value = Is_Null(Txtn_ComLevyMin, True)
    rs("n_comlevymax").Value = Is_Null(Txtn_ComLevyMax, True)
    
    rs("n_comepz").Value = Is_Null(Txtn_ComEPZ, True)
    rs("n_comepzmin").Value = Is_Null(Txtn_ComEPZMin, True)
    rs("n_comepzmax").Value = Is_Null(Txtn_ComEPZMax, True)
    
    rs("n_daymeal").Value = Is_Null(Txtn_DayMeal, True)
    rs("n_nightmeal").Value = Is_Null(Txtn_NightMeal, True)
    
    rs("c_remarks").Value = Is_Null(Txtc_Remarks, False)

    rs("c_rec_sta").Value = "A"
    rs.Update
End Sub

Private Function Start_Generate_New() As String
  Dim MaxNo As ADODB.Recordset
  
  g_Sql = "Select max(right(c_company,4)) from pr_company_mst "
  Set MaxNo = CON.Execute(g_Sql)
  Start_Generate_New = "COM" & Format(Is_Null(MaxNo(0).Value, True) + 1, "0000")
  
End Function

Private Sub Display_Records()
On Error GoTo Err_Disp
  Dim rsDisp As New ADODB.Recordset
  Dim i As Integer
    
    Set rsDisp = Nothing
    g_Sql = "select * from pr_company_mst where c_rec_sta='A' and c_company = '" & Trim(Txtc_Company) & "'"
    rsDisp.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If rsDisp.RecordCount > 0 Then
       Txtc_Company = Is_Null(rsDisp("c_company").Value, False)
       Txtc_DisplayName = Is_Null(rsDisp("c_displayname").Value, False)
       Txtc_CompanyName = Is_Null(rsDisp("c_companyname").Value, False)
       
       Txtc_Ad1 = Is_Null(rsDisp("c_add1").Value, False)
       Txtc_Ad2 = Is_Null(rsDisp("c_add2").Value, False)
       Txtc_Ad3 = Is_Null(rsDisp("c_add3").Value, False)
       Txtc_Country = Is_Null(rsDisp("c_country").Value, False)
       Txtc_Tel = Is_Null(rsDisp("c_tel").Value, False)
       Txtc_Mobile = Is_Null(rsDisp("c_mobile").Value, False)
       Txtc_EMail = Is_Null(rsDisp("c_email").Value, False)
       Txtc_Web = Is_Null(rsDisp("c_web").Value, False)
       
       Txtc_Contact = Is_Null(rsDisp("c_contact").Value, False)
       Txtc_Desig = Is_Null(rsDisp("c_desig").Value, False)
       
       Txtc_VatNo = Is_Null(rsDisp("c_vat").Value, False)
       Txtc_TanNo = Is_Null(rsDisp("c_tan").Value, False)
       Txtc_BrnNo = Is_Null(rsDisp("c_brn").Value, False)
       
       Txtn_EmpNPF = Format_Num(Is_Null(rsDisp("n_empnpf").Value, True))
       Txtn_EmpNPFMin = Format_Num(Is_Null(rsDisp("n_empnpfmin").Value, True))
       Txtn_EmpNPFMax = Format_Num(Is_Null(rsDisp("n_empnpfmax").Value, True))
       
       Txtn_EmpMed = Format_Num(Is_Null(rsDisp("n_empmed").Value, True))
       Txtn_EmpMedMin = Format_Num(Is_Null(rsDisp("n_empmedmin").Value, True))
       Txtn_EmpMedMax = Format_Num(Is_Null(rsDisp("n_empmedmax").Value, True))
       
       Txtn_EmpLevy = Format_Num(Is_Null(rsDisp("n_emplevy").Value, True))
       Txtn_EmpLevyMin = Format_Num(Is_Null(rsDisp("n_emplevymin").Value, True))
       Txtn_EmpLevyMax = Format_Num(Is_Null(rsDisp("n_emplevymax").Value, True))
       
       Txtn_EmpEPZ = Format_Num(Is_Null(rsDisp("n_empepz").Value, True))
       Txtn_EmpEPZMin = Format_Num(Is_Null(rsDisp("n_empepzmin").Value, True))
       Txtn_EmpEPZMax = Format_Num(Is_Null(rsDisp("n_empepzmax").Value, True))
       
       Txtn_ComNPF = Format_Num(Is_Null(rsDisp("n_comnpf").Value, True))
       Txtn_ComNPFMin = Format_Num(Is_Null(rsDisp("n_comnpfmin").Value, True))
       Txtn_ComNPFMax = Format_Num(Is_Null(rsDisp("n_comnpfmax").Value, True))
       
       Txtn_ComMed = Format_Num(Is_Null(rsDisp("n_commed").Value, True))
       Txtn_ComMedMin = Format_Num(Is_Null(rsDisp("n_commedmin").Value, True))
       Txtn_ComMedMax = Format_Num(Is_Null(rsDisp("n_commedmax").Value, True))
       
       Txtn_ComLevy = Format_Num(Is_Null(rsDisp("n_comlevy").Value, True))
       Txtn_ComLevyMin = Format_Num(Is_Null(rsDisp("n_comlevymin").Value, True))
       Txtn_ComLevyMax = Format_Num(Is_Null(rsDisp("n_comlevymax").Value, True))
       
       Txtn_ComEPZ = Format_Num(Is_Null(rsDisp("n_comepz").Value, True))
       Txtn_ComEPZMin = Format_Num(Is_Null(rsDisp("n_comepzmin").Value, True))
       Txtn_ComEPZMax = Format_Num(Is_Null(rsDisp("n_comepzmax").Value, True))
       
       Txtn_DayMeal = Format_Num(Is_Null(rsDisp("n_daymeal").Value, True))
       Txtn_NightMeal = Format_Num(Is_Null(rsDisp("n_nightmeal").Value, True))
       
       Txtc_Remarks = Is_Null(rsDisp("c_remarks").Value, False)
    End If
  Exit Sub

Err_Disp:
  MsgBox "Error while Display - " + Err.Description
End Sub

Private Sub Txtn_ComEPZ_KeyPress(KeyAscii As Integer)
  Call OnlyNumeric(Txtn_EmpNPF, KeyAscii, 3, 2)
End Sub

Private Sub Txtn_ComEPZMax_KeyPress(KeyAscii As Integer)
  Call OnlyNumeric(Txtn_EmpNPF, KeyAscii, 3, 2)
End Sub

Private Sub Txtn_ComEPZMin_KeyPress(KeyAscii As Integer)
  Call OnlyNumeric(Txtn_EmpNPF, KeyAscii, 3, 2)
End Sub

Private Sub Txtn_ComLevy_KeyPress(KeyAscii As Integer)
  Call OnlyNumeric(Txtn_EmpNPF, KeyAscii, 3, 2)
End Sub

Private Sub Txtn_ComLevyMax_KeyPress(KeyAscii As Integer)
  Call OnlyNumeric(Txtn_EmpNPF, KeyAscii, 3, 2)
End Sub

Private Sub Txtn_ComLevyMin_KeyPress(KeyAscii As Integer)
  Call OnlyNumeric(Txtn_EmpNPF, KeyAscii, 3, 2)
End Sub

Private Sub Txtn_EmpEPZ_KeyPress(KeyAscii As Integer)
  Call OnlyNumeric(Txtn_EmpNPF, KeyAscii, 3, 2)
End Sub

Private Sub Txtn_EmpEPZMax_KeyPress(KeyAscii As Integer)
  Call OnlyNumeric(Txtn_EmpNPF, KeyAscii, 3, 2)
End Sub

Private Sub Txtn_EmpEPZMin_KeyPress(KeyAscii As Integer)
  Call OnlyNumeric(Txtn_EmpNPF, KeyAscii, 3, 2)
End Sub

Private Sub Txtn_EmpLevy_KeyPress(KeyAscii As Integer)
  Call OnlyNumeric(Txtn_EmpNPF, KeyAscii, 3, 2)
End Sub

Private Sub Txtn_EmpLevyMax_KeyPress(KeyAscii As Integer)
  Call OnlyNumeric(Txtn_EmpNPF, KeyAscii, 3, 2)
End Sub

Private Sub Txtn_EmpLevyMin_KeyPress(KeyAscii As Integer)
  Call OnlyNumeric(Txtn_EmpNPF, KeyAscii, 3, 2)
End Sub

Private Sub Txtn_EmpNPF_KeyPress(KeyAscii As Integer)
  Call OnlyNumeric(Txtn_EmpNPF, KeyAscii, 3, 2)
End Sub

Private Sub Txtn_ComNPF_KeyPress(KeyAscii As Integer)
  Call OnlyNumeric(Txtn_ComNPF, KeyAscii, 3, 2)
End Sub

Private Sub Txtn_EmpMed_KeyPress(KeyAscii As Integer)
  Call OnlyNumeric(Txtn_EmpMed, KeyAscii, 3, 2)
End Sub

Private Sub Txtn_ComMed_KeyPress(KeyAscii As Integer)
  Call OnlyNumeric(Txtn_ComMed, KeyAscii, 3, 2)
End Sub

Private Sub Txtn_EmpNPFMin_KeyPress(KeyAscii As Integer)
  Call OnlyNumeric(Txtn_EmpNPFMin, KeyAscii, 8, 2)
End Sub

Private Sub Txtn_ComNPFMin_KeyPress(KeyAscii As Integer)
  Call OnlyNumeric(Txtn_ComNPFMin, KeyAscii, 8, 2)
End Sub

Private Sub Txtn_EmpMedMin_KeyPress(KeyAscii As Integer)
  Call OnlyNumeric(Txtn_EmpMedMin, KeyAscii, 8, 2)
End Sub

Private Sub Txtn_ComMedMin_KeyPress(KeyAscii As Integer)
  Call OnlyNumeric(Txtn_ComMedMin, KeyAscii, 8, 2)
End Sub

Private Sub Txtn_EmpNPFMax_KeyPress(KeyAscii As Integer)
  Call OnlyNumeric(Txtn_EmpNPFMax, KeyAscii, 8, 2)
End Sub

Private Sub Txtn_ComNPFMax_KeyPress(KeyAscii As Integer)
  Call OnlyNumeric(Txtn_ComNPFMax, KeyAscii, 8, 2)
End Sub

Private Sub Txtn_EmpMedMax_KeyPress(KeyAscii As Integer)
  Call OnlyNumeric(Txtn_EmpMedMax, KeyAscii, 8, 2)
End Sub

Private Sub Txtn_ComMedMax_KeyPress(KeyAscii As Integer)
  Call OnlyNumeric(Txtn_ComMedMax, KeyAscii, 8, 2)
End Sub

Private Sub Txtn_DayMeal_KeyPress(KeyAscii As Integer)
  Call OnlyNumeric(Txtn_DayMeal, KeyAscii, 8, 2)
End Sub

Private Sub Txtn_NightMeal_KeyPress(KeyAscii As Integer)
  Call OnlyNumeric(Txtn_NightMeal, KeyAscii, 8, 2)
End Sub



