VERSION 5.00
Begin VB.Form frm_About 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   4530
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   7845
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3126.686
   ScaleMode       =   0  'User
   ScaleWidth      =   7366.861
   ShowInTaskbar   =   0   'False
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
      Left            =   6195
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3615
      Width           =   1155
   End
   Begin VB.CommandButton Cmd_Filter 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   0
      Picture         =   "frm_About.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Filter"
      Top             =   0
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   14.086
      X2              =   7141.489
      Y1              =   2236.305
      Y2              =   2236.305
   End
   Begin VB.Label Lbl_Lincence 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   1470
      Left            =   1545
      TabIndex        =   4
      Top             =   1545
      Width           =   4830
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      Caption         =   "Solutions Knitted for Business Ltd."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1170
      TabIndex        =   0
      Top             =   270
      Width           =   6045
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Time  and  Attendance  System"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1140
      TabIndex        =   2
      Top             =   675
      Width           =   6075
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version (1.0)"
      Height          =   225
      Left            =   3195
      TabIndex        =   3
      Top             =   1140
      Width           =   1005
   End
   Begin VB.Label lblDisclaimer 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   195
      TabIndex        =   1
      Top             =   3420
      Width           =   5430
   End
End
Attribute VB_Name = "frm_About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Btn_Ok_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version 1.0"
    Lbl_Lincence.Caption = "This product is licenced to " & vbCrLf & vbCrLf & _
                           "Citrine Solutions " & vbCrLf & "Chennai, India" & vbCrLf & "citrinesolutions@hotmail.com" & vbCrLf & vbCrLf & _
                           "on " & g_Server
    
    lblDisclaimer.Caption = "Warning : This Computer Program is Protected by Copyright law and international treaties. Unauthorised reproduction or any portion of it, may result in severe civil and criminal penalities, and will be prosecuted to the maximum extend possible under law."
End Sub


