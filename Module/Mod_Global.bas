Attribute VB_Name = "Mod_Global"
Option Explicit

Global CON As New ADODB.Connection

Global Const g_ClientCode = "SIFB"
Global Const g_Database = "HITHRPAY_SIFB"
Global Const g_DatabaseUser = "HITHRPAY"
Global Const g_DatabasePwd = "HITHRPAY"
Global Const g_OBDCUser = "HITHRPAY"

Global Const g_Module = "THP"
Global Const g_SalaryScr = "SAL"
Global Const g_AttnScr = "ATTN"
Global Const g_BranchDefault = "Port-Louis"
Global Const g_DeptDefault = "Admin"

Global g_Server As String
Global g_ReportPath As String
Global g_UserName As String
Global g_CurrentDate As Date
Global g_SaveFlagNull As Boolean
Global g_Sql As String

Global g_Author As Boolean
Global g_Admin As Boolean
Global g_FrmSupUser As Boolean
Global g_FrmAddRight As Boolean
Global g_FrmModRight As Boolean
Global g_FrmDelRight As Boolean
Global g_FrmViewRight As Boolean

Global g_FixedRateName2 As String
Global g_FixedRateName3 As String
Global g_FixedRateName4 As String
Global g_FixedRateName5 As String


Public Const g_Nformat = "###,###,###.00"
Public Const g_Nformat0 = "###,###,###"
Public Const g_Nformat3 = "###,###,###.000"

'Open
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal Hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
'Get Value
Private Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal Hkey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
'Close
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long

Const HKEY_CURRENT_USER = &H80000001
Const ERROR_SUCCESS = 0


Public Sub StartDB()
 On Error GoTo ErrTrap
 
    If CON.State = 1 Then CON.Close
    
    CON.CursorLocation = adUseClient
    
    g_Sql = "DSN=" & g_OBDCUser & "; UID=" & g_DatabaseUser & "; PWD=" & g_DatabasePwd & ";"
    
    CON.Open g_Sql
    
    g_Server = GetKeyValueInOBDC("Server")
    
 Exit Sub

ErrTrap:
     Err.Raise Err.Number & Space(5) & Err.Description
End Sub

Sub Print_Rpt(SelFormula As String, RepName As String, Optional Str As String)
On Error GoTo Errhand
    
    With Mdi_Ta_HrPay.CRY1
        Call Clear_RPT_Formula
        .Formulas(0) = "UserName='" & Proper(g_UserName) & "'"
        .Connect = "dsn=" & g_OBDCUser & "; UID=" & g_DatabaseUser & "; pwd=" & g_DatabasePwd & ";"
        .ReportFileName = g_ReportPath & RepName
        .SelectionFormula = SelFormula
    End With
    
 Exit Sub

Errhand:
    MsgBox "Error" & vbLf & "Cause : " & Err.Description, vbInformation, "Printing Error"
End Sub

Sub Clear_RPT_Formula()
  Dim i As Integer
    
    With Mdi_Ta_HrPay.CRY1
         For i = 0 To 25
            .Formulas(i) = ""
         Next i
        .SelectionFormula = ""
        .GroupSelectionFormula = ""
    End With
End Sub

Public Sub OnlyAlpha(KeyCode As Integer, Optional Caps As Variant)
    If (KeyCode = 8 Or KeyCode = 32) Then Exit Sub
    If (KeyCode >= 97 And KeyCode <= 122) Or (KeyCode >= 65 And KeyCode <= 90) Then
            If IsMissing(Caps) Then Exit Sub
            Select Case Caps
                    Case vbUpperCase
                            If (KeyCode >= 97 And KeyCode <= 122) Then KeyCode = KeyCode - 32
                    Case vbLowerCase
                            If (KeyCode >= 65 And KeyCode <= 90) Then KeyCode = KeyCode + 32
            End Select
    Else
            Beep
            KeyCode = 0
    End If
End Sub

Public Sub OnlyNumeric(TextBox As TextBox, KeyCode As Integer, ByVal IntDigits As Integer, Optional DecimalDigits As Variant)
    Dim blnWithPeriod As Boolean
    If Not ((KeyCode >= 48 And KeyCode <= 57) Or KeyCode = 8) Then
            blnWithPeriod = IIf(IsMissing(DecimalDigits), False, True)
            If blnWithPeriod Then
                    Dim blnHasPeriod As Boolean
                    If Not (KeyCode = Asc(".")) Then
                            Call Beep:      KeyCode = 0
                    Else
                            blnHasPeriod = IIf(InStr(1, TextBox.Text, ".") > 0, True, False)
                            If blnHasPeriod Then
                                    Call Beep:      KeyCode = 0
                            End If
                    End If
            Else
                    Call Beep:      KeyCode = 0
            End If
    ElseIf (KeyCode >= 48 And KeyCode <= 57) Then
            Dim iPosDecimal As Integer
            iPosDecimal = InStr(1, TextBox.Text, ".")
            If iPosDecimal > 0 Then
                    If TextBox.SelStart < iPosDecimal Then
                            If Len(Mid(TextBox.Text, 1, iPosDecimal - 1)) >= IntDigits Then
                                    Call Beep:      KeyCode = 0
                            End If
                    ElseIf TextBox.SelStart > iPosDecimal Then
                            If Len(Mid(TextBox.Text, iPosDecimal + 1)) >= DecimalDigits Then
                                    Call Beep:      KeyCode = 0
                            End If
                    End If
            Else
                    If Len(TextBox.Text) >= IntDigits Then
                            Call Beep:      KeyCode = 0
                    End If
            End If
    End If
End Sub

Public Sub OnlyAlphaNumeric(KeyCode As Integer, Optional Caps As Integer)
    If (KeyCode = 8 Or KeyCode = 32) Then Exit Sub
    If (KeyCode >= 48 And KeyCode <= 57) Then Exit Sub
    If (KeyCode >= 65 And KeyCode <= 90) Then
            If IsMissing(Caps) Then Exit Sub
            If Caps = vbLowerCase Then KeyCode = KeyCode + 32
            Exit Sub
    End If
    If (KeyCode >= 97 And KeyCode <= 122) Then
            If IsMissing(Caps) Then Exit Sub
            If Caps = vbUpperCase Then KeyCode = KeyCode - 32
            Exit Sub
    End If
    KeyCode = 0
End Sub

Function Is_Null(arg As Variant, argNum As Boolean) As Variant
    If Not IsNull(arg) Then
         If argNum = True Then 'Number
            Is_Null = Val(arg)
         Else                  'String
            If Trim(arg) = "" And g_SaveFlagNull Then
               Is_Null = Null
            Else
               Is_Null = Trim(arg)
            End If
         End If
    Else
         If argNum = True Then 'Number
            Is_Null = 0
         Else                  'String
            If g_SaveFlagNull Then
               Is_Null = Null
            Else
               Is_Null = ""
            End If
         End If
    End If
End Function

Function Is_Null_D(arg As Variant, argNum As Boolean, Optional dec As Boolean) As Variant
    Dim RStr As String
    If Not IsNull(arg) Then
         If argNum = True Then 'Number
            If Trim(arg) <> "" Then
               If dec = True Then
                  If Right(Trim(arg), 1) = 5 Then
                     If InStr(1, Trim(arg), ".") > 0 And Len(Trim(arg)) - InStr(1, Trim(arg), ".") > 2 Then
                        arg = Left(Trim(arg), Len(Trim(arg)) - 1) & "6"
                     End If
                  End If
                  Is_Null_D = Round(CDbl(arg), 2)
               Else
                  Is_Null_D = CDbl(arg)
               End If
            End If
         Else                  'String
            If Trim(arg) = "" And g_SaveFlagNull Then
                Is_Null_D = Null
            Else
               Is_Null_D = Trim(arg)
            End If
         End If
    Else
         If argNum = True Then 'Number
            Is_Null_D = 0
         Else                  'String
            If g_SaveFlagNull Then
               Is_Null_D = Null
            Else
               Is_Null_D = ""
            End If
         End If
    End If
End Function

Function Is_Date(arg As Variant, ByVal vSaveDispType As String) As Variant
    If UCase(vSaveDispType) = "S" Then
       Is_Date = IIf(IsDate(arg), Format(arg, "yyyy-mm-dd"), Null)
    Else
       Is_Date = IIf(IsDate(arg), Format(arg, "dd/mm/yyyy"), "__/__/____")
    End If
End Function

Function Is_DateTime(arg As Variant) As Variant
    Is_DateTime = IIf(IsDate(arg), Format(arg, "dd/mm/yyyy   Hh:Nn:Ss"), "")
End Function

Function Format_Num(arg As Variant) As Variant
    If Is_Null_D(arg, True) <> 0 Then
       Format_Num = Format(arg, g_Nformat)
    Else
       Format_Num = ""
    End If
End Function

Public Function Proper(ByVal vStrValue As String)
   Dim i As Integer
   Dim ArrStr
 
    ArrStr = Split(vStrValue, " ")
    For i = 0 To UBound(ArrStr)
        If Trim(ArrStr(i)) <> "" Then
           If i = 0 Then
              Proper = UCase(Left(ArrStr(i), 1)) & LCase(Right(ArrStr(i), Len(ArrStr(i)) - 1))
           Else
              Proper = Proper & " " & UCase(Left(ArrStr(i), 1)) & LCase(Right(ArrStr(i), Len(ArrStr(i)) - 1))
           End If
        End If
    Next i
End Function

Public Function SubString(ByVal vStrValue As String, vStart As Long, vChar As Long)
   Dim tmpStr As String
    
    If vStart <= 0 Or vChar <= 0 Then
       SubString = ""
    ElseIf vStart = 1 Then
       SubString = Left(vStrValue, vChar)
    Else
       tmpStr = Right(vStrValue, Len(vStrValue) - (vStart - 1))
       SubString = Left(tmpStr, vChar)
    End If
End Function

Public Sub TGControlProperty(ByVal frm As Form, Optional SpreadProperty As String)
On Error Resume Next
Dim Ctl, i

    For Each Ctl In frm.Controls
        If TypeOf Ctl Is Label Then
           Ctl.FontName = "Arial"
           Ctl.FontSize = 8
           Ctl.FontBold = True
        
        ElseIf TypeOf Ctl Is TextBox Then
           Ctl.FontName = "MS Sans Serif"
           Ctl.FontSize = 7
           Ctl.FontBold = False
           If Ctl.Enabled = False Then
              Ctl.BackColor = &HE0E0E0
           End If
           
        ElseIf TypeOf Ctl Is ComboBox Then
           Ctl.FontName = "MS Sans Serif"
           Ctl.FontSize = 8
           Ctl.FontBold = False
        
        ElseIf TypeOf Ctl Is SSTab Then
           Ctl.ForeColor = RGB(0, 0, 175)
           Ctl.FontName = "Arial"
           Ctl.FontSize = 8
           Ctl.FontBold = False
        
        ElseIf TypeOf Ctl Is Line Then
           Ctl.X1 = 10
           Ctl.BorderColor = RGB(255, 255, 255)
           
        ElseIf TypeOf Ctl Is fpSpread And (SpreadProperty = "" Or SpreadProperty = "Y") Then
           Ctl.Row = -1
           Ctl.Col = -1
               Ctl.FontName = "MS Sans Serif"
               Ctl.FontSize = 8
               Ctl.FontBold = False
           
           ' Spread Details
           Ctl.Row = -1
           Ctl.Col = -1
               Ctl.FontName = "Arial"
               Ctl.FontSize = 8
               Ctl.FontBold = False
               
           
           ' Header
           Ctl.Row = 0
           Ctl.Col = -1
               Ctl.FontName = "Arial"
               Ctl.FontSize = 8
               Ctl.FontBold = False
               Ctl.ShadowColor = RGB(255, 255, 235)
               Ctl.ShadowText = RGB(0, 0, 0)
           
           ' Grid details
           For i = 1 To Ctl.MaxCols
               Ctl.Row = -1
               Ctl.Col = i
                   Ctl.CellBorderType = SS_BORDER_TYPE_OUTLINE
                   Ctl.CellBorderStyle = SS_BORDER_STYLE_SOLID
                   If Ctl.Lock = True Then
                      Ctl.BackColor = RGB(230, 230, 230)
                      Ctl.CellBorderColor = RGB(192, 192, 192)
                   Else
                      Ctl.BackColor = RGB(255, 255, 255)
                      Ctl.CellBorderColor = RGB(192, 192, 192)
                   End If
                   Ctl.Action = SS_ACTION_SET_CELL_BORDER
           Next i
        End If
    Next
End Sub

Public Sub LookUp_Employee(frm As Form, Optional ByVal EmpType As String, Optional ByVal vBranch As String, Optional ByVal vDept As String)
  Dim Search As New Search.MyClass, SerVar
  Dim tmpStr As String
  
    If Trim(EmpType) <> "" Then
       tmpStr = "type = '" & Trim(EmpType) & "'"
    End If
    If Trim(vBranch) <> "" Then
       If Trim(tmpStr) = "" Then
          tmpStr = "branch = '" & Trim(vBranch) & "'"
       Else
          tmpStr = Trim(tmpStr) & " and branch = '" & Trim(vBranch) & "'"
       End If
    End If
    If Trim(vDept) <> "" Then
       If Trim(tmpStr) = "" Then
          tmpStr = "department = '" & Trim(vDept) & "'"
       Else
          tmpStr = Trim(tmpStr) & " and department = '" & Trim(vDept) & "'"
       End If
    End If
    
    If Trim(tmpStr) = "" Then
       Search.Query = "select * from v_pr_emp_lookup_dtl_ALL "
    Else
       Search.Query = "select * from v_pr_emp_lookup_dtl_ALL where " & Trim(tmpStr)
    End If
    Search.CheckFields = "EmployeeNo, Name"
    Search.ReturnField = "EmployeeNo, Name"
    SerVar = Search.Search(, , CON)
    If Len(Search.col1) <> 0 Then
        frm.Txtc_EmployeeName = Search.col2 & Space(100) & Search.col1
    End If
End Sub

Public Function ReportFilterOption(ByVal vF1 As String, ByVal vF2 As String, Optional ByVal vF3 As String, Optional ByVal vF4 As String, Optional ByVal vF5 As String)
  Dim vStrFor As String
     
      If Trim(vF1) <> "" Then
         vStrFor = Trim(vF1)
      End If
      If Trim(vF2) <> "" Then
         If Trim(vStrFor) <> "" Then
            vStrFor = Trim(vStrFor) & " AND " & Trim(vF2)
         Else
            vStrFor = Trim(vF2)
         End If
      End If
      If Trim(vF3) <> "" Then
         If Trim(vStrFor) <> "" Then
            vStrFor = Trim(vStrFor) & " AND " & Trim(vF3)
         Else
            vStrFor = Trim(vF3)
         End If
      End If
      If Trim(vF4) <> "" Then
         If Trim(vStrFor) <> "" Then
            vStrFor = Trim(vStrFor) & " AND " & Trim(vF4)
         Else
            vStrFor = Trim(vF4)
         End If
      End If
      If Trim(vF5) <> "" Then
         If Trim(vStrFor) <> "" Then
            vStrFor = Trim(vStrFor) & " AND " & Trim(vF5)
         Else
            vStrFor = Trim(vF5)
         End If
      End If
    
      ReportFilterOption = Trim(vStrFor)
End Function

Public Function Ceiling(ByVal vValue As Double)
    If vValue - Round(vValue, 0) > 0 Then
       Ceiling = Round(vValue) + 1
    Else
       Ceiling = Round(vValue, 0)
    End If
End Function

Public Function Floor(ByVal vValue As Double, Optional vDigit As Integer)
   If vDigit = 0 Then
      If vValue - Round(vValue, 0) < 0 Then
         Floor = Round(vValue) - 1
      Else
         Floor = Round(vValue, 0)
      End If
   Else
      If vValue - Round(vValue, vDigit) < 0 Then
         If vDigit = 1 Then
            Floor = Round(vValue, vDigit) - 0.1
         ElseIf vDigit = 2 Then
            Floor = Round(vValue, vDigit) - 0.01
         ElseIf vDigit = 3 Then
            Floor = Round(vValue, vDigit) - 0.001
         ElseIf vDigit = 4 Then
            Floor = Round(vValue, vDigit) - 0.0001
         ElseIf vDigit = 5 Then
            Floor = Round(vValue, vDigit) - 0.00001
         Else
            Floor = Round(vValue, vDigit) - 0.00001
         End If
      Else
         Floor = Round(vValue, vDigit)
      End If
   End If
End Function

Public Function ChkUsrEditDelRight(ByVal vUserName As String, ByVal vTableName As String, ByVal vFieldName As String, ByVal vFieldValue As String) As Boolean
  Dim rsChk As New ADODB.Recordset
  
    If Not g_Admin Then
       Set rsChk = Nothing
       g_Sql = "select b.c_usr_id, a.cadmin from users_mst a, " & Trim(vTableName) & " b " & _
               "where a.cUserName = b.c_usr_id and b." & Trim(vFieldName) & " = '" & Trim(vFieldValue) & "'"
       rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
       If rsChk.RecordCount > 0 Then
          If g_FrmSupUser Then
             If Trim(rsChk("c_usr_id").Value) <> Trim(vUserName) Then
                MsgBox "The informations entered / Modified by (" & Proper(Trim(rsChk("c_usr_id").Value)) & ") " & vbCrLf & _
                       "Please make sure your entries. ", vbInformation, "Information"
             End If
          Else
             If Trim(rsChk("c_usr_id").Value) <> Trim(vUserName) Then
                MsgBox "The informations entered (" & Proper(Trim(rsChk("c_usr_id").Value)) & "), who can Edit / Delete. " & vbCrLf & _
                       "No access for others. Contact Supervisor. ", vbInformation, "Information"
                ChkUsrEditDelRight = False
                Exit Function
             End If
          End If
       End If
    End If
    ChkUsrEditDelRight = True
End Function

Public Function ChkUsrRight() As Boolean
    If g_Admin Or g_Author Then
    Else
        MsgBox "No Access to Use this Option. Please Contact Admin", vbInformation, "Information"
        Exit Function
    End If
    ChkUsrRight = True
End Function

Public Function ChkScreenRight(ByVal vScreenId As String) As Boolean
  Dim rsChk As New ADODB.Recordset
  
    If g_Admin Or g_Author Then
       ChkScreenRight = True
       Exit Function
    End If
    
    Set rsChk = Nothing
    g_Sql = "select * from pr_user_dtl where c_user_id = '" & Trim(g_UserName) & "' and c_screen_id = '" & Trim(vScreenId) & "'"
    rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If rsChk.RecordCount = 0 Then
        MsgBox "No Access to Use this Option. Please Contact Admin", vbInformation, "Information"
        Exit Function
    Else
        g_FrmSupUser = Is_Null(rsChk("c_adminright").Value, True)
        g_FrmAddRight = Is_Null(rsChk("c_addright").Value, True)
        g_FrmDelRight = Is_Null(rsChk("c_delright").Value, True)
        g_FrmViewRight = Is_Null(rsChk("c_viewright").Value, True)
        
        If g_FrmSupUser Or g_FrmAddRight Or g_FrmDelRight Or g_FrmViewRight Then
        Else
           MsgBox "No Access to Use this Option. Please Contact Admin", vbInformation, "Information"
           Exit Function
        End If
    End If
    ChkScreenRight = True
End Function

Public Sub ScreenUserRight(frm As Form)
    If g_Admin Then
       Exit Sub
    ElseIf g_FrmSupUser Then
       Exit Sub
    End If
    
    frm.Btn_Save.Visible = False
    frm.Btn_View.Visible = False
    frm.Btn_Delete.Visible = False
    
    If g_FrmAddRight Then
       frm.Btn_Save.Visible = True
       frm.Btn_View.Visible = True
    End If
    If g_FrmDelRight Then
       frm.Btn_Delete.Visible = True
       frm.Btn_View.Visible = True
    End If
    If g_FrmViewRight Then
       frm.Btn_View.Visible = True
    End If
End Sub

Public Function ChkShiftChangeRight(Optional vMsgSuppress As Boolean) As Boolean
  Dim rsChk As New ADODB.Recordset
  
   If Not g_Admin Then
      Set rsChk = Nothing
      g_Sql = "select cshiftchange from users_mst where cusername = '" & g_UserName & "'"
      rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
      If rsChk.RecordCount > 0 Then
         If Is_Null(rsChk("cshiftchange").Value, False) <> "Y" Then
            If Not vMsgSuppress Then
               MsgBox "No Access to Use this Option. Please Contact Supervisor", vbInformation, "Information"
            End If
            ChkShiftChangeRight = False
            Exit Function
         End If
      Else
         ChkShiftChangeRight = False
         Exit Function
      End If
   End If
   ChkShiftChangeRight = True
End Function

Public Function TimeToMins(ByVal DTime As Double)
  Dim dHrs As Double, dMins As Double, dTmp As Double
      dMins = (DTime * 100) Mod 100
      If Val(dMins) > 0 Then
         dTmp = Val(dMins / 100)
      End If
      dHrs = Val(DTime - dTmp)
      TimeToMins = Val(dHrs) * 60 + Val(dMins)
End Function

Public Function MinsToTime(ByVal dTotalMins As Double)
   Dim dHrs As Double, dMins As Double
       dHrs = dTotalMins - (dTotalMins Mod 60)
       dHrs = Val(dHrs / 60)
       
       dMins = (dTotalMins Mod 60) / 100
       MinsToTime = Val(dHrs) + Val(dMins)
End Function

Public Function MinsToDecimal(ByVal DTime As Double)
   Dim dMins As Double
       dMins = (DTime * 100) Mod 100
       If dMins = 0 Then
          MinsToDecimal = DTime
       Else
          dMins = Round(dMins / 60, 2)
          MinsToDecimal = (DTime - (((DTime * 100) Mod 100) / 100)) + dMins
       End If
End Function

Public Function DecimalToMins(ByVal DTime As Double)
   Dim dMins As Double
       dMins = (DTime * 100) Mod 100
       If dMins = 0 Then
          DecimalToMins = DTime
       Else
          dMins = Round((dMins * 60) / 100, 0)
          DecimalToMins = (DTime - (((DTime * 100) Mod 100) / 100)) + (dMins / 100)
       End If
End Function

Public Function FloorTo15Mins(ByVal vValue As Double)
    FloorTo15Mins = vValue - (vValue Mod 15)
End Function

Public Function GetDelFlag(Optional ByVal vDelRemarks As String)
   If Trim(vDelRemarks) = "" Then
      GetDelFlag = " c_rec_sta = 'I', c_dusr_id = '" & Trim(g_UserName) & "',  d_deleted='" & GetDateTime & "'"
   Else
      GetDelFlag = " c_rec_sta = 'I', c_dusr_id = '" & Trim(g_UserName) & "',  d_deleted='" & GetDateTime & "', c_delremarks = '" & Left(Trim(vDelRemarks), 100) & "' "
   End If
End Function

Public Sub Enable_Controls(ByVal frm As Form, ByVal tag As Boolean)
On Error Resume Next
Dim Ctl
    For Each Ctl In frm.Controls
        If Not (TypeOf Ctl Is Label Or TypeOf Ctl Is Frame Or TypeOf Ctl Is SSTab) Then Ctl.Enabled = tag
    Next

    frm.Btn_Exit.Enabled = True
End Sub

Public Sub Clear_Controls(ByVal frm As Form)
On Error Resume Next
Dim Ctl
    For Each Ctl In frm.Controls
        If (TypeOf Ctl Is TextBox) Then
            Ctl.Text = Empty
        ElseIf TypeOf Ctl Is MaskEdBox Then
            Ctl.Text = ""

        ElseIf TypeOf Ctl Is DateControl Then
            Ctl.Text = "__/__/____"
        ElseIf TypeOf Ctl Is ComboBox Then
            If Ctl.Style = 2 Then
                Ctl.ListIndex = -1
            Else
                Ctl.Text = ""
            End If
        ElseIf TypeOf Ctl Is ImageCombo Then
            Ctl.SelectedItem = Nothing
            Ctl.Text = ""
        ElseIf TypeOf Ctl Is CheckBox Then
            Ctl.Value = False
        End If
    Next
End Sub

Public Sub Enable_Buttons(ByVal frm As Form, btns)
 Dim i As Integer
 Dim Button
    For i = 1 To 10
        frm.Controls("tlb1").enablebutton(i) = False
    Next i
    For Each Button In btns
        frm.Controls("tlb1").enablebutton(Button) = True
    Next
End Sub

Public Function GetDateTime()
  Dim rsDt As New ADODB.Recordset
    Set rsDt = Nothing
    g_Sql = "select getdate()"
    rsDt.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If rsDt.RecordCount > 0 Then
       GetDateTime = Format(rsDt(0).Value, "yyyy-mm-dd hh:mm:ss")
    Else
       GetDateTime = Format(Now, "yyyy-mm-dd hh:mm:ss")
    End If
End Function


Public Function New_Val(ByVal s As Variant) As Double
    If IsNull(s) Then
        New_Val = 0
    ElseIf IsNumeric(s) Then
        New_Val = CDbl(s)
    Else
        New_Val = Val(s)
    End If
End Function


Public Function ChkPeriodOpen(ByVal vPeriod As Long, ByVal vEmpType As String) As Boolean
 On Error GoTo Err_Proc
   Dim rsChk As New ADODB.Recordset
    
     Set rsChk = Nothing
     g_Sql = "select * from pr_payperiod_dtl where n_period = " & vPeriod & " and c_type = '" & Trim(vEmpType) & "'"
     rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
     If rsChk.RecordCount > 0 Then
        If Is_Null_D(rsChk("c_period_closed").Value, False) = "Y" Then
           MsgBox "Pay Period is closed. Please contact Admin", vbInformation, "Information"
           ChkPeriodOpen = False
           Exit Function
        End If
     Else
        MsgBox "Period Not Found. Please Check the Pay Period", vbInformation, "Information"
        ChkPeriodOpen = False
        Exit Function
     End If
     ChkPeriodOpen = True
   Exit Function

Err_Proc:
   MsgBox "Critical Error - " + Err.Description, vbCritical, "Critical"
End Function

Public Function ChkPeriodExists(ByVal vPeriod As Long, ByVal vEmpType As String) As Boolean
 On Error GoTo Err_Proc
   Dim rsChk As New ADODB.Recordset
    
     Set rsChk = Nothing
     g_Sql = "select * from pr_payperiod_dtl where n_period = " & vPeriod & " and c_type = '" & Trim(vEmpType) & "'"
     rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
     If rsChk.RecordCount = 0 Then
        MsgBox "Period Not Found. Please Check the Pay Period", vbInformation, "Information"
        ChkPeriodExists = False
        Exit Function
     End If
     ChkPeriodExists = True
   Exit Function

Err_Proc:
   MsgBox "Critical Error - " + Err.Description, vbCritical, "Critical"
End Function


Public Sub LoadComboCompany(frm As Form)
  Dim rsCombo As New ADODB.Recordset
  Dim i As Integer
    
    Set rsCombo = Nothing
    g_Sql = "select c_company, c_companyname from pr_company_mst where c_rec_sta='A' " & _
            "order by c_companyname "
    rsCombo.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    frm.Cmb_Company.Clear
    frm.Cmb_Company.AddItem ""
    For i = 0 To rsCombo.RecordCount - 1
        frm.Cmb_Company.AddItem rsCombo("c_companyname").Value & Space(100) & rsCombo("c_company").Value
        rsCombo.MoveNext
    Next i
End Sub

Public Sub DisplayComboCompany(frm As Form, ByVal vComboVal As String)
  Dim i As Integer
  
    For i = 0 To frm.Cmb_Company.ListCount - 1
      If Right(Trim(frm.Cmb_Company.List(i)), 7) = Trim(vComboVal) Then
         frm.Cmb_Company.ListIndex = i
         Exit For
      End If
    Next i
End Sub

Public Sub LoadComboBranch(frm As Form)
  Dim rsCombo As New ADODB.Recordset
  Dim i As Integer
    
    Set rsCombo = Nothing
    g_Sql = "select distinct c_branch from pr_emp_mst where c_rec_sta = 'A' and c_branch is not null order by c_branch "
    rsCombo.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    frm.Cmb_Branch.Clear
    frm.Cmb_Branch.AddItem ""
    For i = 0 To rsCombo.RecordCount - 1
        If Is_Null(rsCombo("c_branch").Value, False) <> "" Then
           frm.Cmb_Branch.AddItem rsCombo("c_branch").Value
        End If
        rsCombo.MoveNext
    Next i
End Sub

Public Sub DisplayComboBranch(frm As Form, ByVal vComboVal As String)
  Dim i As Integer
    For i = 0 To frm.Cmb_Branch.ListCount - 1
      If Trim(frm.Cmb_Branch.List(i)) = Trim(vComboVal) Then
         frm.Cmb_Branch.ListIndex = i
         Exit For
      End If
    Next i
End Sub

Public Sub LoadComboDept(frm As Form)
  Dim rsCombo As New ADODB.Recordset
  Dim i As Integer
    
    Set rsCombo = Nothing
    g_Sql = "select distinct c_dept from pr_emp_mst where c_rec_sta = 'A' and c_dept is not null order by c_dept "
    rsCombo.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    frm.Cmb_Dept.Clear
    frm.Cmb_Dept.AddItem ""
    For i = 0 To rsCombo.RecordCount - 1
        If Is_Null(rsCombo("c_dept").Value, False) <> "" Then
           frm.Cmb_Dept.AddItem rsCombo("c_dept").Value
        End If
        rsCombo.MoveNext
    Next i
End Sub

Public Sub DisplayComboDept(frm As Form, ByVal vComboVal As String)
  Dim i As Integer
    For i = 0 To frm.Cmb_Dept.ListCount - 1
      If Trim(frm.Cmb_Dept.List(i)) = Trim(vComboVal) Then
         frm.Cmb_Dept.ListIndex = i
         Exit For
      End If
    Next i
End Sub

Public Sub LoadComboDesig(frm As Form)
  Dim rsCombo As New ADODB.Recordset
  Dim i As Integer
    
    Set rsCombo = Nothing
    g_Sql = "select distinct c_desig from pr_emp_mst where c_rec_sta = 'A' and c_desig is not null order by c_desig "
    rsCombo.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    frm.Cmb_Desig.Clear
    frm.Cmb_Desig.AddItem ""
    For i = 0 To rsCombo.RecordCount - 1
        If Is_Null(rsCombo("c_desig").Value, False) <> "" Then
           frm.Cmb_Desig.AddItem rsCombo("c_desig").Value
        End If
        rsCombo.MoveNext
    Next i
End Sub

Public Sub DisplayComboDesig(frm As Form, ByVal vComboVal As String)
  Dim i As Integer
    For i = 0 To frm.Cmb_Desig.ListCount - 1
      If Trim(frm.Cmb_Desig.List(i)) = Trim(vComboVal) Then
         frm.Cmb_Desig.ListIndex = i
         Exit For
      End If
    Next i
End Sub

Public Sub LoadComboEmpType(frm As Form)
  Dim rsCombo As New ADODB.Recordset
  Dim i As Integer
    
    Set rsCombo = Nothing
    g_Sql = "select distinct c_emptype from pr_emp_mst where c_rec_sta = 'A' and c_emptype is not null order by c_emptype "
    rsCombo.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    frm.Cmb_EmpType.Clear
    frm.Cmb_EmpType.AddItem ""
    For i = 0 To rsCombo.RecordCount - 1
        If Is_Null(rsCombo("c_emptype").Value, False) <> "" Then
           frm.Cmb_EmpType.AddItem rsCombo("c_emptype").Value
        End If
        rsCombo.MoveNext
    Next i
End Sub

Public Sub DisplayComboEmpType(frm As Form, ByVal vComboVal As String)
  Dim i As Integer
    For i = 0 To frm.Cmb_EmpType.ListCount - 1
      If Trim(frm.Cmb_EmpType.List(i)) = Trim(vComboVal) Then
         frm.c_emptype.ListIndex = i
         Exit For
      End If
    Next i
End Sub

Public Sub LoadComboShift(frm As Form)
  Dim rsCombo As New ADODB.Recordset
  Dim i As Integer
    
    Set rsCombo = Nothing
    g_Sql = "select c_code, c_shiftname from pr_shiftstructure_mst where c_rec_sta='A' order by c_shiftname "
    rsCombo.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    frm.Cmb_Shift.Clear
    frm.Cmb_Shift.AddItem ""
    For i = 0 To rsCombo.RecordCount - 1
        frm.Cmb_Shift.AddItem rsCombo("c_shiftname").Value & Space(100) & rsCombo("c_code").Value
        rsCombo.MoveNext
    Next i
End Sub

Public Sub LoadComboLeave(frm As Form)
  Dim rsCombo As New ADODB.Recordset
  Dim i As Integer
    
    Set rsCombo = Nothing
    g_Sql = "select c_leave, c_leavename from pr_leave_mst where c_rec_sta='A' order by c_leavename "
    rsCombo.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    frm.Cmb_Leave.Clear
    frm.Cmb_Leave.AddItem ""
    For i = 0 To rsCombo.RecordCount - 1
        frm.Cmb_Leave.AddItem rsCombo("c_leavename").Value & Space(100) & rsCombo("c_leave").Value
        rsCombo.MoveNext
    Next i
End Sub

Public Sub MakeMonthTwoDigits(frm As Form)
    If Val(frm.Txtc_Month) = 0 Then
       frm.Txtc_Month.Text = ""
    ElseIf Len(frm.Txtc_Month) < 2 Then
       frm.Txtc_Month.Text = "0" & Trim(frm.Txtc_Month)
    End If
End Sub

Public Function GetMonthNameInFrench(ByVal vMonth As Integer)
   Dim vMonthName As String
   
    If vMonth = 1 Then
       vMonthName = "janvier"
    ElseIf vMonth = 2 Then
       vMonthName = "fevrier"
    ElseIf vMonth = 3 Then
       vMonthName = "mars"
    ElseIf vMonth = 4 Then
       vMonthName = "avril"
    ElseIf vMonth = 5 Then
       vMonthName = "mai"
    ElseIf vMonth = 6 Then
       vMonthName = "juin"
    ElseIf vMonth = 7 Then
       vMonthName = "juillet"
    ElseIf vMonth = 8 Then
       vMonthName = "aout"
    ElseIf vMonth = 9 Then
       vMonthName = "septembre"
    ElseIf vMonth = 10 Then
       vMonthName = "octobre"
    ElseIf vMonth = 11 Then
       vMonthName = "novembre"
    ElseIf vMonth = 12 Then
       vMonthName = "decembre"
    Else
       vMonthName = ""
    End If
    
    GetMonthNameInFrench = UCase(Trim(vMonthName))
    
End Function

Public Function EmpMaster_RepFilter(frm As Form) As String
  Dim vF1 As String, vF2 As String, vF3 As String, vF4 As String, vF5 As String
  Dim SelFor As String
  
     vF1 = "": vF2 = "": vF3 = "": vF4 = "": vF5 = ""
     SelFor = ""
     
     If Right(Trim(frm.Cmb_Company), 7) <> "" Then
        vF1 = "{V_PR_EMP_MST.C_COMPANY}='" & Right(Trim(frm.Cmb_Company), 7) & "'"
     End If
     
     If Trim(frm.Cmb_Branch) <> "" Then
        vF2 = "{V_PR_EMP_MST.C_BRANCH}='" & Trim(frm.Cmb_Branch) & "'"
     End If
     
     If Trim(frm.Cmb_Dept) <> "" Then
        vF3 = "{V_PR_EMP_MST.C_DEPT}='" & Trim(frm.Cmb_Dept) & "'"
     End If
     
     If Trim(frm.Cmb_Desig) <> "" Then
        vF4 = "{V_PR_EMP_MST.C_DESIG}='" & Trim(frm.Cmb_Desig) & "'"
     End If
     
     If Trim(frm.Cmb_EmpType) <> "" Then
        vF5 = "{V_PR_EMP_MST.C_EMPTYPE}='" & Trim(frm.Cmb_EmpType) & "'"
     End If
     
     SelFor = ReportFilterOption(vF1, vF2, vF3, vF4, vF5)
     vF1 = "": vF2 = "": vF3 = "": vF4 = "": vF5 = ""
     
     If Trim(frm.Txtc_EmployeeName) <> "" Then
        vF1 = "{V_PR_EMP_MST.C_EMPNO}='" & Trim(Right(Trim(frm.Txtc_EmployeeName), 7)) & "'"
     End If
     
     If Trim(frm.Cmb_Sex) <> "" Then
        vF2 = "{V_PR_EMP_MST.C_SEX}='" & Trim(frm.Cmb_Sex) & "'"
     End If
     
     If Trim(frm.Cmb_Nationality) <> "" Then
        vF3 = "{V_PR_EMP_MST.C_NATIONALITY}='" & Trim(frm.Cmb_Nationality) & "'"
     End If
     
     If Trim(frm.Cmb_DayWork) <> "" Then
        vF4 = "{V_PR_EMP_MST.C_DAYWORK}='" & Trim(Right(Trim(frm.Cmb_DayWork), 5)) & "'"
     End If
     
     SelFor = ReportFilterOption(SelFor, vF1, vF2, vF3, vF4)
     vF1 = "": vF2 = "": vF3 = "": vF4 = "": vF5 = ""
     
     
     If frm.Opt_Active.Value = True Then
        vF1 = "ISNULL({V_PR_EMP_MST.D_DOL})"
     ElseIf frm.Opt_Left.Value = True Then
        vF1 = "NOT ISNULL({V_PR_EMP_MST.D_DOL})"
     End If
     
     If frm.Opt_StaffTypeFlat.Value = True Then
        vF2 = "{V_PR_EMP_MST.C_STAFFTYPE}='Flat'"
     ElseIf frm.Opt_StaffTypeOt.Value = True Then
        vF2 = "{V_PR_EMP_MST.C_STAFFTYPE}='OverTime'"
     End If
     
     If frm.Opt_SalaryTypeMon.Value = True Then
        vF3 = "{V_PR_EMP_MST.C_SALARYTYPE}='Monthly'"
     ElseIf frm.Opt_SalaryTypeHr.Value = True Then
        vF3 = "{V_PR_EMP_MST.C_SALARYTYPE}='Hourly'"
     End If
     
     If frm.Opt_PayTypeCash.Value = True Then
        vF4 = "{V_PR_EMP_MST.C_PAYTYPE}='CA'"
     ElseIf frm.Opt_PayTypeBank.Value = True Then
        vF4 = "{V_PR_EMP_MST.C_PAYTYPE}='BA'"
     End If
     
     EmpMaster_RepFilter = ReportFilterOption(SelFor, vF1, vF2, vF3, vF4)
     
End Function

Public Function SalaryMaster_RepFilter(frm As Form) As String
  Dim vF1 As String, vF2 As String, vF3 As String, vF4 As String, vF5 As String
  Dim SelFor As String
  
     vF1 = "": vF2 = "": vF3 = "": vF4 = "": vF5 = ""
     SelFor = ""
     
     If Right(Trim(frm.Cmb_Company), 7) <> "" Then
        vF1 = "{V_PR_SALARY_MST.C_COMPANY}='" & Right(Trim(frm.Cmb_Company), 7) & "'"
     End If
     
     If Trim(frm.Cmb_Branch) <> "" Then
        vF2 = "{V_PR_SALARY_MST.C_BRANCH}='" & Trim(frm.Cmb_Branch) & "'"
     End If
     
     If Trim(frm.Cmb_Dept) <> "" Then
        vF3 = "{V_PR_SALARY_MST.C_DEPT}='" & Trim(frm.Cmb_Dept) & "'"
     End If
     
     If Trim(frm.Cmb_Desig) <> "" Then
        vF4 = "{V_PR_SALARY_MST.C_DESIG}='" & Trim(frm.Cmb_Desig) & "'"
     End If
     
     If Trim(frm.Cmb_EmpType) <> "" Then
        vF5 = "{V_PR_SALARY_MST.C_EMPTYPE}='" & Trim(frm.Cmb_EmpType) & "'"
     End If
     
     SelFor = ReportFilterOption(vF1, vF2, vF3, vF4, vF5)
     vF1 = "": vF2 = "": vF3 = "": vF4 = "": vF5 = ""
     
     If Trim(frm.Txtc_EmployeeName) <> "" Then
        vF1 = "{V_PR_SALARY_MST.C_EMPNO}='" & Trim(Right(Trim(frm.Txtc_EmployeeName), 7)) & "'"
     End If
     
     If Trim(frm.Cmb_Sex) <> "" Then
        vF2 = "{PR_EMP_MST.C_SEX}='" & Trim(frm.Cmb_Sex) & "'"
     End If
     
     If Trim(frm.Cmb_Nationality) <> "" Then
        vF3 = "{PR_EMP_MST.C_NATIONALITY}='" & Trim(frm.Cmb_Nationality) & "'"
     End If
     
     If Trim(frm.Cmb_DayWork) <> "" Then
        vF4 = "{PR_EMP_MST.C_DAYWORK}='" & Trim(Right(Trim(frm.Cmb_DayWork), 5)) & "'"
     End If
     
     SelFor = ReportFilterOption(SelFor, vF1, vF2, vF3, vF4)
     vF1 = "": vF2 = "": vF3 = "": vF4 = "": vF5 = ""
     
     
     If frm.Opt_Active.Value = True Then
        vF1 = "ISNULL({PR_EMP_MST.D_DOL})"
     ElseIf frm.Opt_Left.Value = True Then
        vF1 = "NOT ISNULL({PR_EMP_MST.D_DOL})"
     End If
     
     If frm.Opt_StaffTypeFlat.Value = True Then
        vF2 = "{V_PR_SALARY_MST.C_STAFFTYPE}='Flat'"
     ElseIf frm.Opt_StaffTypeOt.Value = True Then
        vF2 = "{V_PR_SALARY_MST.C_STAFFTYPE}='OverTime'"
     End If
     
     If frm.Opt_SalaryTypeMon.Value = True Then
        vF3 = "{V_PR_SALARY_MST.C_SALARYTYPE}='Monthly'"
     ElseIf frm.Opt_SalaryTypeHr.Value = True Then
        vF3 = "{V_PR_SALARY_MST.C_SALARYTYPE}='Hourly'"
     End If
     
     If frm.Opt_PayTypeCash.Value = True Then
        vF4 = "{V_PR_SALARY_MST.C_PAYTYPE}='CA'"
     ElseIf frm.Opt_PayTypeBank.Value = True Then
        vF4 = "{V_PR_SALARY_MST.C_PAYTYPE}='BA'"
     End If
     
     SelFor = ReportFilterOption(SelFor, vF1, vF2, vF3, vF4)
     vF1 = "": vF2 = "": vF3 = "": vF4 = "": vF5 = ""
     
     vF1 = "{V_PR_SALARY_MST.N_PERIOD}=" & Is_Null(Format(frm.Txtc_Year, "0000") & Format(frm.Txtc_Month, "00"), True)
     
     SalaryMaster_RepFilter = ReportFilterOption(SelFor, vF1)
     
End Function

Public Function MakeReportHead(frm As Form, ByVal vTitle As String, Optional ByVal AsOn As Boolean) As String
  Dim RepTitle As String
    ' Company name - Report title - Emp status - As on date
    
    ' Company name
    If Trim(frm.Cmb_Company) <> "" Then
       RepTitle = Proper(Trim(Left(Trim(frm.Cmb_Company), 50)))
    End If
    
    ' Report title
    If Trim(RepTitle) = "" Then
       RepTitle = vTitle
    Else
       If Trim(vTitle) <> "" Then
          RepTitle = RepTitle & "  -  " & vTitle
       End If
    End If
    
    ' Emp status
    vTitle = ""
    If frm.Opt_Active.Value = True Then
       vTitle = "Active "
    ElseIf frm.Opt_Left.Value = True Then
       vTitle = "Left "
    ElseIf frm.Opt_BothStatus.Value = True Then
       vTitle = "Both Active and Left "
    End If
    
    If Trim(RepTitle) = "" Then
       RepTitle = vTitle
    Else
       If Trim(vTitle) <> "" Then
          RepTitle = RepTitle & "  -  " & vTitle
       End If
    End If
    
    ' As on
    If AsOn Then
       RepTitle = RepTitle & "  -  " & "As on : " & Format(g_CurrentDate, "dd/mm/yyyy")
    End If
    
    MakeReportHead = RepTitle
End Function

Public Function MakeReportHeadShort(frm As Form, ByVal vTitle As String, Optional ByVal AsOn As Boolean) As String
  Dim RepTitle As String
    ' Company name - Report title - Emp status - As on date
    
    ' Company name
    If Trim(frm.Cmb_Company) <> "" Then
       RepTitle = Proper(Trim(Left(Trim(frm.Cmb_Company), 50)))
    End If
    
    ' Report title
    If Trim(RepTitle) = "" Then
       RepTitle = vTitle
    Else
       If Trim(vTitle) <> "" Then
          RepTitle = RepTitle & "  -  " & vTitle
       End If
    End If
    
    ' As on
    If AsOn Then
       RepTitle = RepTitle & "  -  " & "As on : " & Format(g_CurrentDate, "dd/mm/yyyy")
    End If
    
    MakeReportHeadShort = RepTitle
End Function

Public Function MakeReportSubHead(frm As Form)
  Dim RepDate As String
    ' Filters headers
      
    If Trim(frm.Cmb_Branch) <> "" Then
       RepDate = "Branch : " & Trim(frm.Cmb_Branch)
    End If
    If Trim(frm.Cmb_Dept) <> "" Then
       RepDate = RepDate & Space(15) & "Dept : " & Trim(frm.Cmb_Dept)
    End If
    If Trim(frm.Cmb_Desig) <> "" Then
       RepDate = RepDate & Space(15) & "Desig : " & Trim(frm.Cmb_Desig)
    End If
    If Trim(frm.Cmb_EmpType) <> "" Then
       RepDate = RepDate & Space(15) & "Type : " & Trim(frm.Cmb_EmpType)
    End If
    
    MakeReportSubHead = RepDate
End Function

Public Sub Create_Default_Employee()
'On Error GoTo ErrSave
On Error Resume Next

  Dim rsChk As New ADODB.Recordset, rs As New ADODB.Recordset
  Dim i As Integer
  Dim IsFound As Boolean
  Dim vSqlHIT As String
  
    vSqlHIT = "select employeeid, departmentcode, employeebranchcode, employeefirstname, employeemiddlename, employeelastname, " & _
            "employeeaddress, employeecity, employeecountry, employeetitle, employeegender, employeebirthdate, employeehiredate, " & _
            "employeeoutdate, employeehomephone, employeehandphone, employeeemail " & _
            "From hitfpta.dbo.employee "
  
    Set rsChk = Nothing
    g_Sql = "select c_empno from pr_emp_mst "
    rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    If rsChk.RecordCount > 0 Then
       vSqlHIT = vSqlHIT & " Where employeeid not in (Select c_empno from pr_emp_mst)"
    End If
    
    Set rsChk = Nothing
    rsChk.Open vSqlHIT, CON, adOpenForwardOnly, adLockOptimistic
    If rsChk.RecordCount > 0 Then
       Screen.MousePointer = vbHourglass
       g_SaveFlagNull = True
       
       Set rs = Nothing
       g_Sql = "select * from pr_emp_mst "
       rs.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
       
       For i = 1 To rsChk.RecordCount
           CON.BeginTrans
           rs.AddNew
           
           rs("c_empno").Value = Is_Null(rsChk("employeeid").Value, False)
           If rsChk("employeetitle").Value <> "" Then
              rs("c_title").Value = Proper(Is_Null(rsChk("employeetitle").Value, False))
           End If
           rs("c_name").Value = Is_Null(rsChk("employeefirstname").Value, False)
           rs("c_othername").Value = Is_Null(rsChk("employeemiddlename").Value, False) & " " & Is_Null(rsChk("employeelastname").Value, False)
           If rsChk("employeegender").Value <> "" Then
              rs("c_sex").Value = Proper(Is_Null(rsChk("employeegender").Value, False))
           End If
           
           If rsChk("employeeaddress").Value <> "" Then
              rs("c_address").Value = Proper(Is_Null(rsChk("employeeaddress").Value, False)) & vbCrLf & Is_Null(rsChk("employeecity").Value, False)
           End If
           rs("c_phone").Value = Trim(Is_Null(rsChk("employeehomephone").Value, False) & "    " & Is_Null(rsChk("employeehandphone").Value, False))
           rs("c_email").Value = Is_Null(rsChk("employeeemail").Value, False)
           rs("c_nationality").Value = Is_Null(rsChk("employeecountry").Value, False)
           
           rs("d_dob").Value = Is_Date(Is_Null(rsChk("employeebirthdate").Value, False), "S")
           rs("d_doj").Value = Is_Date(Is_Null(rsChk("employeehiredate").Value, False), "S")
           rs("d_dol").Value = Is_Date(Is_Null(rsChk("employeeoutdate").Value, False), "S")
           
           rs("c_company").Value = "COM0001"
           If rsChk("employeebranchcode").Value <> "" Then
              rs("c_branch").Value = Is_Null(rsChk("employeebranchcode").Value, False)
           Else
              rs("c_branch").Value = g_BranchDefault
           End If
           If rsChk("departmentcode").Value <> "" Then
              rs("c_dept").Value = Proper(Is_Null(rsChk("departmentcode").Value, False))
           Else
              rs("c_dept").Value = g_DeptDefault
           End If
           
           rs("c_expatriate").Value = "Yes"
           rs("c_emptype").Value = "Worker"
           rs("c_stafftype").Value = "O"
           rs("c_clockcard").Value = "1"
           rs("c_edfcat").Value = "A"
           rs("c_tpflag").Value = "No"
           rs("c_paytype").Value = "CA"
           rs("c_daywork").Value = "5D"
           rs("c_shiftcode").Value = "S01"
           
           rs("c_rec_sta").Value = "A"
           rs("c_usr_id").Value = g_UserName
           rs("d_created").Value = GetDateTime
           
           rs.Update
           CON.CommitTrans
           
           rsChk.MoveNext
       Next i
       Screen.MousePointer = vbDefault
       g_SaveFlagNull = False
    
    End If
   
   Exit Sub
     
ErrSave:
     CON.RollbackTrans
     g_SaveFlagNull = False
     Screen.MousePointer = vbDefault
    ' MsgBox "Error while Saving - " + Err.Description, vbCritical, "Critical"
End Sub

Public Sub Create_Default_Company()
On Error GoTo Err_Save
   Dim DyDisp As New ADODB.Recordset

     ' Creating default
     Set DyDisp = Nothing
     g_Sql = "select * from pr_company_mst "
     DyDisp.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
   
     If DyDisp.RecordCount = 0 Then
        g_Sql = "Insert into pr_company_mst " & _
                "       (c_company, c_companyname, c_displayname, c_add1, c_add2, " & _
                "        c_country, c_tel, c_email, c_web, c_rec_sta, c_usr_id, d_created, " & _
                "        n_empnpf, n_empnpfmin, n_empnpfmax, n_comnpf, n_comnpfmin, n_comnpfmax, " & _
                "        n_empmed, n_empmedmin, n_empmedmax, n_commed, n_commedmin, n_commedmax, " & _
                "        n_emplevy, n_emplevymin, n_emplevymax, n_comlevy, n_comlevymin, n_comlevymax, " & _
                "        n_empepz, n_empepzmin, n_empepzmax, n_comepz, n_comepzmin, n_comepzmax)   " & _
                "values ('COM0001', 'SIFB', 'Sugar Insurance Fund Board', '18, Sir Seewoosagur Ramgoolam Street', 'Port-Louis', " & _
                "        'Mauritius', '(230) 208 3236', 'contactus@sifb.mu', 'www.sifb.mu', 'A', 'Admin', getdate(), " & _
                "        3, 2, 510, 6, 4, 1020, 1, 1, 170, 2.5, 2, 425, 0, 0, 0, 1.5, 0, 0, 3, 0, 0, 9, 0, 0)"
        CON.Execute g_Sql
     End If
  
 Exit Sub

Err_Save:
   MsgBox "Error on creating default company - " + Err.Description, vbCritical, "Critical"
End Sub

Public Sub Create_Default_Bank()
On Error GoTo Err_Save
   Dim DyDisp As New ADODB.Recordset

     ' Creating default
     Set DyDisp = Nothing
     g_Sql = "select * from pr_bankmast "
     DyDisp.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
   
     If DyDisp.RecordCount = 0 Then
        g_Sql = "Insert into pr_bankmast " & _
                "       (c_code, c_bankname, c_shortname, c_bankcode,  c_rec_sta, d_created) " & _
                "values ('B01', 'State Bank of Mauritius', 'SBM', '001', 'A', getdate())"
        CON.Execute g_Sql
     
        g_Sql = "Insert into pr_bankmast " & _
                "       (c_code, c_bankname, c_shortname, c_bankcode,  c_rec_sta, d_created) " & _
                "values ('B02', 'Mauritian Commercial Bank', 'MCB', '002', 'A', getdate())"
        CON.Execute g_Sql
     
        g_Sql = "Insert into pr_bankmast " & _
                "       (c_code, c_bankname, c_shortname, c_bankcode,  c_rec_sta, d_created) " & _
                "values ('B03', 'Barclays Bank', 'BB', '003', 'A', getdate())"
        CON.Execute g_Sql
     End If
  
 Exit Sub

Err_Save:
   MsgBox "Error on creating default company - " + Err.Description, vbCritical, "Critical"
End Sub

Public Sub Create_Default_Shift()
On Error GoTo Err_Save
   Dim DyDisp As New ADODB.Recordset

     ' Creating default
     Set DyDisp = Nothing
     g_Sql = "select * from pr_clock_shift "
     DyDisp.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
   
     If DyDisp.RecordCount = 0 Then
        g_Sql = "Insert into pr_clock_shift " & _
                "       (c_shiftcode, n_starthrs, n_endhrs, n_breakhrs, n_latehrs, n_shifthrs, n_break1, n_mins1, n_break2, n_mins2, n_break3, n_mins3, " & _
                "        n_permhrs, n_clmin, n_cutoffhrs, n_maxhrs, c_desp, " & _
                "        starthrs, endhrs, breakhrs, latehrs, shifthrs, break1, mins1, break2, mins2, break3, mins3, permhrs, clmin, cutoffhrs, maxhrs) " & _
                "values ('S1', 8.00, 16.30, 30, 10, 8, 12.15, 30, 0, 0, 0, 0, 10, 10, 24, 24, 'General Shift', " & _
                "        480, 990, 35, 10, 480, 735, 30, 0, 0, 0, 0, 10, 10, 1440, 1440) "
        CON.Execute g_Sql
     
        g_Sql = "Insert into pr_clock_shift " & _
                "       (c_shiftcode, n_starthrs, n_endhrs, n_breakhrs, n_latehrs, n_shifthrs, n_break1, n_mins1, n_break2, n_mins2, n_break3, n_mins3, " & _
                "        n_permhrs, n_clmin, n_cutoffhrs, n_maxhrs, c_desp, " & _
                "        starthrs, endhrs, breakhrs, latehrs, shifthrs, break1, mins1, break2, mins2, break3, mins3, permhrs, clmin, cutoffhrs, maxhrs) " & _
                "values ('WO', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 'Work Off', " & _
                "        0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0) "
        CON.Execute g_Sql
     End If
  
 Exit Sub

Err_Save:
   MsgBox "Error on creating default company - " + Err.Description, vbCritical, "Critical"
End Sub

Public Sub Create_Default_ShiftStructure()
On Error GoTo Err_Save
   Dim DyDisp As New ADODB.Recordset

     ' Creating default
     Set DyDisp = Nothing
     g_Sql = "select * from pr_shiftstructure_mst "
     DyDisp.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
   
     If DyDisp.RecordCount = 0 Then
        g_Sql = "Insert into pr_shiftstructure_mst " & _
                "       (c_code, c_shiftname,  c_rec_sta, d_created) " & _
                "values ('S01', 'General Shift', 'A', getdate())"
        CON.Execute g_Sql
     
     
        g_Sql = "Insert into pr_shiftstructure_dtl " & _
                "(c_code, n_wkday, c_shiftcode, c_wocode, c_altcode, n_slno) values ('S01', 1, 'WO', 'S1', NULL, 7)"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_shiftstructure_dtl " & _
                "(c_code, n_wkday, c_shiftcode, c_wocode, c_altcode, n_slno) values ('S01', 2, 'S1', NULL, NULL, 1)"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_shiftstructure_dtl " & _
                "(c_code, n_wkday, c_shiftcode, c_wocode, c_altcode, n_slno) values ('S01', 3, 'S1', NULL, NULL, 2)"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_shiftstructure_dtl " & _
                "(c_code, n_wkday, c_shiftcode, c_wocode, c_altcode, n_slno) values ('S01', 4, 'S1', NULL, NULL, 3)"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_shiftstructure_dtl " & _
                "(c_code, n_wkday, c_shiftcode, c_wocode, c_altcode, n_slno) values ('S01', 5, 'S1', NULL, NULL, 4)"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_shiftstructure_dtl " & _
                "(c_code, n_wkday, c_shiftcode, c_wocode, c_altcode, n_slno) values ('S01', 6, 'S1', NULL, NULL, 5)"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_shiftstructure_dtl " & _
                "(c_code, n_wkday, c_shiftcode, c_wocode, c_altcode, n_slno) values ('S01', 7, 'WO', 'S1', NULL, 6)"
        CON.Execute g_Sql
     
  
     End If
  
 Exit Sub

Err_Save:
   MsgBox "Error on creating default company - " + Err.Description, vbCritical, "Critical"
End Sub

Public Sub Create_Default_LeaveAllot()
On Error GoTo Err_Save
   Dim DyDisp As New ADODB.Recordset

     ' Creating default
     Set DyDisp = Nothing
     g_Sql = "select * from pr_leaveallot_mst "
     DyDisp.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
   
     If DyDisp.RecordCount = 0 Then
        g_Sql = "Insert into pr_leaveallot_mst " & _
                "       (c_leave, n_yearfrom,  c_rec_sta, d_created) " & _
                "values ('VL', 2000, 'A', getdate())"
        CON.Execute g_Sql
    
     
        g_Sql = "Insert into pr_leaveallot_dtl " & _
                "(c_leave, n_yearfrom, n_from, n_to, n_allot, n_max) values ('VL', 2000, 0, 5, 25, 105)"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_leaveallot_dtl " & _
                "(c_leave, n_yearfrom, n_from, n_to, n_allot, n_max) values ('VL', 2000, 5, 10, 30, 140)"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_leaveallot_dtl " & _
                "(c_leave, n_yearfrom, n_from, n_to, n_allot, n_max) values ('VL', 2000, 10, 15, 35, 175)"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_leaveallot_dtl " & _
                "(c_leave, n_yearfrom, n_from, n_to, n_allot, n_max) values ('VL', 2000, 15, 100, 35, 210)"
        CON.Execute g_Sql
     End If
  
 Exit Sub

Err_Save:
   MsgBox "Error on creating default company - " + Err.Description, vbCritical, "Critical"
End Sub

Public Sub Create_Default_PayStructure(ByVal vCompany As String)
  Dim DyDisp As New ADODB.Recordset
    
     If Trim(vCompany) = "" Then
        Exit Sub
     End If
  
     vCompany = Trim(vCompany)
     
     ' Creating defaul pay structure master
     Set DyDisp = Nothing
     g_Sql = "select * from pr_paystructure_mst where c_company = '" & vCompany & "'"
     DyDisp.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
   
     If DyDisp.RecordCount = 0 Then
        g_Sql = "Insert into pr_paystructure_mst (c_company, c_remarks, c_rec_sta, d_created) values " & _
                "('" & vCompany & "', 'System default', 'A', getdate())"
        CON.Execute g_Sql
     End If
  
     ' Creating defaul pay structure details
     Set DyDisp = Nothing
     g_Sql = "select * from pr_paystructure_dtl where c_company = '" & vCompany & "'"
     DyDisp.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
   
     If DyDisp.RecordCount = 0 Then
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0001', 1, 'Basic', '1', 'Y', 'Y', 'Y', 'Y', 'A')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0002', 2, 'Fixed Rate 1', '1', 'Y', 'Y', 'Y', 'Y', 'A')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0003', 3, 'Fixed Rate 2', '1', 'Y', 'Y', 'Y', 'Y', 'A')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0004', 4, 'Allowance', '1', 'Y', 'Y', 'Y', 'Y', 'A')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0005', 5, 'Special Allowance', '1', 'Y', 'Y', 'Y', 'Y', 'A')"
        CON.Execute g_Sql
        
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0006', 6, 'OT 1.5', '1', 'Y', 'Y', 'Y', 'N', 'A')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0007', 7, 'OT 2', '1', 'Y', 'Y', 'Y', 'N', 'A')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0008', 8, 'OT 3', '1', 'Y', 'Y', 'Y', 'N', 'A')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0009', 9, 'SUN/PH 2', '1', 'Y', 'Y', 'Y', 'N', 'A')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0010', 10, 'SUN/PH 3', '1', 'Y', 'Y', 'Y', 'N', 'A')"
        CON.Execute g_Sql
        
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0015', 11, 'Travel Allowance', '1', 'N', 'N', 'Y', 'Y', 'E')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0016', 12, 'Travel Allowance Taxable', '1', 'Y', 'N', 'Y', 'Y', 'M')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0017', 13, 'Attendance Bonus', '1', 'Y', 'Y', 'Y', 'N', 'A')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0018', 14, 'Meal Allowance', '1', 'Y', 'Y', 'Y', 'N', 'A')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0019', 15, 'Night Allowance', '1', 'Y', 'Y', 'Y', 'N', 'A')"
        CON.Execute g_Sql
        
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0021', 16, 'Public Holiday', '1', 'Y', 'Y', 'Y', 'N', 'A')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0022', 17, 'Local Leave', '1', 'Y', 'Y', 'Y', 'N', 'A')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0023', 18, 'Sick Leave', '1', 'Y', 'Y', 'Y', 'N', 'A')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0024', 19, 'Injury Leave', '1', 'Y', 'Y', 'Y', 'N', 'A')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0025', 20, 'Prolong Illness', '1', 'Y', 'Y', 'Y', 'N', 'A')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0026', 21, 'Wedding Leave', '1', 'Y', 'Y', 'Y', 'N', 'A')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0027', 22, 'Maternity Leave', '1', 'Y', 'Y', 'Y', 'N', 'A')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0028', 23, 'Paternity Leave', '1', 'Y', 'Y', 'Y', 'N', 'A')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0029', 24, 'Company Leave', '1', 'Y', 'Y', 'Y', 'N', 'A')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0030', 25, 'Other Leave', '1', 'Y', 'Y', 'Y', 'N', 'A')"
        CON.Execute g_Sql
        
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0035', 26, 'E.O.Y Bonus', '1', 'Y', 'N', 'Y', 'N', 'A')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0036', 27, 'Adjusted E.O.Y Bonus', '1', 'Y', 'N', 'Y', 'N', 'A')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0041', 28, 'Cash added (Round-off)', '1', 'Y', 'N', 'Y', 'N', 'A')"
        CON.Execute g_Sql
        
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0051', 29, 'Absent', '2', 'Y', 'Y', 'Y', 'N', 'A')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0052', 30, 'Dedtn. on Late/Perm', '2', 'Y', 'Y', 'Y', 'N', 'A')"
        CON.Execute g_Sql
        
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0056', 31, 'N.P.S', '2', 'N', 'N', 'Y', 'N', 'A')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0057', 32, 'EPZ', '2', 'N', 'N', 'Y', 'N', 'A')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0058', 33, 'N.S.F', '2', 'N', 'N', 'Y', 'N', 'A')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0059', 34, 'LEVY', '2', 'N', 'N', 'Y', 'N', 'A')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0060', 35, 'PAYE', '2', 'N', 'N', 'Y', 'N', 'A')"
        CON.Execute g_Sql
        
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0061', 36, 'Salary Advance', '2', 'N', 'N', 'Y', 'N', 'A')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0066', 37, 'Cash deducted (Round-off)', '2', 'Y', 'N', 'Y', 'N', 'A')"
        CON.Execute g_Sql
        
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0071', 38, 'Company N.P.S', '3', 'N', 'N', 'Y', 'N', Null)"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0072', 39, 'Company EPZ', '3', 'N', 'N', 'Y', 'N', Null)"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0073', 40, 'Company E.W.F', '3', 'N', 'N', 'Y', 'N', Null)"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0074', 41, 'Company Levy', '3', 'N', 'N', 'Y', 'N', Null)"
        CON.Execute g_Sql

        ' client specific pay types
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0081', 42, 'Travel Refund', '1', 'N', 'N', 'N', 'N', 'E')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0082', 43, 'Bonus', '1', 'Y', 'Y', 'N', 'N', 'B')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0083', 44, 'Production Bonus', '1', 'Y', 'Y', 'N', 'N', 'B')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_paystructure_dtl (c_company, c_salary, n_seq, c_payname, c_type, c_paye, c_bonus, c_syscal, c_master, c_category) " & _
                "values ('" & vCompany & "', 'SAL0084', 45, 'Performance Bonus', '1', 'Y', 'Y', 'N', 'N', 'B')"
        CON.Execute g_Sql

     End If

 Exit Sub

Err_Save:
   MsgBox "Error on creating default company - " + Err.Description, vbCritical, "Critical"
End Sub

Public Sub Create_Default_LeaveTypes()
On Error GoTo Err_Save
   Dim DyDisp As New ADODB.Recordset

     ' Creating defaul leavetypes
     Set DyDisp = Nothing
     g_Sql = "select * from pr_leave_mst "
     DyDisp.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
   
     If DyDisp.RecordCount = 0 Then
        g_Sql = "Insert into pr_leave_mst (c_leave, c_leavename, c_remarks, c_rec_sta, d_created)  values " & _
                "('P', 'Present', 'System default', 'A', getdate() )"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_leave_mst (c_leave, c_leavename, c_remarks, c_rec_sta, d_created)  values " & _
                "('A', 'Absent', 'System default', 'A', getdate() )"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_leave_mst (c_leave, c_leavename, c_remarks, c_rec_sta, d_created)  values " & _
                "('WO', 'Work Off', 'System default', 'A', getdate() )"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_leave_mst (c_leave, c_leavename, c_remarks, c_rec_sta, d_created)  values " & _
                "('PH', 'Public Holiday', 'System default', 'A', getdate() )"
        CON.Execute g_Sql
        
        g_Sql = "Insert into pr_leave_mst (c_leave, c_leavename, c_remarks, c_rec_sta, d_created)  values " & _
                "('CL', 'Casual Leave', 'System default', 'A', getdate() )"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_leave_mst (c_leave, c_leavename, c_remarks, c_rec_sta, d_created)  values " & _
                "('SL', 'Sick Leave', 'System default', 'A', getdate() )"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_leave_mst (c_leave, c_leavename, c_remarks, c_rec_sta, d_created)  values " & _
                "('IL', 'Injury Leave', 'System default', 'A', getdate() )"
        CON.Execute g_Sql
        
        g_Sql = "Insert into pr_leave_mst (c_leave, c_leavename, c_remarks, c_rec_sta, d_created)  values " & _
                "('DV', 'Daily Vacation', 'System default', 'A', getdate() )"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_leave_mst (c_leave, c_leavename, c_remarks, c_rec_sta, d_created)  values " & _
                "('VL', 'Vacation Leave', 'System default', 'A', getdate() )"
        CON.Execute g_Sql
        
        g_Sql = "Insert into pr_leave_mst (c_leave, c_leavename, c_remarks, c_rec_sta, d_created)  values " & _
                "('ML', 'Maternity Leave', 'System default', 'A', getdate() )"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_leave_mst (c_leave, c_leavename, c_remarks, c_rec_sta, d_created)  values " & _
                "('SP', 'Special Leave', 'System default', 'A', getdate() )"
        CON.Execute g_Sql
        
        g_Sql = "Insert into pr_leave_mst (c_leave, c_leavename, c_remarks, c_rec_sta, d_created)  values " & _
                "('LWP', 'Leave Without Pay', 'System default', 'A', getdate() )"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_leave_mst (c_leave, c_leavename, c_remarks, c_rec_sta, d_created)  values " & _
                "('UNL', 'Unauthorized Leave', 'System default', 'A', getdate() )"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_leave_mst (c_leave, c_leavename, c_remarks, c_rec_sta, d_created)  values " & _
                "('PRL', 'Pre-retirement Leave', 'System default', 'A', getdate() )"
        CON.Execute g_Sql
     
        g_Sql = "Insert into pr_leave_mst (c_leave, c_leavename, c_remarks, c_rec_sta, d_created)  values " & _
                "('TO', 'Time off in lieu of Overtime', 'System default', 'A', getdate() )"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_leave_mst (c_leave, c_leavename, c_remarks, c_rec_sta, d_created)  values " & _
                "('IN', 'Interdiction ', 'System default', 'A', getdate() )"
        CON.Execute g_Sql
     End If
  
 Exit Sub

Err_Save:
   MsgBox "Error on creating default leaves - " + Err.Description, vbCritical, "Critical"
End Sub


Public Sub Create_Default_EDFTypes()
On Error GoTo Err_Display
Dim DyDisp As New ADODB.Recordset

     ' Creating defaul edf categories
     Set DyDisp = Nothing
     g_Sql = "select * from pr_edfmast"
     DyDisp.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
   
     If DyDisp.RecordCount = 0 Then
        g_Sql = "Insert into pr_edfmast (c_category, c_desp, n_edfamt)  values " & _
                "('A', 'Personal', 300000 )"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_edfmast (c_category, c_desp, n_edfamt)  values " & _
                "('B', 'Personal + One Dependent', 410000 )"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_edfmast (c_category, c_desp, n_edfamt)  values " & _
                "('C', 'Personal + Two Dependents', 475000 )"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_edfmast (c_category, c_desp, n_edfamt)  values " & _
                "('D', 'Personal + Three Dependents', 520000 )"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_edfmast (c_category, c_desp, n_edfamt)  values " & _
                "('E', 'Personal + Four Dependents', 550000 )"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_edfmast (c_category, c_desp, n_edfamt)  values " & _
                "('F', 'Retired', 350000 )"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_edfmast (c_category, c_desp, n_edfamt)  values " & _
                "('G', 'Retired + One Dependents', 460000 )"
        CON.Execute g_Sql
    End If
 
 Exit Sub

Err_Display:
     MsgBox "Critical Erro - EDF Cateogry Creation - " & Err.Description, vbCritical, "Critical"
End Sub

Public Sub Create_Default_Screens()
On Error GoTo Err_Save
   Dim DyDisp As New ADODB.Recordset

     ' Creating default
     Set DyDisp = Nothing
     g_Sql = "select * from pr_screen_mst "
     DyDisp.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
   
     If DyDisp.RecordCount = 0 Then
        
        g_Sql = "Insert into pr_screen_mst (c_screen_id, c_screen_name) values ('frm_company', 'Company Master')"
        CON.Execute g_Sql
        
        g_Sql = "Insert into pr_screen_mst (c_screen_id, c_screen_name) values ('frm_shifts', 'Shift Master')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_screen_mst (c_screen_id, c_screen_name) values ('frm_shift_structure', 'Shift Structure Master')"
        CON.Execute g_Sql
        
        g_Sql = "Insert into pr_screen_mst (c_screen_id, c_screen_name) values ('frm_payperiod', 'Pay Period Master')"
        CON.Execute g_Sql
        
        g_Sql = "Insert into pr_screen_mst (c_screen_id, c_screen_name) values ('frm_leavemaster', 'Leave Master')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_screen_mst (c_screen_id, c_screen_name) values ('frm_leave_allotment', 'Leave Allotment')"
        CON.Execute g_Sql
     
        g_Sql = "Insert into pr_screen_mst (c_screen_id, c_screen_name) values ('frm_holiday', 'Holiday Master')"
        CON.Execute g_Sql
        
        g_Sql = "Insert into pr_screen_mst (c_screen_id, c_screen_name) values ('frm_emp_master', 'Employee Master')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_screen_mst (c_screen_id, c_screen_name) values ('frm_emp_upd_E', 'Employee Master Multi-Updates')"
        CON.Execute g_Sql
        
        g_Sql = "Insert into pr_screen_mst (c_screen_id, c_screen_name) values ('frm_Leave_Adj', 'Leave Encash/Adjustment')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_screen_mst (c_screen_id, c_screen_name) values ('frm_Leave_Entry', 'Leave Entry Details')"
        CON.Execute g_Sql
        
        g_Sql = "Insert into pr_screen_mst (c_screen_id, c_screen_name) values ('frm_clock_dataprocess_prp', 'Process Screen')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_screen_mst (c_screen_id, c_screen_name) values ('frm_clock_dataprocess_dwh', 'Data Capture and Process')"
        CON.Execute g_Sql
        
        g_Sql = "Insert into pr_screen_mst (c_screen_id, c_screen_name) values ('frm_clock_emp', 'Attendance Details')"
        CON.Execute g_Sql
        
        
'        g_Sql = "Insert into pr_screen_mst (c_screen_id, c_screen_name) values ('frm_bankmaster', 'Bank Master')"
'        CON.Execute g_Sql
'        g_Sql = "Insert into pr_screen_mst (c_screen_id, c_screen_name) values ('frm_pay_structure', 'Pay Structure Master')"
'        CON.Execute g_Sql
'        g_Sql = "Insert into pr_screen_mst (c_screen_id, c_screen_name) values ('frm_edf_master', 'EDF Master')"
'        CON.Execute g_Sql

'        g_Sql = "Insert into pr_screen_mst (c_screen_id, c_screen_name) values ('frm_emp_upd_T', 'Transport Planning')"
'        CON.Execute g_Sql
'        g_Sql = "Insert into pr_screen_mst (c_screen_id, c_screen_name) values ('frm_loan', 'Loan/Advance Details')"
'        CON.Execute g_Sql
'        g_Sql = "Insert into pr_screen_mst (c_screen_id, c_screen_name) values ('frm_addpay_details', 'Additional Income/Deduction Details')"
'        CON.Execute g_Sql

'        g_Sql = "Insert into pr_screen_mst (c_screen_id, c_screen_name) values ('frm_salary_details', 'Salary Details')"
'        CON.Execute g_Sql
        
     End If
  
 Exit Sub

Err_Save:
   MsgBox "Error on creating default company - " + Err.Description, vbCritical, "Critical"
End Sub


Public Sub Create_Default_PayPeriods(ByVal vYear As Integer)
  Dim DyDisp As New ADODB.Recordset
  Dim vFYStart As String, vFYEnd As String
  
     If vYear = 0 Then
        Exit Sub
     End If
  
     vFYStart = Trim(Str(vYear - 1)) & "-07-01"
     vFYEnd = Trim(Str(vYear)) & "-06-30"
     
     ' Creating defaul pay period master
     Set DyDisp = Nothing
     g_Sql = "select * from pr_payperiod_mst where n_year = " & vYear & " and c_type = 'W'"
     DyDisp.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
     If DyDisp.RecordCount = 0 Then
        g_Sql = "Insert into pr_payperiod_mst ( n_year, c_type, d_fystart, d_fyend, c_remarks, c_rec_sta, d_created) values " & _
                "(" & vYear & ", 'W', '" & vFYStart & "', '" & vFYEnd & "', 'Pay periods for " & Trim(Str(vYear)) & "', 'A', getdate())"
        CON.Execute g_Sql
     End If
  
     ' Creating defaul pay period details
     Set DyDisp = Nothing
     g_Sql = "select * from pr_payperiod_dtl where n_year = " & vYear & " and c_type = 'W'"
     DyDisp.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
     If DyDisp.RecordCount = 0 Then
        g_Sql = "Insert into pr_payperiod_dtl (c_type, n_period, n_year, d_fromdate, d_todate, c_period_closed) " & _
                "values ('W', " & Val(Trim(Str(vYear)) & "01") & ", " & vYear & ", '" & Trim(Str(vYear)) & "-01-01" & "', '" & Trim(Str(vYear)) & "-01-31" & "', 'N')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_payperiod_dtl (c_type, n_period, n_year, d_fromdate, d_todate, c_period_closed) " & _
                "values ('W', " & Val(Trim(Str(vYear)) & "02") & ", " & vYear & ", '" & Trim(Str(vYear)) & "-02-01" & "', '" & Trim(Str(vYear)) & "-02-28" & "', 'N')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_payperiod_dtl (c_type, n_period, n_year, d_fromdate, d_todate, c_period_closed) " & _
                "values ('W', " & Val(Trim(Str(vYear)) & "03") & ", " & vYear & ", '" & Trim(Str(vYear)) & "-03-01" & "', '" & Trim(Str(vYear)) & "-03-31" & "', 'N')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_payperiod_dtl (c_type, n_period, n_year, d_fromdate, d_todate, c_period_closed) " & _
                "values ('W', " & Val(Trim(Str(vYear)) & "04") & ", " & vYear & ", '" & Trim(Str(vYear)) & "-04-01" & "', '" & Trim(Str(vYear)) & "-04-30" & "', 'N')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_payperiod_dtl (c_type, n_period, n_year, d_fromdate, d_todate, c_period_closed) " & _
                "values ('W', " & Val(Trim(Str(vYear)) & "05") & ", " & vYear & ", '" & Trim(Str(vYear)) & "-05-01" & "', '" & Trim(Str(vYear)) & "-05-31" & "', 'N')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_payperiod_dtl (c_type, n_period, n_year, d_fromdate, d_todate, c_period_closed) " & _
                "values ('W', " & Val(Trim(Str(vYear)) & "06") & ", " & vYear & ", '" & Trim(Str(vYear)) & "-06-01" & "', '" & Trim(Str(vYear)) & "-06-30" & "', 'N')"
        CON.Execute g_Sql

        g_Sql = "Insert into pr_payperiod_dtl (c_type, n_period, n_year, d_fromdate, d_todate, c_period_closed) " & _
                "values ('W', " & Val(Trim(Str(vYear)) & "07") & ", " & vYear & ", '" & Trim(Str(vYear)) & "-07-01" & "', '" & Trim(Str(vYear)) & "-07-31" & "', 'N')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_payperiod_dtl (c_type, n_period, n_year, d_fromdate, d_todate, c_period_closed) " & _
                "values ('W', " & Val(Trim(Str(vYear)) & "08") & ", " & vYear & ", '" & Trim(Str(vYear)) & "-08-01" & "', '" & Trim(Str(vYear)) & "-08-31" & "', 'N')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_payperiod_dtl (c_type, n_period, n_year, d_fromdate, d_todate, c_period_closed) " & _
                "values ('W', " & Val(Trim(Str(vYear)) & "09") & ", " & vYear & ", '" & Trim(Str(vYear)) & "-09-01" & "', '" & Trim(Str(vYear)) & "-09-30" & "', 'N')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_payperiod_dtl (c_type, n_period, n_year, d_fromdate, d_todate, c_period_closed) " & _
                "values ('W', " & Val(Trim(Str(vYear)) & "10") & ", " & vYear & ", '" & Trim(Str(vYear)) & "-10-01" & "', '" & Trim(Str(vYear)) & "-10-31" & "', 'N')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_payperiod_dtl (c_type, n_period, n_year, d_fromdate, d_todate, c_period_closed) " & _
                "values ('W', " & Val(Trim(Str(vYear)) & "11") & ", " & vYear & ", '" & Trim(Str(vYear)) & "-11-01" & "', '" & Trim(Str(vYear)) & "-11-30" & "', 'N')"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_payperiod_dtl (c_type, n_period, n_year, d_fromdate, d_todate, c_period_closed) " & _
                "values ('W', " & Val(Trim(Str(vYear)) & "12") & ", " & vYear & ", '" & Trim(Str(vYear)) & "-12-01" & "', '" & Trim(Str(vYear)) & "-12-31" & "', 'N')"
        CON.Execute g_Sql

        g_Sql = "Insert into pr_payperiod_dtl (c_type, n_period, n_year, d_fromdate, d_todate, c_period_closed) " & _
                "values ('W', " & Val(Trim(Str(vYear)) & "13") & ", " & vYear & ", '" & Trim(Str(vYear)) & "-01-01" & "', '" & Trim(Str(vYear)) & "-12-31" & "', 'N')"
        CON.Execute g_Sql

     End If

 Exit Sub

Err_Save:
   MsgBox "Error on creating default company - " + Err.Description, vbCritical, "Critical"
End Sub

Public Sub Create_Default_User()
On Error GoTo Err_Save
   Dim DyDisp As New ADODB.Recordset

     ' Creating default
     Set DyDisp = Nothing
     g_Sql = "select * from pr_user_mst "
     DyDisp.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
   
     If DyDisp.RecordCount = 0 Then
        g_Sql = "Insert into pr_user_mst " & _
                "       (c_user_id, c_user_name, c_user_pwd, c_user_level, d_pwd_on, c_rec_sta, d_created) " & _
                "values ('Admin', 'Admin', 'Admin', 2, getdate(), 'A', getdate())"
        CON.Execute g_Sql
     
        g_Sql = "Insert into pr_user_mst " & _
                "       (c_user_id, c_user_name, c_user_pwd, c_user_level, d_pwd_on, c_rec_sta, d_created) " & _
                "values ('Guest', 'Guest', 'Guest', 0, getdate(), 'A', getdate())"
        CON.Execute g_Sql
     End If
  
 Exit Sub

Err_Save:
   MsgBox "Error on creating default company - " + Err.Description, vbCritical, "Critical"
End Sub

Public Sub Create_Default_Holidays(ByVal vYear As Integer)
  Dim DyDisp As New ADODB.Recordset
     
     If vYear = 0 Then
        Exit Sub
     End If
     
     ' Creating defaul public holidays
     Set DyDisp = Nothing
     g_Sql = "select * from pr_holiday_mst where year(d_phdate) = " & vYear
     DyDisp.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
     If DyDisp.RecordCount = 0 Then
        g_Sql = "Insert into pr_holiday_mst (n_uno, d_phdate, c_desp, c_type, n_no, c_rec_sta, d_created) " & _
                "values ((Select max(n_uno) from pr_holiday_mst)+1, '" & Trim(Str(vYear) & "-01-01") & "', 'New Year', 'PH', 1, 'A', getdate())"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_holiday_mst (n_uno, d_phdate, c_desp, c_type, n_no, c_rec_sta, d_created) " & _
                "values ((Select max(n_uno) from pr_holiday_mst)+1, '" & Trim(Str(vYear) & "-02-01") & "', 'Abolition of Slavery', 'PH', 1, 'A', getdate())"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_holiday_mst (n_uno, d_phdate, c_desp, c_type, n_no, c_rec_sta, d_created) " & _
                "values ((Select max(n_uno) from pr_holiday_mst)+1, '" & Trim(Str(vYear) & "-03-02") & "', 'National Day', 'PH', 1, 'A', getdate())"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_holiday_mst (n_uno, d_phdate, c_desp, c_type, n_no, c_rec_sta, d_created) " & _
                "values ((Select max(n_uno) from pr_holiday_mst)+1, '" & Trim(Str(vYear) & "-05-01") & "', 'Labour Day', 'PH', 1, 'A', getdate())"
        CON.Execute g_Sql
        g_Sql = "Insert into pr_holiday_mst (n_uno, d_phdate, c_desp, c_type, n_no, c_rec_sta, d_created) " & _
                "values ((Select max(n_uno) from pr_holiday_mst)+1, '" & Trim(Str(vYear) & "-12-25") & "', 'Christmas', 'PH', 1, 'A', getdate())"
        CON.Execute g_Sql
     End If

 Exit Sub

Err_Save:
   MsgBox "Error on creating default company - " + Err.Description, vbCritical, "Critical"
End Sub

Public Sub Create_Default_Clock_Filter()
On Error GoTo Err_Save
   Dim DyDisp As New ADODB.Recordset

     ' Creating default
     Set DyDisp = Nothing
     g_Sql = "select * from pr_clock_filter "
     DyDisp.Open g_Sql, CON, adOpenDynamic, adLockOptimistic
   
     If DyDisp.RecordCount = 0 Then
        g_Sql = "Insert into pr_clock_filter " & _
                "       (desp, fieldname, type, flag) " & _
                "values ('FP/Card No.', 'b.c_clockidno', 'C', 'U')"
        CON.Execute g_Sql
        
        g_Sql = "Insert into pr_clock_filter " & _
                "       (desp, fieldname, type, flag) " & _
                "values ('Shift','c_shift','C','U')"
        CON.Execute g_Sql
        
        g_Sql = "Insert into pr_clock_filter " & _
                "       (desp, fieldname, type, flag) " & _
                "values ('In Time','n_arrtime','N','U')"
        CON.Execute g_Sql
        
        g_Sql = "Insert into pr_clock_filter " & _
                "       (desp, fieldname, type, flag) " & _
                "values ('Late Hrs','n_latehrs','N','U')"
        CON.Execute g_Sql
        
        g_Sql = "Insert into pr_clock_filter " & _
                "       (desp, fieldname, type, flag) " & _
                "values ('Out Time','n_deptime','N','U')"
        CON.Execute g_Sql
        
        g_Sql = "Insert into pr_clock_filter " & _
                "       (desp, fieldname, type, flag) " & _
                "values ('Early Hrs','n_earlhrs','N','U')"
        CON.Execute g_Sql
        
        g_Sql = "Insert into pr_clock_filter " & _
                "       (desp, fieldname, type, flag) " & _
                "values ('Permission Hrs','n_permhrs','N','U')"
        CON.Execute g_Sql
        
        g_Sql = "Insert into pr_clock_filter " & _
                "       (desp, fieldname, type, flag) " & _
                "values ('Work Hrs','n_workhrs','N','U')"
        CON.Execute g_Sql

        g_Sql = "Insert into pr_clock_filter " & _
                "       (desp, fieldname, type, flag) " & _
                "values ('OT Total','n_overtime','N','U')"
        CON.Execute g_Sql

        g_Sql = "Insert into pr_clock_filter " & _
                "       (desp, fieldname, type, flag) " & _
                "values ('Present Period','n_present','N','U')"
        CON.Execute g_Sql

        g_Sql = "Insert into pr_clock_filter " & _
                "       (desp, fieldname, type, flag) " & _
                "values ('Status','c_presabs','C','U')"
        CON.Execute g_Sql

        g_Sql = "Insert into pr_clock_filter " & _
                "       (desp, fieldname, type, flag) " & _
                "values ('No. of Clocking','n_noclock','N','U')"
        CON.Execute g_Sql

        g_Sql = "Insert into pr_clock_filter " & _
                "       (desp, fieldname, type, flag) " & _
                "values ('User Name','a.c_usr_id','C','U')"
        CON.Execute g_Sql

        g_Sql = "Insert into pr_clock_filter " & _
                "       (desp, fieldname, type, flag) " & _
                "values ('User Entry','c_flag','C','U')"
        CON.Execute g_Sql

        g_Sql = "Insert into pr_clock_filter " & _
                "       (desp, fieldname, type, flag) " & _
                "values ('No. of Clocking in a Day','n_noclock_day','N','U')"
        CON.Execute g_Sql
        
        g_Sql = "Insert into pr_clock_filter " & _
                "       (desp, fieldname, type, flag) " & _
                "values ('OT In Approved','n_otin_str','S','U')"
        CON.Execute g_Sql
        
        g_Sql = "Insert into pr_clock_filter " & _
                "       (desp, fieldname, type, flag) " & _
                "values ('OT Out Approved','n_otout_str','S','U')"
        CON.Execute g_Sql
        
        g_Sql = "Insert into pr_clock_filter " & _
                "       (desp, fieldname, type, flag) " & _
                "values ('Staff Type','b.c_stafftype','C','U')"
        CON.Execute g_Sql
        
        g_Sql = "Insert into pr_clock_filter " & _
                "       (desp, fieldname, type, flag) " & _
                "values ('OT In','n_otin','N','U')"
        CON.Execute g_Sql
        
        g_Sql = "Insert into pr_clock_filter " & _
                "       (desp, fieldname, type, flag) " & _
                "values ('OT Out','n_otout','N','U')"
        CON.Execute g_Sql
        
        g_Sql = "Insert into pr_clock_filter " & _
                "       (desp, fieldname, type, flag) " & _
                "values ('OT 1.5','n_ot15','N','U')"
        CON.Execute g_Sql
        
        g_Sql = "Insert into pr_clock_filter " & _
                "       (desp, fieldname, type, flag) " & _
                "values ('OT 2.0','n_ot20','N','U')"
        CON.Execute g_Sql
        
        g_Sql = "Insert into pr_clock_filter " & _
                "       (desp, fieldname, type, flag) " & _
                "values ('OT 3.0','n_ot30','N','U')"
        CON.Execute g_Sql
        

     End If
  
 Exit Sub

Err_Save:
   MsgBox "Error on creating default company - " + Err.Description, vbCritical, "Critical"
End Sub


Public Sub Get_FixedRate_Name()
  Dim rsChk As New ADODB.Recordset
  Dim i As Integer
  
    Set rsChk = Nothing
    g_Sql = "select c_salary, c_payname from pr_paystructure_dtl where c_salary in ('SAL0002','SAL0003','SAL0004','SAL0005')"
    rsChk.Open g_Sql, CON, adOpenForwardOnly, adLockOptimistic
    For i = 1 To rsChk.RecordCount
       If rsChk("c_salary").Value = "SAL0002" Then
          g_FixedRateName2 = Is_Null(rsChk("c_payname").Value, False)
       ElseIf rsChk("c_salary").Value = "SAL0003" Then
          g_FixedRateName3 = Is_Null(rsChk("c_payname").Value, False)
       ElseIf rsChk("c_salary").Value = "SAL0004" Then
          g_FixedRateName4 = Is_Null(rsChk("c_payname").Value, False)
       ElseIf rsChk("c_salary").Value = "SAL0005" Then
          g_FixedRateName5 = Is_Null(rsChk("c_payname").Value, False)
       End If
       rsChk.MoveNext
    Next i
End Sub

Public Function GetKeyValueInOBDC(vReqValue As String) As String
  Dim i As Long
  Dim rc As Long
  Dim Hkey As Long
  Dim hDepth As Long
  Dim KeyValType As Long
  Dim tmpVal As String
  Dim KeyValSize As Long
        
    rc = RegOpenKey(HKEY_CURRENT_USER, "SOFTWARE\ODBC\ODBC.INI\" & g_OBDCUser, Hkey)
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
    
    tmpVal = String$(1024, 0)
    KeyValSize = 1024
    
    rc = RegQueryValueEx(Hkey, vReqValue, 0, KeyValType, tmpVal, KeyValSize)
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then
        tmpVal = Left(tmpVal, KeyValSize - 1)
    Else
        tmpVal = Left(tmpVal, KeyValSize)
    End If
    
    GetKeyValueInOBDC = tmpVal
    rc = RegCloseKey(Hkey)
  Exit Function
    
GetKeyError:
    GetKeyValueInOBDC = ""
    rc = RegCloseKey(Hkey)
End Function

