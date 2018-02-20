Attribute VB_Name = "Cost_Spread"
' General Module. Using by all the modules
' module spread
Option Explicit
Option Base 0

'------------------------------------------------------------------------------------------
'
'               Created by Vijaya Kumar Kanneganti
'               Date    :       01-Sep-1999
'
'                   All the Constants delcared below are taken from the
'               Far Point Spread Constants File and any change in the
'               constants value will not give the desired result.
'------------------------------------------------------------------------------------------

' Action property settings
Public Const SS_ACTION_ACTIVE_CELL = 0
Public Const SS_ACTION_GOTO_CELL = 1
Public Const SS_ACTION_SELECT_BLOCK = 2
Public Const SS_ACTION_CLEAR = 3
Public Const SS_ACTION_DELETE_COL = 4
Public Const SS_ACTION_DELETE_ROW = 5
Public Const SS_ACTION_INSERT_COL = 6
Public Const SS_ACTION_INSERT_ROW = 7
Public Const SS_ACTION_RECALC = 11
Public Const SS_ACTION_CLEAR_TEXT = 12
Public Const SS_ACTION_PRINT = 13
Public Const SS_ACTION_DESELECT_BLOCK = 14
Public Const SS_ACTION_DSAVE = 15
Public Const SS_ACTION_SET_CELL_BORDER = 16
Public Const SS_ACTION_ADD_MULTISELBLOCK = 17
Public Const SS_ACTION_GET_MULTI_SELECTION = 18
Public Const SS_ACTION_COPY_RANGE = 19
Public Const SS_ACTION_MOVE_RANGE = 20
Public Const SS_ACTION_SWAP_RANGE = 21
Public Const SS_ACTION_CLIPBOARD_COPY = 22
Public Const SS_ACTION_CLIPBOARD_CUT = 23
Public Const SS_ACTION_CLIPBOARD_PASTE = 24
Public Const SS_ACTION_SORT = 25
Public Const SS_ACTION_COMBO_CLEAR = 26
Public Const SS_ACTION_COMBO_REMOVE = 27
Public Const SS_ACTION_RESET = 28
Public Const SS_ACTION_SEL_MODE_CLEAR = 29
Public Const SS_ACTION_VMODE_REFRESH = 30
Public Const SS_ACTION_SMARTPRINT = 32

' SelectBlockOptions property settings
Public Const SS_SELBLOCKOPT_COLS = 1
Public Const SS_SELBLOCKOPT_ROWS = 2
Public Const SS_SELBLOCKOPT_BLOCKS = 4
Public Const SS_SELBLOCKOPT_ALL = 8

' DAutoSize property settings
Public Const SS_AUTOSIZE_NO = 0
Public Const SS_AUTOSIZE_MAX_COL_WIDTH = 1
Public Const SS_AUTOSIZE_BEST_GUESS = 2

' BackColorStyle property settings
Public Const SS_BACKCOLORSTYLE_OVERGRID = 0
Public Const SS_BACKCOLORSTYLE_UNDERGRID = 1

' CellType property settings
Public Const SS_CELL_TYPE_DATE = 0
Public Const SS_CELL_TYPE_EDIT = 1
Public Const SS_CELL_TYPE_FLOAT = 2
Public Const SS_CELL_TYPE_INTEGER = 3
Public Const SS_CELL_TYPE_PIC = 4
Public Const SS_CELL_TYPE_STATIC_TEXT = 5
Public Const SS_CELL_TYPE_TIME = 6
Public Const SS_CELL_TYPE_BUTTON = 7
Public Const SS_CELL_TYPE_COMBOBOX = 8
Public Const SS_CELL_TYPE_PICTURE = 9
Public Const SS_CELL_TYPE_CHECKBOX = 10
Public Const SS_CELL_TYPE_OWNER_DRAWN = 11

' CellBorderType property settings
Public Const SS_BORDER_TYPE_NONE = 0
Public Const SS_BORDER_TYPE_OUTLINE = 16
Public Const SS_BORDER_TYPE_LEFT = 1
Public Const SS_BORDER_TYPE_RIGHT = 2
Public Const SS_BORDER_TYPE_TOP = 4
Public Const SS_BORDER_TYPE_BOTTOM = 8

' CellBorderStyle property settings
Public Const SS_BORDER_STYLE_DEFAULT = 0
Public Const SS_BORDER_STYLE_SOLID = 1
Public Const SS_BORDER_STYLE_DASH = 2
Public Const SS_BORDER_STYLE_DOT = 3
Public Const SS_BORDER_STYLE_DASH_DOT = 4
Public Const SS_BORDER_STYLE_DASH_DOT_DOT = 5
Public Const SS_BORDER_STYLE_BLANK = 6
Public Const SS_BORDER_STYLE_FINE_SOLID = 11
Public Const SS_BORDER_STYLE_FINE_DASH = 12
Public Const SS_BORDER_STYLE_FINE_DOT = 13
Public Const SS_BORDER_STYLE_FINE_DASH_DOT = 14
Public Const SS_BORDER_STYLE_FINE_DASH_DOT_DOT = 15

' ColHeaderDisplay and RowHeaderDisplay property settings
Public Const SS_HEADER_BLANK = 0
Public Const SS_HEADER_NUMBERS = 1
Public Const SS_HEADER_LETTERS = 2

' TypeCheckTextAlign property settings
Public Const SS_CHECKBOX_TEXT_LEFT = 0
Public Const SS_CHECKBOX_TEXT_RIGHT = 1

' CursorStyle property settings
Public Const SS_CURSOR_STYLE_USER_DEFINED = 0
Public Const SS_CURSOR_STYLE_DEFAULT = 1
Public Const SS_CURSOR_STYLE_ARROW = 2
Public Const SS_CURSOR_STYLE_DEFCOLRESIZE = 3
Public Const SS_CURSOR_STYLE_DEFROWRESIZE = 4

' CursorType property settings
Public Const SS_CURSOR_TYPE_DEFAULT = 0
Public Const SS_CURSOR_TYPE_COLRESIZE = 1
Public Const SS_CURSOR_TYPE_ROWRESIZE = 2
Public Const SS_CURSOR_TYPE_BUTTON = 3
Public Const SS_CURSOR_TYPE_GRAYAREA = 4
Public Const SS_CURSOR_TYPE_LOCKEDCELL = 5
Public Const SS_CURSOR_TYPE_COLHEADER = 6
Public Const SS_CURSOR_TYPE_ROWHEADER = 7

' OperationMode property settings
Public Const SS_OP_MODE_NORMAL = 0
Public Const SS_OP_MODE_READONLY = 1
Public Const SS_OP_MODE_ROWMODE = 2
Public Const SS_OP_MODE_SINGLE_SELECT = 3
Public Const SS_OP_MODE_MULTI_SELECT = 4
Public Const SS_OP_MODE_EXT_SELECT = 5

' SortKeyOrder property settings
Public Const SS_SORT_ORDER_NONE = 0
Public Const SS_SORT_ORDER_ASCENDING = 1
Public Const SS_SORT_ORDER_DESCENDING = 2

' SortBy property settings
Public Const SS_SORT_BY_ROW = 0
Public Const SS_SORT_BY_COL = 1

' UserResize property settings
Public Const SS_USER_RESIZE_COL = 1
Public Const SS_USER_RESIZE_ROW = 2

' UserResizeCol and UserResizeRow property settings
Public Const SS_USER_RESIZE_DEFAULT = 0
Public Const SS_USER_RESIZE_ON = 1
Public Const SS_USER_RESIZE_OFF = 2

' VScrollSpecialType property settings
Public Const SS_VSCROLLSPECIAL_NO_HOME_END = 1
Public Const SS_VSCROLLSPECIAL_NO_PAGE_UP_DOWN = 2
Public Const SS_VSCROLLSPECIAL_NO_LINE_UP_DOWN = 4

' Position property settings
Public Const SS_POSITION_UPPER_LEFT = 0
Public Const SS_POSITION_UPPER_CENTER = 1
Public Const SS_POSITION_UPPER_RIGHT = 2
Public Const SS_POSITION_CENTER_LEFT = 3
Public Const SS_POSITION_CENTER_CENTER = 4
Public Const SS_POSITION_CENTER_RIGHT = 5
Public Const SS_POSITION_BOTTOM_LEFT = 6
Public Const SS_POSITION_BOTTOM_CENTER = 7
Public Const SS_POSITION_BOTTOM_RIGHT = 8

' ScrollBars property settings
Public Const SS_SCROLLBAR_NONE = 0
Public Const SS_SCROLLBAR_H_ONLY = 1
Public Const SS_SCROLLBAR_V_ONLY = 2
Public Const SS_SCROLLBAR_BOTH = 3

' PrintOrientation property settings
Public Const SS_PRINTORIENT_DEFAULT = 0
Public Const SS_PRINTORIENT_PORTRAIT = 1
Public Const SS_PRINTORIENT_LANDSCAPE = 2

' PrintType property settings
Public Const SS_PRINT_ALL = 0
Public Const SS_PRINT_CELL_RANGE = 1
Public Const SS_PRINT_CURRENT_PAGE = 2
Public Const SS_PRINT_PAGE_RANGE = 3

' TypeButtonType property settings
Public Const SS_CELL_BUTTON_NORMAL = 0
Public Const SS_CELL_BUTTON_TWO_STATE = 1

' TypeButtonAlign property settings
Public Const SS_CELL_BUTTON_ALIGN_BOTTOM = 0
Public Const SS_CELL_BUTTON_ALIGN_TOP = 1
Public Const SS_CELL_BUTTON_ALIGN_LEFT = 2
Public Const SS_CELL_BUTTON_ALIGN_RIGHT = 3

' ButtonDrawMode property settings
Public Const SS_BDM_ALWAYS = 0
Public Const SS_BDM_CURRENT_CELL = 1
Public Const SS_BDM_CURRENT_COLUMN = 2
Public Const SS_BDM_CURRENT_ROW = 4

' TypeDateFormat property settings
Public Const SS_CELL_DATE_FORMAT_DDMONYY = 0
Public Const SS_CELL_DATE_FORMAT_DDMMYY = 1
Public Const SS_CELL_DATE_FORMAT_MMDDYY = 2
Public Const SS_CELL_DATE_FORMAT_YYMMDD = 3

' TypeEditCharCase property settings
Public Const SS_CELL_EDIT_CASE_LOWER_CASE = 0
Public Const SS_CELL_EDIT_CASE_NO_CASE = 1
Public Const SS_CELL_EDIT_CASE_UPPER_CASE = 2

' TypeEditCharSet property settings
Public Const SS_CELL_EDIT_CHAR_SET_ASCII = 0
Public Const SS_CELL_EDIT_CHAR_SET_ALPHA = 1
Public Const SS_CELL_EDIT_CHAR_SET_ALPHANUMERIC = 2
Public Const SS_CELL_EDIT_CHAR_SET_NUMERIC = 3

' TypeTextAlignVert property settings
Public Const SS_CELL_STATIC_V_ALIGN_BOTTOM = 0
Public Const SS_CELL_STATIC_V_ALIGN_CENTER = 1
Public Const SS_CELL_STATIC_V_ALIGN_TOP = 2

' TypeTime24Hour property settings
Public Const SS_CELL_TIME_12_HOUR_CLOCK = 0
Public Const SS_CELL_TIME_24_HOUR_CLOCK = 1

'Unit type
Public Const SS_CELL_UNIT_NORMAL = 0
Public Const SS_CELL_UNIT_VGA = 1
Public Const SS_CELL_UNIT_TWIPS = 2

' TypeHAlign property settings
Public Const SS_CELL_H_ALIGN_LEFT = 0
Public Const SS_CELL_H_ALIGN_RIGHT = 1
Public Const SS_CELL_H_ALIGN_CENTER = 2

' EditEnterAction property settings
Public Const SS_CELL_EDITMODE_EXIT_NONE = 0
Public Const SS_CELL_EDITMODE_EXIT_UP = 1
Public Const SS_CELL_EDITMODE_EXIT_DOWN = 2
Public Const SS_CELL_EDITMODE_EXIT_LEFT = 3
Public Const SS_CELL_EDITMODE_EXIT_RIGHT = 4
Public Const SS_CELL_EDITMODE_EXIT_NEXT = 5
Public Const SS_CELL_EDITMODE_EXIT_PREVIOUS = 6
Public Const SS_CELL_EDITMODE_EXIT_SAME = 7
Public Const SS_CELL_EDITMODE_EXIT_NEXTROW = 8

' Custom function parameter type used with CFGetParamInfo method
Public Const SS_VALUE_TYPE_LONG = 0
Public Const SS_VALUE_TYPE_DOUBLE = 1
Public Const SS_VALUE_TYPE_STR = 2
Public Const SS_VALUE_TYPE_CELL = 3
Public Const SS_VALUE_TYPE_RANGE = 4

' Custom function parameter status used with CFGetParamInfo method
Public Const SS_VALUE_STATUS_OK = 0
Public Const SS_VALUE_STATUS_ERROR = 1
Public Const SS_VALUE_STATUS_EMPTY = 2

' Reference style settings used with GetRefStyle/SetRefStyle methods
Public Const SS_REFSTYLE_DEFAULT = 0
Public Const SS_REFSTYLE_A1 = 1
Public Const SS_REFSTYLE_R1C1 = 2

' Options used with Flags parameter of AddCustomFunctionExt method
Public Const SS_CUSTFUNC_WANTCELLREF = 1
Public Const SS_CUSTFUNC_WANTRANGEREF = 2

'               -------------------------------------------------------------------------------
'                        This Function return Boolean stating wheather
'                        a particular row in a spread is null or not
'               -------------------------------------------------------------------------------
Function isRowNull(Spread As fpSpread, ByVal RowNum As Integer, Optional IgnoreCols As Variant) As Boolean
        Dim ColNotNull As Boolean, iIndex As Integer, blnIgnore As Boolean, iTemp As Integer
        For iIndex = 0 To Spread.MaxCols - 1
                blnIgnore = False
                If Not IsMissing(IgnoreCols) Then
                        For iTemp = 0 To UBound(IgnoreCols)
                                If IgnoreCols(iTemp) = iIndex Then
                                        blnIgnore = True:       Exit For
                                End If
                        Next iTemp
                End If
                If Not blnIgnore Then
                        Spread.Row = RowNum:                Spread.Col = iIndex + 1
                        ColNotNull = IIf(Len(Trim(Spread.Text)), True, False)
                        If ColNotNull Then isRowNull = False: Exit Function
                End If
        Next iIndex
        isRowNull = True
End Function

'               -------------------------------------------------------------------------------------------------------
'                        This Function return a integer value specifing where the duplicated
'                        value for a particular column/row value in the specified col
'               -------------------------------------------------------------------------------------------------------
Function IsColDupValues(Spread As fpSpread, ByVal ColNum As Integer, Optional RowNum As Variant) As Integer
        Dim iChk As Integer, iLoop As Integer, iStartRow As Integer, iEndRow As Integer
        Dim strChk As String, strLoop As String
        
        If Not (TypeOf Spread Is fpSpread) Then Exit Function
        If IsMissing(RowNum) Then
           iStartRow = 1:      iEndRow = Spread.DataRowCnt
        Else
           iStartRow = CInt(RowNum):   iEndRow = CInt(RowNum)
        End If
        For iChk = iStartRow To iEndRow
            Spread.Row = iChk:    Spread.Col = ColNum:   strChk = Spread.Text
            If Trim(strChk) <> "" Then
               For iLoop = 1 To Spread.DataRowCnt
                   If Not isRowNull(Spread, iLoop) Then
                       If iLoop <> iChk Then
                           Spread.Row = iLoop:    Spread.Col = ColNum:  strLoop = Spread.Text
                           If StrComp(strChk, strLoop) = 0 Then
                              IsColDupValues = iChk
                              MsgBox strChk, vbInformation, "Duplicate"
                              Exit Function
                           End If
                       End If
                   End If
               Next iLoop
            End If
        Next iChk
        IsColDupValues = 0
End Function

'               -------------------------------------------------------------------------------------------------------
'                        This Function return a integer value specifing where the duplicated
'                        value for a particular column/row value in the specified col
'               -------------------------------------------------------------------------------------------------------
Function IsColTwoDupValues(Spread As fpSpread, ByVal ColNum As Integer, ByVal ColNumTwo As Integer, Optional RowNum As Variant) As Integer
        Dim iChk As Integer, iLoop As Integer, iStartRow As Integer, iEndRow As Integer
        Dim strChk As String, strLoop As String
         
        If Not (TypeOf Spread Is fpSpread) Then Exit Function
        If IsMissing(RowNum) Then
           iStartRow = 1:      iEndRow = Spread.DataRowCnt
        Else
           iStartRow = CInt(RowNum):   iEndRow = CInt(RowNum)
        End If
        For iChk = iStartRow To iEndRow
            Spread.Row = iChk:    Spread.Col = ColNum:     strChk = Spread.Text
            Spread.Row = iChk:    Spread.Col = ColNumTwo:  strChk = strChk + Spread.Text
            If Trim(strChk) <> "" Then
               For iLoop = 1 To Spread.DataRowCnt
                   If Not isRowNull(Spread, iLoop) Then
                      If iLoop <> iChk Then
                         Spread.Row = iLoop:    Spread.Col = ColNum:     strLoop = Spread.Text
                             Spread.Row = iLoop:    Spread.Col = ColNumTwo:  strLoop = strLoop + Spread.Text
                         If StrComp(strChk, strLoop) = 0 Then
                            IsColTwoDupValues = iChk
                            MsgBox strChk, vbInformation, "Duplicate"
                            Exit Function
                         End If
                      End If
                   End If
               Next iLoop
            End If
        Next iChk
        IsColTwoDupValues = 0
End Function

'               -------------------------------------------------------------------------------------------------------------------------------------
'                        This Function returns the column number where a null value in the row is identified
'                        this function main purpose is to validate a row, while moveing focus from one row to other
'               ------------------------------------------------------------------------------------------------------------------------------------
Function IsRowColsNull(Spread As fpSpread, ByVal RowNum As Integer, Optional NullCols As Variant) As Integer
        Dim ColNull As Boolean, Mandatory As Boolean
        Dim iIndex As Integer, ColIndex As Integer
        
        If Not (TypeOf Spread Is fpSpread) Then Exit Function
        For iIndex = 1 To Spread.MaxCols - 1
                Spread.Row = RowNum:                Spread.Col = iIndex
                ColNull = IIf(Len(Trim(Spread.Text)), False, True)
                If ColNull Then
                        If IsMissing(NullCols) Then IsRowColsNull = iIndex:    Exit Function
                        Mandatory = True
                        For ColIndex = 0 To UBound(NullCols)
                                If iIndex = NullCols(ColIndex) Then
                                        Mandatory = False:              Exit For
                                End If
                        Next ColIndex
                        If Mandatory Then IsRowColsNull = iIndex: Exit Function
                End If
        Next iIndex
        IsRowColsNull = False
End Function

'               -------------------------------------------------------------------------------
'                        This Function return the sum of all the values in the specified column
'               -------------------------------------------------------------------------------
Function GetColumnTotal(Spread As fpSpread, ByVal ColNum As Integer) As Double
        Dim dblTotal As Double, iCounter As Integer, varValue As Variant
        For iCounter = 1 To Spread.MaxRows
                Call Spread.GetText(ColNum, iCounter, varValue)
                If Len(varValue) > 0 Then dblTotal = dblTotal + CDbl(varValue)
        Next iCounter
        GetColumnTotal = dblTotal
End Function

Function GetRowTotal(Spread As fpSpread, ByVal RowNum As Integer) As Double
        Dim dblTotal As Double, iCounter As Integer, varValue As Variant
        For iCounter = 2 To Spread.MaxCols
                Call Spread.GetText(iCounter, RowNum, varValue)
                dblTotal = dblTotal + IIf(Len(varValue) = 0, 0, varValue)
        Next iCounter
        GetRowTotal = dblTotal
End Function

'               -------------------------------------------------------------------------------
'                        The purpose of this function is to lock cell
'               -------------------------------------------------------------------------------
Function LockCells(Spread As fpSpread, blnLock As Boolean, row1 As Integer, col1 As Integer, row2 As Integer, col2 As Integer) As Integer
        Dim iRow As Integer, iCol As Integer, iCellCount As Integer, iCounter As Integer
        For iRow = row1 To row2
                For iCol = col1 To col2
                        Spread.Row = iRow:         Spread.Col = iCol
                        Spread.Lock = blnLock:     Spread.Protect = True
                Next iCol
                For iCounter = 1 To 10: DoEvents:       Next iCounter
        Next iRow
End Function

'               -------------------------------------------------------------------------------
'                        This Function Delete all the Rows from Row1 to Row2
'                        Always the Row1 should be greater than Row2
'               -------------------------------------------------------------------------------
Function DeleteRows(Spread As fpSpread, ByVal row1 As Integer, ByVal row2 As Integer, Optional AddRow As Variant) As Integer
        Dim iRow As Integer, iDelRows As Integer
        For iRow = row1 To row2
                If Spread.MaxRows <= 0 Then Exit Function
                Spread.Row = row1:                Spread.Action = 5
                If IsMissing(AddRow) Then AddRow = False
                iDelRows = iDelRows + 1
        Next iRow
        Spread.MaxRows = IIf(AddRow, Spread.MaxRows, Spread.MaxRows - iDelRows)
        DeleteRows = iDelRows
End Function

'               ----------------------------------------------------------------------------------------------------------------
'                        This Function Clears all the Columns from Col1 to Col2 for the
'                        specified row Range(Row1-Row2)
'                        Always the Col1 should be greater than Col2
'               ----------------------------------------------------------------------------------------------------------------
Function ClearCols(Spread As fpSpread, ByVal row1 As Integer, ByVal col1 As Integer, ByVal row2 As Integer, ByVal col2 As Integer) As Integer
        Spread.Row = row1:        Spread.row2 = row2
        Spread.Col = col1:        Spread.col2 = col2
        Spread.BlockMode = True
        Spread.Action = SS_ACTION_CLEAR
        Spread.BlockMode = False
        ClearCols = Abs(col2 - col1) + 1
End Function

'               -------------------------------------------------------------------------------
'                        This Function Delete all the Columns from Col1 to Col2
'                        Always the Col1 should be greater than Col2
'               -------------------------------------------------------------------------------
Function DeleteCols(Spread As fpSpread, col1 As Integer, col2 As Integer) As Integer
        Dim iCol As Integer, iDelCols As Integer
        For iCol = col1 To col2
                If Spread.MaxCols > 0 Then Exit Function
                Spread.Col = iCol:               Spread.Action = 5
                Spread.MaxCols = Spread.MaxCols - 1
                iDelCols = iDelCols + 1
        Next iCol
        DeleteCols = iDelCols
End Function

'               ----------------------------------------------------------------------------------------------------------
'                        This Procedure sets the font attributes for the spread header, if
'                        you wish to have different font parameters for the header and for the cells
'               ----------------------------------------------------------------------------------------------------------
Public Sub SpreadHeaderFont(Spread As fpSpread, ByVal FontName As String, ByVal size As Integer, ByVal Bold As Boolean)
        Dim iCol As Integer
        Spread.Row = 0
        For iCol = 1 To Spread.MaxCols
                Spread.Col = iCol
                Spread.FontName = FontName
                Spread.FontSize = size
                Spread.FontBold = Bold
        Next iCol
End Sub

Public Function HLookUp(Spread As fpSpread, ByVal Col As Integer, SearchVal As Variant) As Integer
        Dim iRow As Integer
        For iRow = 1 To Spread.MaxRows
                Spread.Row = iRow:                Spread.Col = Col
                If StrComp(Spread.Text, CStr(SearchVal)) = 0 Then
                        HLookUp = iRow:         Exit Function
                End If
        Next iRow
End Function

'               ----------------------------------------------------------------------------------------------------------
'                       Marks/UnMarks the specified row for delete
'                       Mark         :       Set the font attribute FontStrikeThrough to True
'                       UnMark    :       Set the font attribute FontStrikeThrough to False
'               ----------------------------------------------------------------------------------------------------------
Public Function MarkRowForDelete(ByVal Mark As Boolean, Spread As fpSpread, ByVal Row As Integer) As Boolean
        Dim iRow As Integer, iCol As Integer, sngRowHeight As Single
        sngRowHeight = Spread.RowHeight(Row)
        For iCol = 1 To Spread.MaxCols
                Spread.Row = Row:                Spread.Col = iCol:                Spread.FontStrikethru = Mark
        Next iCol
        Spread.RowHeight(Row) = sngRowHeight
        MarkRowForDelete = Mark
End Function

Public Function SelectComboString(Spread As fpSpread, ByVal strSearch As String, ByVal Row As Integer, ByVal Col As Integer, ByVal lngth As Integer) As Integer
        Dim intIndex As Integer, strList As String, strJustOne As String, intPos As Integer
        Spread.Col = Col:       Spread.Row = Row:        strList = Spread.TypeComboBoxList
        If Spread.TypeComboBoxCount = 0 Then Exit Function
        Do
                intIndex = intIndex + 1
                intPos = InStr(strList, vbTab):                strJustOne = Mid(strList, 1, intPos - 1)
                        If Right(strJustOne, lngth) = strSearch Then
                                SelectComboString = intIndex
                                Spread.TypeComboBoxCurSel = intIndex - 1
                                Exit Function
                        End If
                strList = Mid(strList, intPos + 1, Len(strList))
        Loop Until strList = ""
        SelectComboString = -1
End Function

Public Function SelComboString(Spread As fpSpread, ByVal strSearch As String, ByVal Row As Integer, ByVal Col As Integer, Optional PatternCheck As Variant) As Integer
        Dim intIndex As Integer, strList As String, strJustOne As String, intPos As Integer
        Spread.Col = Col:       Spread.Row = Row:        strList = Spread.TypeComboBoxList
        If Spread.TypeComboBoxCount = 0 Then Exit Function
        Do
                intIndex = intIndex + 1
                intPos = InStr(strList, vbTab):                strJustOne = Mid(strList, 1, intPos - 1)
                If IsMissing(PatternCheck) Then
                        If strJustOne = strSearch Then
                                SelComboString = intIndex
                                Spread.TypeComboBoxCurSel = intIndex - 1
                                Exit Function
                        End If
                ElseIf CBool(PatternCheck) = True Then
                        If strJustOne Like "*" & strSearch & "*" Then
                                SelComboString = intIndex
                                Spread.TypeComboBoxCurSel = intIndex - 1
                                Exit Function
                        End If
                End If
                strList = Mid(strList, intPos + 1, Len(strList))
        Loop Until strList = ""
        SelComboString = -1
End Function

Public Sub SpreadColSort(Spread As fpSpread, ByVal Col As Long, ByVal Row As Long)
  Dim AlreadySorted As Integer
     If Row = 0 Then
        Spread.Col = 1
        Spread.col2 = Spread.DataColCnt
        Spread.Row = 1
        Spread.row2 = Spread.DataRowCnt
        Spread.SortBy = 0
        Spread.SortKey(1) = Col
        AlreadySorted = Is_Null((Spread.GetColItemData(Col)), True)
        Spread.SortKeyOrder(1) = IIf(AlreadySorted = 0, 1, 2)
        Call Spread.SetColItemData(Col, IIf(AlreadySorted = 0, 1, 0))
        Spread.Action = SS_ACTION_SORT
     End If
End Sub

Public Sub SpreadCellDataClear(Spread As fpSpread, ByVal Row As Long, ByVal Col As Long)
     If Row > 0 Then
        Spread.Row = Row
        Spread.Col = Col
           If Spread.Lock = False Then
              Spread.Text = ""
           End If
     End If
End Sub

Public Sub SpreadInsertRow(Spread As fpSpread, ByVal Row As Long, Optional ByVal Sheet As Long)
     If Row > 0 Then
        Spread.Sheet = Sheet
        Spread.MaxRows = Spread.MaxRows
        Spread.Row = Row
        Spread.Action = 7
        Spread.Row = Row
        Spread.Col = 1
        Spread.Action = 0
     End If
End Sub

Public Sub SpreadDeleteRow(Spread As fpSpread, ByVal Row As Long, Optional ByVal Sheet As Long)
     If Row > 0 Then
        Spread.Sheet = Sheet
        Spread.Row = Row
        Spread.Col = 1
        Spread.Action = 5
     End If
End Sub


Public Sub SpreadBlockCopy(Spread As fpSpread, ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long, Optional vMsgSuppress As Boolean, Optional vRowCopy As Boolean, Optional vNotCopyLockCell As Boolean)
  Dim tmpStr As String, vConfFlag
  Dim i As Long, j As Long, k As Long
  
    If BlockRow > 0 Then
       If Not vMsgSuppress Then
          vConfFlag = MsgBox("Do you want to change", vbOKCancel + vbDefaultButton2, "Information")
       Else
          vConfFlag = vbOK
       End If
       If vConfFlag = vbOK Then
          If BlockCol = BlockCol2 Then  'Col copy
             Spread.Row = BlockRow
             Spread.Col = BlockCol
                tmpStr = Trim(Spread.Text)
              
             For i = BlockRow To IIf(BlockRow2 > Spread.DataRowCnt, Spread.DataRowCnt, BlockRow2)
                 Spread.Row = i
                 Spread.Col = BlockCol
                    If vNotCopyLockCell Then
                       If Spread.Lock = False Then
                          Spread.Text = tmpStr
                       End If
                    Else
                       Spread.Text = tmpStr
                    End If
             Next i
          ElseIf BlockRow = BlockRow2 And vRowCopy Then ' Row Copy
             Spread.Row = BlockRow
             Spread.Col = BlockCol
                tmpStr = Trim(Spread.Text)
              
             For i = BlockCol To BlockCol2
                 Spread.Row = BlockRow
                 Spread.Col = i
                    If Spread.Lock = False Then
                       Spread.Text = tmpStr
                    End If
             Next i
          ElseIf vRowCopy Then  ' Row & Col Copy.
             For i = BlockCol To BlockCol2
                 Spread.Row = BlockRow
                 Spread.Col = i
                    tmpStr = Trim(Spread.Text)
                 For j = BlockRow To IIf(BlockRow2 > Spread.DataRowCnt, Spread.DataRowCnt, BlockRow2)
                    Spread.Row = j
                    Spread.Col = i
                       If vNotCopyLockCell Then
                          If Spread.Lock = False Then
                             Spread.Text = tmpStr
                          End If
                       Else
                             Spread.Text = tmpStr
                       End If
                 Next j
             Next i
          End If
       End If
    End If
End Sub

Public Sub SpreadBlockMultiCopy(Spread As fpSpread, ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
  Dim tmpStr As String
  Dim i, j, k As Long
     
     If BlockRow > 0 Then
        If MsgBox("Do you want to change", vbOKCancel, "Information") = vbOK Then
           k = 0
           For i = BlockCol To BlockCol2
               Spread.Row = BlockRow
               Spread.Col = BlockCol + k
                  tmpStr = Trim(Spread.Text)
               For j = BlockRow To BlockRow2
                   Spread.Row = j
                   Spread.Col = BlockCol + k
                      If Trim(tmpStr) <> "" Then
                         Spread.Text = tmpStr
                      End If
                Next j
                k = k + 1
            Next i
        End If
     End If
End Sub

Public Function Spread_NumFormat(ByVal vValue As Variant, Optional vDecReq As Boolean, Optional NoOfDecimal As Integer)
   vValue = Is_Null(vValue, True)
   If vDecReq = False Then
      Spread_NumFormat = Format(IIf(CDbl(vValue) = 0, "", vValue), "###,###,###")
   Else
      Spread_NumFormat = Format(IIf(CDbl(vValue) = 0, "", vValue), g_Nformat)
   End If
End Function

Public Function Spread_IntFormat(ByVal vValue As Variant, Optional vCommaReq As Boolean)
   vValue = Is_Null(vValue, True)
   If vCommaReq = False Then
      Spread_IntFormat = IIf(CDbl(vValue) = 0, "", vValue)
   Else
      Spread_IntFormat = Format(IIf(CDbl(vValue) = 0, "", vValue), "###,###,###")
   End If
End Function


Public Sub SpreadScrollAllow(ByVal frm As Form, ByVal vFlag As Boolean)
On Error Resume Next
  Dim Ctl
  Dim i As Integer
  
    For Each Ctl In frm.Controls
        If TypeOf Ctl Is fpSpread Then
           Ctl.Enabled = vFlag
           For i = 1 To Ctl.MaxCols
               Ctl.Row = -1
               Ctl.Col = i
                  Ctl.Lock = True
           Next i
        End If
    Next
End Sub

Function Is_DateSpread(arg As Variant, Optional vShortDate As Boolean) As Variant
    If vShortDate Then
       Is_DateSpread = IIf(IsDate(arg), Format(arg, "dd/mm/yy"), "")
    Else
       Is_DateSpread = IIf(IsDate(arg), Format(arg, "dd/mm/yyyy"), "")
    End If
End Function

Public Sub Clear_Spread(ByVal Spread As fpSpread)
    Spread.BlockMode = True
    Spread.Col = 1
    Spread.col2 = Spread.DataColCnt
    Spread.Row = 1
    Spread.row2 = Spread.DataRowCnt
    Spread.Action = 3
    Spread.BlockMode = False
End Sub

Public Sub Spread_Row_Height(ByVal Spread As fpSpread, Optional ByVal vHeight As Integer, Optional ByVal vHeaderHeight As Integer)
  Dim i As Integer
    ' row height
    Spread.Row = 0
    Spread.Col = -1
       If vHeaderHeight = 0 Then
          Spread.RowHeight(0) = 25
       Else
          Spread.RowHeight(0) = vHeaderHeight
       End If
       
    'col height
    For i = 1 To Spread.MaxRows
        Spread.Row = i
        Spread.Col = -1
           If vHeight = 0 Then
              Spread.RowHeight(i) = 15
           Else
              Spread.RowHeight(i) = vHeight
           End If
    Next i
End Sub
