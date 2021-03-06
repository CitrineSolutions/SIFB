USE [HITHRPAY_SIFB]
GO


SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO



CREATE PROCEDURE [dbo].[HR_CLOCK_EMP_UPD]
       @d_frdate   DateTime, 		
       @d_todate   DateTime,	
       @c_company  Varchar(7),
       @c_branch   Varchar(50),
       @c_dept     Varchar(50),	
       @c_empno    Varchar(7)
AS

If IsNull(@c_company,'A')='A'	Set @c_company = Null
If IsNull(@c_branch,'A')='A'    Set @c_branch = Null
If IsNull(@c_dept,'A')='A'		Set @c_dept = Null
If IsNull(@c_empno,'A')='A'		Set @c_empno = Null

If @d_frdate Is Null 
   Set @d_frdate = (Select d_fromdate From pr_payperiod_dtl Where c_type = 'W' and 
                           n_period = (Select Min(n_period) From pr_payperiod_dtl Where c_type = 'W' and c_period_closed = 'N'))

If @d_todate Is Null 
   Set @d_todate = (Select d_todate From pr_payperiod_dtl Where c_type = 'W' and 
                           n_period = (Select Min(n_period) From pr_payperiod_dtl Where c_type = 'W' and c_period_closed = 'N'))

SET DATEFIRST 7

Declare Pr_Cursor Cursor Scroll Static For
       Select a.c_empno, a.c_shiftcode 
       From pr_emp_mst a 
       Where a.c_clockcard = '1' and a.c_rec_sta = 'A'  and
            (a.d_dol is null or (a.d_dol >= @d_frdate and a.d_dol <= @d_todate)) and
			a.c_company = IsNull(@c_company, a.c_company) and a.c_branch = IsNull(@c_branch, a.c_branch) and 
			a.c_dept = IsNull(@c_dept, a.c_dept) and a.c_empno = IsNull(@c_empno, a.c_empno) 
		
Open Pr_Cursor

Declare @c_shiftcode Varchar(3), @c_shift Varchar(5)

Declare @d_tmpDate as DateTime

Fetch Next From Pr_Cursor Into @c_empno, @c_shiftcode

WHILE @@FETCH_STATUS = 0
  BEGIN

	Set @d_tmpdate = @d_frdate	

	While @d_tmpdate <= @d_todate
		Begin
			Set @c_shift = (Select c_shiftcode From pr_shiftstructure_dtl
			Where c_code = @c_shiftcode and n_wkday = DatePart(dw,@d_tmpdate))   

			Update pr_clock_emp Set c_shift = @c_shift Where c_empno = @c_empno and d_date = @d_tmpdate

			If @@RowCount = 0 	
				Insert Into pr_clock_emp
							(c_empno, d_date, n_wkday, c_shift, c_presabs, n_present, c_flag) 	
					 Values (@c_empno, @d_tmpdate, DatePart(dw,@d_tmpdate), @c_shift, Case when @c_shift = 'WO' then 'WO' else 'A' End, 1, 'B')	


			Update pr_clock_emp 
				Set c_shift = @c_shift, c_presabs = Case @c_shift When 'WO' then @c_shift else c_presabs End
			Where c_empno = @c_empno and d_date = @d_tmpdate and c_flag <> 'U' and IsNull(c_sh_flag,'A') <> 'U'


			Set @d_tmpdate = @d_tmpdate + 1		
		End		

     Fetch Next From Pr_Cursor Into @c_empno, @c_shiftcode
  END

Close Pr_Cursor
Deallocate Pr_Cursor

	
-- Exec HR_CLOCK_EMP_UPD '2017-09-01', '2017-09-30', 'A', 'A', 'A', 'A'





















GO
/****** Object:  StoredProcedure [dbo].[HR_CLOCK_PRESABS_UPD]    Script Date: 28/01/2018 12:13:57 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO




CREATE PROCEDURE [dbo].[HR_CLOCK_PRESABS_UPD]
       @d_frdate   DateTime, 		
       @d_todate   DateTime,	
       @c_company  Varchar(7),
       @c_branch   Varchar(50),
       @c_dept     Varchar(50),	
       @c_empno    Varchar(7)

AS

If IsNull(@c_company,'A')='A' 
   Set @c_company = Null

If IsNull(@c_branch,'A')='A' 
   Set @c_branch = Null

If IsNull(@c_dept,'A')='A'
   Set @c_dept = Null

If IsNull(@c_empno,'A')='A'
   Set @c_empno = Null


-- Update 'P' Present
Update pr_clock_emp 
   Set c_presabs = 'P', n_present = 1 
From pr_clock_emp a, pr_emp_mst b 
Where a.c_empno = b.c_empno and a.c_flag = 'B' and a.n_workhrs > 0 and 
      a.d_date >= @d_frdate and a.d_date <= @d_todate and 
      b.c_company = IsNull(@c_company, b.c_company) and b.c_branch = IsNull(@c_branch, b.c_branch) and 
      b.c_dept = IsNull(@c_dept, b.c_dept) and b.c_empno = IsNull(@c_empno, b.c_empno) 


-- Update 'WO' work off
Update pr_clock_emp 
   Set c_presabs = 'WO', n_present = 1 
From pr_clock_emp a, pr_emp_mst b 
Where a.c_empno = b.c_empno and a.c_flag = 'B' and a.n_workhrs = 0 and a.c_presabs = 'A' and 
      a.d_date >= @d_frdate and a.d_date <= @d_todate and a.c_shift = 'WO'  and 
      b.c_company = IsNull(@c_company, b.c_company) and b.c_branch = IsNull(@c_branch, b.c_branch) and 
      b.c_dept = IsNull(@c_dept, b.c_dept) and b.c_empno = IsNull(@c_empno, b.c_empno) 


-- Update 'A' again in case of leave entry updates
Update pr_clock_emp 
   Set c_presabs = 'A', n_present = 1 
From pr_clock_emp a, pr_emp_mst b 
Where a.c_empno = b.c_empno and a.c_flag = 'B' and a.n_workhrs = 0 and a.n_arrtime = 0 and a.c_presabs <> 'A' and 
      a.d_date >= @d_frdate and a.d_date <= @d_todate and a.c_shift <> 'WO'  and 
      b.c_company = IsNull(@c_company, b.c_company) and b.c_branch = IsNull(@c_branch, b.c_branch) and 
      b.c_dept = IsNull(@c_dept, b.c_dept) and b.c_empno = IsNull(@c_empno, b.c_empno) 


-- Update 'ML' from master
Update pr_clock_emp 
   Set c_presabs = 'ML', n_present = 1 
From pr_clock_emp a, pr_emp_mst b 
Where a.c_empno = b.c_empno and a.c_flag = 'B' and n_arrtime = 0 and n_deptime = 0 and 
      b.d_mlfrom is not null and b.d_mlto is not null and
      a.d_date between b.d_mlfrom and b.d_mlto and a.d_date >= @d_frdate and a.d_date <= @d_todate and 
      b.c_company = IsNull(@c_company, b.c_company) and b.c_branch = IsNull(@c_branch, b.c_branch) and 
      b.c_dept = IsNull(@c_dept, b.c_dept) and b.c_empno = IsNull(@c_empno, b.c_empno) 


-- Update 'PH' Public Holiday 
Declare Pr_Cursor Cursor Scroll Static For

        Select d_phdate From pr_holiday_mst 
        Where c_rec_sta = 'A' and d_phdate >= @d_frdate and d_phdate <= @d_todate
	Order by d_phdate

Open Pr_Cursor

Declare @d_phdate DateTime

Fetch Next From Pr_Cursor Into @d_phdate

While @@Fetch_Status = 0
  Begin
     
     Update pr_clock_emp 
	Set c_presabs = 'PH', n_present = 1 
     From pr_clock_emp a, pr_emp_mst b 
     Where a.c_empno = b.c_empno and a.d_date = @d_phdate and
           b.c_company = IsNull(@c_company, b.c_company) and b.c_branch = IsNull(@c_branch, b.c_branch) and 
           b.c_dept = IsNull(@c_dept, b.c_dept) and b.c_empno = IsNull(@c_empno, b.c_empno) 

      Fetch Next From Pr_Cursor Into @d_phdate
  End

Close Pr_Cursor
Deallocate Pr_Cursor


-- Update 'WO' before date join and after date left and Actual Date left
Update pr_clock_emp 
   Set c_presabs = 'WO', n_present = 1 
From pr_clock_emp a, pr_emp_mst b 
Where a.c_empno = b.c_empno and a.n_workhrs = 0 and a.n_time1 = 0 and
      a.c_presabs <> 'WO' and (a.d_date < b.d_doj or a.d_date > b.d_dol) and
      b.c_company = IsNull(@c_company, b.c_company) and b.c_branch = IsNull(@c_branch, b.c_branch) and 
      b.c_dept = IsNull(@c_dept, b.c_dept) and b.c_empno = IsNull(@c_empno, b.c_empno) 
	



-- HR_CLOCK_PRESABS_UPD '2017-03-01', '2017-03-31', 'A', 'A', 'A', 'A'


GO

/****** Object:  StoredProcedure [dbo].[HR_LEAVE_ALLOT_PROC]    Script Date: 28/01/2018 12:13:57 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[HR_LEAVE_ALLOT_PROC] 
       @Year  Int
AS

If @Year < 2018
   Return

-- One Year Completed employee Leave Updation Process

Declare @YearBegin Varchar(10), @YearEnd Varchar(10)
   Set @YearBegin = ltrim(rtrim(str(@Year)))+'-01-01'
   Set @YearEnd = ltrim(rtrim(str(@Year)))+'-12-31'


-- Prev year
Update pr_emp_leave_dtl Set c_yrstatus = 'O' Where d_prfrom < @YearBegin and c_yrstatus <> 'O'

Declare PrLeave_Cursor Cursor Scroll Static For

	Select c_empno, c_emptype, c_daywork, c_stafftype, d_doj, c_dept
	From   pr_emp_mst
	Where  d_dol is null and c_rec_sta = 'A' and DateAdd(Year,1,d_doj) >= @YearBegin and 
               DateAdd(Year,1,d_doj) <= (Select min(d_todate) from pr_payperiod_dtl Where c_type = 'W' and Getdate() between d_fromdate and d_todate)
 	Order By c_empno

Open PrLeave_Cursor

Declare @c_empno Varchar(7), @c_emptype Varchar(25), @c_daywork Varchar(3), @c_stafftype Varchar(1),
	@d_doj DateTime, @c_dept Varchar(50)

Declare @VacationLeave Decimal(10,2), @SickLeave Decimal(10,2)
Declare @Allot Decimal(10,2)

Fetch Next From PrLeave_Cursor Into @c_empno, @c_emptype, @c_daywork, @c_stafftype, @d_doj, @c_dept

WHILE @@FETCH_STATUS = 0
	BEGIN
		Set @VacationLeave = 0  Set @SickLeave = 0	

		If DateAdd(Year,1,@d_doj) < @YearBegin 
   		   Set @d_doj = @YearBegin

		If @c_emptype <> 'C'
			Begin
				-- Sick Leave
				Set @SickLeave = Round((DateDiff(Day, DateAdd(Year,1,@d_doj), @YearEnd)*21.00)/365.00,1)

				If @SickLeave - Floor(@SickLeave) = 0 
					Set @SickLeave = @SickLeave
				Else IF @SickLeave - Floor(@SickLeave) > 0.5
					Set @SickLeave = Ceiling(@SickLeave)
				Else
					Set @SickLeave = Floor(@SickLeave)+ 0.5

				Update pr_emp_leave_dtl Set n_alloted = @SickLeave Where c_empno = @c_empno and c_leave = 'SL' and year(d_prfrom) = @Year
				If @@Rowcount = 0 
					Insert Into pr_emp_leave_dtl (c_empno, d_prfrom, c_leave, n_entitle, n_alloted, c_yrstatus) 
										  Values (@c_empno, dateadd(year,1,@d_doj), 'SL', @SickLeave, @SickLeave, 'C')

				-- Vacation Leave
				Select @Allot = n_allot
				From pr_leaveallot_dtl Where c_leave = 'VL' and 1 BETWEEN n_from AND n_to and 
					 n_yearfrom in (Select Max(n_yearfrom) From pr_leaveallot_mst Where n_yearfrom < year(@YearEnd))

				Set @VacationLeave = Round((DateDiff(Day, DateAdd(Year,1,@d_doj), @YearEnd)*@Allot)/365.00,1)

				If @VacationLeave - Floor(@VacationLeave) = 0 
					Set @VacationLeave = @VacationLeave
				Else IF @VacationLeave - Floor(@VacationLeave) > 0.5
					Set @VacationLeave = Ceiling(@VacationLeave)
				Else
					Set @VacationLeave = Floor(@VacationLeave)+ 0.5

				Update pr_emp_leave_dtl Set n_alloted = @VacationLeave Where c_empno = @c_empno and c_leave = 'VL' and year(d_prfrom) = @Year
				If @@Rowcount = 0 
					Insert Into pr_emp_leave_dtl (c_empno, d_prfrom, c_leave, n_entitle, n_alloted, c_yrstatus) 
										  Values (@c_empno, dateadd(year,1,@d_doj), 'VL', @VacationLeave, @VacationLeave, 'C')	
			End            

		Update pr_emp_leave_dtl Set n_clbal = (n_opbal+n_alloted)-(n_utilised+n_adjusted)
		Where c_empno = @c_empno and year(d_prfrom) = @Year

		Fetch Next From PrLeave_Cursor Into @c_empno, @c_emptype, @c_daywork, @c_stafftype, @d_doj, @c_dept
	END

Close PrLeave_Cursor
Deallocate PrLeave_Cursor

Delete From pr_emp_leave_dtl
From pr_emp_mst a, pr_emp_leave_dtl b
Where a.c_empno = b.c_empno and a.c_rec_sta = 'A' and year(b.d_prfrom) = @Year and a.d_dol is not null and b.d_prfrom >= a.d_dol




-- Exec HR_LEAVE_ALLOT_PROC 2018


GO
/****** Object:  StoredProcedure [dbo].[HR_LEAVE_ENTITLE_PROC]    Script Date: 28/01/2018 12:13:57 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


CREATE PROCEDURE [dbo].[HR_LEAVE_ENTITLE_PROC] 
       @Year  Int
AS

If @Year < 2018
   Return

-- Update leave entitle update proc

Declare @c_empno Varchar(7), @n_entitle Decimal(10,2), @n_utilised Decimal(10,2)
Declare @VacationLeave Decimal(10,2), @OthLeave Decimal(10,2)

IF @Year = 2018
	BEGIN
		Declare Opbal_Cursor Cursor Scroll Static For

			Select a.c_empno, b.n_entitle, 
					IsNull((select n_utilised from pr_emp_leave_clbal_dtl where c_empno = a.c_empno and c_leave = 'VL' and 
									year(d_clbalon) = @Year-1),0) n_utilised
			From   pr_emp_mst a, pr_emp_leave_dtl b
			Where  a.c_empno = b.c_empno and b.c_leave = 'VL' and year(b.d_prfrom) = @Year and a.c_rec_sta = 'A' 
			Order By a.c_empno

		Open Opbal_Cursor

		Fetch Next From Opbal_Cursor Into @c_empno, @n_entitle, @n_utilised

		While @@FETCH_STATUS = 0
			Begin

				Set @VacationLeave = Round((((365.00 - @n_utilised) * @n_entitle) / 365.00),1)

				If @VacationLeave - Floor(@VacationLeave) = 0 
					Set @VacationLeave = @VacationLeave
				Else IF @VacationLeave - Floor(@VacationLeave) > 0.5
					Set @VacationLeave = Ceiling(@VacationLeave)
				Else
					Set @VacationLeave = Floor(@VacationLeave)+ 0.5

				Update pr_emp_leave_dtl Set n_alloted = @VacationLeave 
				Where c_empno = @c_empno and c_leave = 'VL' and year(d_prfrom) = @Year

				Update pr_emp_leave_dtl Set n_clbal = (n_opbal+n_alloted)-(n_utilised+n_adjusted)
				Where c_empno = @c_empno and c_leave = 'VL' and year(d_prfrom) = @Year


				Fetch Next From Opbal_Cursor Into @c_empno, @n_entitle, @n_utilised
			End

		Close Opbal_Cursor
		Deallocate Opbal_Cursor
	END
ELSE
	BEGIN
		Declare PrLeave_Cursor Cursor Scroll Static For

			Select a.c_empno, b.n_entitle, 
					IsNull((select n_utilised from pr_emp_leave_dtl where c_empno = a.c_empno and c_leave = 'VL' and 
									year(d_prfrom) = @Year-1),0) n_utilised
			From   pr_emp_mst a, pr_emp_leave_dtl b
			Where  a.c_empno = b.c_empno and b.c_leave = 'VL' and year(b.d_prfrom) = @Year and a.c_rec_sta = 'A' 
			Order By a.c_empno

		Open PrLeave_Cursor

		Fetch Next From PrLeave_Cursor Into @c_empno, @n_entitle, @n_utilised

		While @@FETCH_STATUS = 0
			Begin
				Set @OthLeave = 0
				Set @OthLeave = IsNull((Select sum(n_present) From pr_clock_emp Where c_empno = @c_empno and 
				                 year(d_date) = @Year-1 and c_presabs not in ('P','A','WO','PH')),0)

				Set @VacationLeave = Round((((365.00 - (@n_utilised+@OthLeave)) * @n_entitle) / 365.00),1)

				If @VacationLeave - Floor(@VacationLeave) = 0 
					Set @VacationLeave = @VacationLeave
				Else IF @VacationLeave - Floor(@VacationLeave) > 0.5
					Set @VacationLeave = Ceiling(@VacationLeave)
				Else
					Set @VacationLeave = Floor(@VacationLeave)+ 0.5

				Update pr_emp_leave_dtl Set n_alloted = @VacationLeave 
				Where c_empno = @c_empno and c_leave = 'VL' and year(d_prfrom) = @Year

				Update pr_emp_leave_dtl Set n_clbal = (n_opbal+n_alloted)-(n_utilised+n_adjusted)
				Where c_empno = @c_empno and c_leave = 'VL' and year(d_prfrom) = @Year


				Fetch Next From PrLeave_Cursor Into @c_empno, @n_entitle, @n_utilised
			End

		Close PrLeave_Cursor
		Deallocate PrLeave_Cursor
	END


-- Exec HR_LEAVE_ENTITLE_PROC 2018


GO
/****** Object:  StoredProcedure [dbo].[HR_LEAVE_REUPDATE_PROC]    Script Date: 28/01/2018 12:13:57 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


CREATE PROCEDURE [dbo].[HR_LEAVE_REUPDATE_PROC] 
       @n_year as Int,
       @empno Varchar(7) = NULL
AS

If @n_year < 2018
   Return

-- Update opbal from old system

Declare @c_empno Varchar(7), @c_leave Varchar(7), @n_clbal Decimal (10,2), @n_approvedays Decimal (10,2), @n_adjusted Decimal(10,2) 
Set @c_empno = @empno

If @n_year = 2018
	Begin
		Declare Opbal_Cursor Cursor Scroll Static For

			Select c_empno, c_leave, sum(n_clbal) n_clbal 
			From   pr_emp_leave_clbal_dtl 
			Where  c_leave in ('LL','SL','VL') and year(d_clbalon) = (@n_year-1) 
			Group By c_empno, c_leave
			Order By c_empno, c_leave

		Open Opbal_Cursor

		Fetch Next From Opbal_Cursor Into @c_empno, @c_leave, @n_clbal

		While @@FETCH_STATUS = 0
		  Begin

			  Update pr_emp_leave_dtl Set n_opbal = @n_clbal
			  Where c_empno = @c_empno and c_leave = @c_leave and year(d_prfrom) = @n_year 

			  Fetch Next From Opbal_Cursor Into @c_empno, @c_leave, @n_clbal
  
		  End

		Update pr_emp_leave_dtl Set n_clbal = (n_opbal+n_alloted)-(n_utilised+n_adjusted)
		Where year(d_prfrom) = @n_year 

		Close Opbal_Cursor
		Deallocate Opbal_Cursor
	End



-- Clbal update for prev year 

Set @c_empno = @empno
Update pr_emp_leave_dtl Set n_clbal = (n_opbal+n_alloted)-(n_utilised+n_adjusted)
Where year(d_prfrom) = (@n_year-1) and c_empno = IsNull(@c_empno, c_empno)

-- Update prev year clbal to current year opbal 

Set @c_empno = @empno
Declare PrOpbal_Cursor Cursor Scroll Static For

	Select c_empno, c_leave, sum(n_clbal) n_clbal 
	From   pr_emp_leave_dtl 
	Where  c_leave in ('LL','SL','VL') and year(d_prfrom) = (@n_year-1) and c_empno = IsNull(@c_empno, c_empno)
	Group By c_empno, c_leave

Open PrOpbal_Cursor

Fetch Next From PrOpbal_Cursor Into @c_empno, @c_leave, @n_clbal

While @@FETCH_STATUS = 0
  Begin

      Update pr_emp_leave_dtl Set n_opbal = @n_clbal
      Where c_empno = @c_empno and c_leave = @c_leave and year(d_prfrom) = @n_year and c_empno = IsNull(@c_empno, c_empno)

	  Fetch Next From PrOpbal_Cursor Into @c_empno, @c_leave, @n_clbal
  
  End

Close PrOpbal_Cursor
Deallocate PrOpbal_Cursor

-- Leave re-update for current year

Set @c_empno = @empno
Update pr_emp_leave_dtl Set n_utilised = 0, n_adjusted = 0 
Where year(d_prfrom) = @n_year and c_empno = IsNull(@c_empno, c_empno)

Update pr_emp_leave_dtl Set n_clbal = (n_opbal+n_alloted)-(n_utilised+n_adjusted)
Where year(d_prfrom) = @n_year and c_empno = IsNull(@c_empno, c_empno)


Set @c_empno = @empno
Declare PrUtilised_Cursor Cursor Scroll Static For

	Select c_empno, c_presabs c_leave, Sum(n_present) n_approvedays
	From   pr_clock_emp
	Where  c_presabs in ('LL','SL','VL') and year(d_date) = @n_year and c_empno = IsNull(@c_empno, c_empno)
	Group By c_empno, c_presabs
	Order By c_empno, c_presabs

Open PrUtilised_Cursor

Fetch Next From PrUtilised_Cursor Into @c_empno, @c_leave, @n_approvedays

While @@FETCH_STATUS = 0
  Begin

      Update pr_emp_leave_dtl Set n_utilised = n_utilised + @n_approvedays
      Where c_empno = @c_empno and c_leave = @c_leave and year(d_prfrom) = @n_year and c_empno = IsNull(@c_empno, c_empno)

	  Update pr_emp_leave_dtl Set n_clbal = (n_opbal+n_alloted)-(n_utilised+n_adjusted)
	  Where year(d_prfrom) = @n_year and c_empno = IsNull(@c_empno, c_empno)

      Fetch Next From PrUtilised_Cursor Into @c_empno, @c_leave, @n_approvedays
  End

Close PrUtilised_Cursor
Deallocate PrUtilised_Cursor


-- Update for adjusted leaves

Set @c_empno = @empno
Declare PrAdj_Cursor Cursor Scroll Static For

	Select c_empno, c_leave, Sum(n_adjusted) n_adjusted
	From   pr_leave_adj
	Where  c_leave in ('LL','SL','VL') and c_rec_sta = 'A' and n_year = @n_year and c_empno = IsNull(@c_empno, c_empno)
	Group By c_empno, c_leave
	Order By c_empno, c_leave

Open PrAdj_Cursor

Fetch Next From PrAdj_Cursor Into @c_empno, @c_leave, @n_adjusted 

While @@FETCH_STATUS = 0
  Begin

      Update pr_emp_leave_dtl Set n_adjusted = n_adjusted + @n_adjusted
      Where c_empno = @c_empno and c_leave = @c_leave and year(d_prfrom) = @n_year and c_empno = IsNull(@c_empno, c_empno)
	  
	  Update pr_emp_leave_dtl Set n_clbal = (n_opbal+n_alloted)-(n_utilised+n_adjusted)
	  Where year(d_prfrom) = @n_year and c_empno = IsNull(@c_empno, c_empno)

      Fetch Next From PrAdj_Cursor Into @c_empno, @c_leave, @n_adjusted 
  End

Close PrAdj_Cursor
Deallocate PrAdj_Cursor



-- Exec HR_LEAVE_REUPDATE_PROC 2018 
-- Exec HR_LEAVE_REUPDATE_PROC 2018, '306'
-- Exec HR_LEAVE_REUPDATE_PROC 2018, NULL


GO
/****** Object:  StoredProcedure [dbo].[HR_LEAVE_YEAR_ALLOT_PROC]    Script Date: 28/01/2018 12:13:57 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

CREATE PROCEDURE [dbo].[HR_LEAVE_YEAR_ALLOT_PROC] 
       @Year  Int
AS

If @Year < 2018
   Return

-- Update leave allotment at the begining of the year. 

DECLARE @YearBegin Varchar(10)
   SET @YearBegin = ltrim(rtrim(str(@Year)))+'-01-01'

Update pr_emp_leave_dtl Set c_yrstatus = 'O' Where d_prfrom < @YearBegin and c_yrstatus <> 'O'

DECLARE PrLeave_Cursor CURSOR SCROLL STATIC FOR

	SELECT c_empno, c_emptype, c_daywork, c_stafftype, d_doj
	FROM   pr_emp_mst
	WHERE  d_doj+365 <= @YearBegin and (d_dol is null) and c_rec_sta = 'A' 
	ORDER BY c_empno

OPEN PrLeave_Cursor

DECLARE @c_empno Varchar(7), @c_emptype Varchar(25), @c_daywork Varchar(3), @c_stafftype Varchar(1), @d_doj DateTime

DECLARE @VacationLeave Decimal(10,2), @SickLeave Decimal(10,2)

FETCH NEXT FROM PrLeave_Cursor INTO @c_empno, @c_emptype, @c_daywork, @c_stafftype, @d_doj

WHILE @@FETCH_STATUS = 0
	BEGIN

		Set @SickLeave = 21
		Update pr_emp_leave_dtl Set n_entitle = @SickLeave, n_alloted = @SickLeave 
		Where c_empno = @c_empno and c_leave = 'SL' and year(d_prfrom) = @Year
		If @@Rowcount = 0 
			Insert Into pr_emp_leave_dtl (c_empno, d_prfrom, c_leave, n_entitle, n_alloted, c_yrstatus) 
									Values (@c_empno, @YearBegin, 'SL', @SickLeave, @SickLeave, 'C')

		Set @VacationLeave = 0
		Select @VacationLeave = n_allot
		From pr_leaveallot_dtl Where c_leave = 'VL' and datediff(year, @d_doj, @YearBegin) BETWEEN n_from AND n_to and 
				n_yearfrom in (Select Max(n_yearfrom) From pr_leaveallot_mst Where n_yearfrom < year(@YearBegin))

		Update pr_emp_leave_dtl Set n_entitle = @VacationLeave, n_alloted = @VacationLeave 
		Where c_empno = @c_empno and c_leave = 'VL' and year(d_prfrom) = @Year
		If @@Rowcount = 0 
			Insert Into pr_emp_leave_dtl (c_empno, d_prfrom, c_leave, n_entitle, n_alloted, c_yrstatus) 
									Values (@c_empno, @YearBegin, 'VL', @VacationLeave, @VacationLeave, 'C')	


		FETCH NEXT FROM PrLeave_Cursor INTO @c_empno, @c_emptype, @c_daywork, @c_stafftype, @d_doj
	END

CLOSE PrLeave_Cursor
DEALLOCATE PrLeave_Cursor



-- Exec HR_LEAVE_YEAR_ALLOT_PROC 2018


GO
/****** Object:  StoredProcedure [dbo].[HR_LEAVEENTRY_DTL_UPD]    Script Date: 28/01/2018 12:13:57 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER OFF
GO



CREATE PROCEDURE [dbo].[HR_LEAVEENTRY_DTL_UPD]
AS


Declare Pr_Cursor Cursor Scroll Static For
	Select b.c_empno, b.c_leave, b.d_leavefrom, b.d_leaveto, b.c_type
	From pr_leaveentry_mst a, pr_leaveentry_dtl b
	Where a.c_code = b.c_code and a.c_rec_sta = 'A' and 
		a.n_period > (Select max(n_period) From pr_payperiod_dtl where c_type = 'W' and c_period_closed = 'Y') 
	Order by c_empno, d_leavefrom
		
Open Pr_Cursor

Declare @c_empno Varchar(7), @c_leave Varchar(7), @d_leavefrom Datetime, @d_leaveto Datetime, @c_type Varchar(10)

Fetch Next From Pr_Cursor Into @c_empno, @c_leave, @d_leavefrom, @d_leaveto, @c_type

While @@FETCH_STATUS = 0
  Begin

	 If @c_type = 'AM' or @c_type = 'PM'
	    Begin
           Update pr_clock_emp Set c_presabs = @c_leave, n_present = 0.5 
	       Where n_arrtime = 0 and n_deptime = 0 and c_empno = @c_empno and d_date = @d_leavefrom 

           Update pr_clock_emp Set c_presabs = @c_leave 
   	       Where n_arrtime = 0 and n_deptime = 0 and c_empno = @c_empno and d_date >= dateadd(d,1,@d_leavefrom) and d_date <= @d_leaveto

		End
	 Else
	    Begin
           Update pr_clock_emp Set c_presabs = @c_leave 
   	       Where n_arrtime = 0 and n_deptime = 0 and c_empno = @c_empno and d_date >= @d_leavefrom and d_date <= @d_leaveto
        End
     
	 Fetch Next From Pr_Cursor Into @c_empno, @c_leave, @d_leavefrom, @d_leaveto, @c_type
  End

Close Pr_Cursor
Deallocate Pr_Cursor


-- Exec HR_LEAVEENTRY_DTL_UPD

GO
