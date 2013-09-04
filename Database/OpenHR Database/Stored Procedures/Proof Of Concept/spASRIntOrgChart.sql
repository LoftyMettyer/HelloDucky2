CREATE PROCEDURE [dbo].[spASRIntOrgChart]
	@RootID int
AS
	BEGIN
	/*	DON'T SHIP - Harcoded data structures for npg_openhr8!!! */
	SET NOCOUNT ON;
	declare @staff_number varchar(max);

	SELECT @staff_number = staff_number from Personnel_Records where id=@RootID;
	
	WITH Emp_CTE AS (
	SELECT forenames, surname as name, staff_number, line_Manager_staff_number, job_Title, 1 as HierarchyLevel
	FROM Personnel_Records
	WHERE line_manager_staff_number = @staff_number
	UNION ALL
	SELECT e.forenames, e.Surname, e.staff_number, e.Line_Manager_Staff_Number, e.Job_Title, ecte.HierarchyLevel + 1 as HierarchyLevel
	FROM Personnel_Records e
	INNER JOIN Emp_CTE ecte ON ecte.Staff_Number = e.Line_Manager_Staff_Number
	)
	SELECT *
	FROM Emp_CTE
	order by hierarchylevel, Job_Title, name

	
END
