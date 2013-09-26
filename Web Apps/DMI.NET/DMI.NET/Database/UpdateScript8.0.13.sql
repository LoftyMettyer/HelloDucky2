
IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[spASRIntOrgChart]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[spASRIntOrgChart];
GO

CREATE PROCEDURE [dbo].[spASRIntOrgChart] (@RootID int)
AS
BEGIN
       SET NOCOUNT ON;
       DECLARE @staff_number VARCHAR(MAX);
       DECLARE @today DATETIME = DATEADD(dd, 0, DATEDIFF(dd, 0,  getdate()));
 
       -- Fetch Absences from DB
       DECLARE @ids TABLE (id INT, TYPE VARCHAR(50), reason VARCHAR(50));
 
       INSERT @ids
       SELECT id_1, TYPE, reason FROM absence a WHERE a.Start_Date <= @today AND (End_Date >= @today OR isnull(End_Date, '') = '')
 
       -- Fetch Training bookings from DB
       DECLARE @trainingIDs TABLE (id INT, course_title VARCHAR(50))
       INSERT @trainingIDs
              SELECT id_1, course_title FROM Training_Booking
              WHERE Start_Date <= @today and (End_Date >= @today or ISNULL(End_Date, '') = '') 
 
       SELECT @staff_number = staff_number FROM Personnel_Records WHERE id=@RootID;
       
       WITH Emp_CTE AS (
              SELECT id, forenames, surname AS name, staff_number, line_Manager_staff_number, job_Title, 1 AS HierarchyLevel
                     FROM Personnel_Records
                     WHERE line_manager_staff_number = @staff_number
              UNION ALL
                     SELECT e.id, e.forenames, e.Surname, e.staff_number, e.Line_Manager_Staff_Number, e.Job_Title, ecte.HierarchyLevel + 1 AS HierarchyLevel
                     FROM Personnel_Records e
              INNER JOIN Emp_CTE ecte ON ecte.Staff_Number = e.Line_Manager_Staff_Number
       )
       
       SELECT p.*, a.type, a.reason, t.course_title FROM Emp_CTE p
       LEFT JOIN @ids a ON a.id = p.id
       LEFT JOIN @trainingIDs t ON t.id = p.ID
       ORDER BY hierarchylevel, Job_Title, name
END
 



 GO


DECLARE @sSQL nvarchar(MAX),
		@sGroup sysname,
		@sObject sysname,
		@sObjectType char(2);

/*---------------------------------------------*/
/* Ensure the required permissions are granted */
/*---------------------------------------------*/
DECLARE curObjects CURSOR LOCAL FAST_FORWARD FOR
SELECT sysobjects.name, sysobjects.xtype
FROM sysobjects
     INNER JOIN sysusers ON sysobjects.uid = sysusers.uid
WHERE (((sysobjects.xtype = 'p') AND (sysobjects.name LIKE 'sp_asr%' OR sysobjects.name LIKE 'spasr%'))
    OR ((sysobjects.xtype = 'u') AND (sysobjects.name LIKE 'asrsys%'))
    OR ((sysobjects.xtype = 'fn') AND (sysobjects.name LIKE 'udf_ASRFn%')))
    AND (sysusers.name = 'dbo')

OPEN curObjects
FETCH NEXT FROM curObjects INTO @sObject, @sObjectType
WHILE (@@fetch_status = 0)
BEGIN
    IF rtrim(@sObjectType) = 'P' OR rtrim(@sObjectType) = 'FN'
    BEGIN
        SET @sSQL = 'GRANT EXEC ON [' + @sObject + '] TO [ASRSysGroup]'
        EXEC(@sSQL)
    END
    ELSE
    BEGIN
        SET @sSQL = 'GRANT SELECT,INSERT,UPDATE,DELETE ON [' + @sObject + '] TO [ASRSysGroup]'
        EXEC(@sSQL)
    END

    FETCH NEXT FROM curObjects INTO @sObject, @sObjectType
END
CLOSE curObjects
DEALLOCATE curObjects
