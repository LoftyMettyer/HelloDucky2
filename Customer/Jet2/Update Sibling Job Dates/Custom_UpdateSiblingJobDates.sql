
IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[Custom_UpdateSiblingJobDates]') AND xtype in (N'TR'))
	DROP TRIGGER [dbo].[Custom_UpdateSiblingJobDates]
GO

CREATE TRIGGER [dbo].[Custom_UpdateSiblingJobDates] ON [dbo].[tbuser_Salary]
    AFTER INSERT, UPDATE
AS
BEGIN
    SET NOCOUNT ON;

	IF @@NESTLEVEL > 5 RETURN;

 	MERGE Salary AS s
		USING (SELECT TOP 1 s.ID, s.Job_Start_Date, i._deleted
				, CASE WHEN i._deleted IS NULL THEN i.Job_Start_Date - 1 ELSE i.Job_End_Date END AS [Job_End_Date]
			FROM Salary s
			INNER JOIN inserted i on i.ID_1 = s.ID_1
			WHERE s.Job_Start_Date < i.Job_Start_Date
			ORDER BY s.Job_End_Date DESC)
			AS i ON (i.ID = s.ID) 
	WHEN MATCHED
		THEN UPDATE SET Job_End_Date = i.Job_End_Date;

 END
 
 GO
