
IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[trcustom_Appointment_Working_Patterns_P&E]') AND xtype in (N'TR'))
	DROP TRIGGER [dbo].[trcustom_Appointment_Working_Patterns_P&E]
GO

CREATE TRIGGER [dbo].[trcustom_Appointment_Working_Patterns_P&E] ON [dbo].[tbuser_Appointment_Working_Patterns]
    AFTER INSERT
AS
BEGIN
    SET NOCOUNT ON;

	MERGE Working_Patterns AS wp
		USING (	SELECT a.ID_1, SUM(wp.Sunday_Hours_AM) + SUM(wp.Sunday_Hours_PM) AS Sunday
		, SUM(wp.Monday_Hours_AM) + SUM(wp.Monday_Hours_PM) AS Monday
		, SUM(wp.Tuesday_Hours_AM) + SUM(wp.Tuesday_Hours_PM) AS Tuesday
		, SUM(wp.Wednesday_Hours_AM) + SUM(wp.Wednesday_Hours_PM) AS Wednesday
		, SUM(wp.Thursday_Hours_AM) + SUM(wp.Thursday_Hours_PM) AS Thursday
		, SUM(wp.Friday_Hours_AM) + SUM(wp.Friday_Hours_PM) AS Friday
		, SUM(wp.Saturday_Hours_AM) + SUM(wp.Saturday_Hours_PM) AS Saturday
		FROM inserted i
			INNER JOIN Appointment_Working_Patterns wp ON wp.ID_3 = i.ID_3
			INNER JOIN Appointments a ON a.ID = i.ID_3
		GROUP BY a.ID_1)
	AS awp ON (wp.ID_1 = awp.ID_1) 
	WHEN NOT MATCHED BY TARGET
		THEN INSERT(ID_1, Sunday_Hours, Monday_Hours, Tuesday_Hours, Wednesday_Hours, Thursday_Hours, Friday_Hours, Saturday_Hours) 
			VALUES(awp.ID_1, awp.Sunday, awp.Monday, awp.Tuesday, awp.Wednesday, awp.Thursday, awp.Friday, awp.Saturday)
	WHEN MATCHED 
		THEN UPDATE SET Sunday_Hours = awp.Sunday, Monday_Hours = awp.Monday, Tuesday_Hours = awp.Tuesday, Wednesday_Hours = awp.Wednesday
			, Thursday_Hours = awp.Thursday, Friday_Hours = awp.Friday, Saturday_Hours = awp.Saturday;
		--WHEN NOT MATCHED BY SOURCE
	--    THEN DELETE ;

END