IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[trcustom_Post_P&E]') AND xtype in (N'TR'))
	DROP TRIGGER [dbo].[trcustom_Post_P&E]
GO

CREATE TRIGGER [dbo].[trcustom_Post_P&E] ON [dbo].[tbuser_Post_Records]
    AFTER INSERT
AS
BEGIN
    SET NOCOUNT ON;

	INSERT Post_Holiday_Schemes (ID_219, Effective_Date, Holiday_Scheme)
		SELECT i.ID, i.Effective_Date, jhs.Holiday_Scheme
			FROM inserted i
			INNER JOIN Job_Families jf ON jf.Job_Family = i.Job_Family
			INNER JOIN Job_Family_Holiday_Schemes jhs ON jhs.ID_215 = jf.ID AND (jhs.End_Date >= GETDATE() OR End_Date IS NULL);

	INSERT Post_OMP_Schemes (ID_219, Effective_Date, OMP_Scheme, Description)
		SELECT i.ID, i.Effective_Date, omp.OMP_Scheme, omp.Description
			FROM inserted i
			INNER JOIN Job_Families jf ON jf.Job_Family = i.Job_Family
			INNER JOIN Job_Family_OMP_Schemes omp ON omp.ID_215 = jf.ID AND (omp.End_Date >= GETDATE() OR End_Date IS NULL);

	INSERT Post_OSP_Schemes (ID_219, Effective_Date, OSP_Scheme, Description)
		SELECT i.ID, i.Effective_Date, osp.OSP_Scheme, osp.Description
			FROM inserted i
			INNER JOIN Job_Families jf ON jf.Job_Family = i.Job_Family
			INNER JOIN Job_Family_OSP_Schemes osp ON osp.ID_215 = jf.ID AND (osp.End_Date >= GETDATE() OR End_Date IS NULL);

	INSERT Post_Pension_Schemes (ID_219, Effective_Date, Pension_Scheme)
		SELECT i.ID, i.Effective_Date, pen.Pension_Scheme
			FROM inserted i
			INNER JOIN Job_Families jf ON jf.Job_Family = i.Job_Family
			INNER JOIN Job_Family_Pension_Schemes pen ON pen.ID_215 = jf.ID AND (pen.End_Date >= GETDATE() OR End_Date IS NULL);

	INSERT Post_Working_Patterns (ID_219, Effective_Date, Regional_ID, Absence_In, Day_Pattern, Sunday_Hours_AM, Sunday_Hours_PM, Monday_Hours_AM, Monday_Hours_PM
									, Tuesday_Hours_AM, Tuesday_Hours_PM, Wednesday_Hours_AM, Wednesday_Hours_PM, Thursday_Hours_AM, Thursday_Hours_PM
									, Friday_Hours_AM, Friday_Hours_PM, Saturday_Hours_AM, Saturday_Hours_PM)
		SELECT i.ID, i.Effective_Date, wp.Regional_ID, wp.Absence_In, wp.Day_Pattern, wp.Sunday_Hours_AM, wp.Sunday_Hours_PM, wp.Monday_Hours_AM,wp.Monday_Hours_PM
									, wp.Tuesday_Hours_AM, wp.Tuesday_Hours_PM, wp.Wednesday_Hours_AM, wp.Wednesday_Hours_PM, wp.Thursday_Hours_AM, wp.Thursday_Hours_PM
									, wp.Friday_Hours_AM, wp.Friday_Hours_PM, wp.Saturday_Hours_AM, wp.Saturday_Hours_PM
			FROM inserted i
			INNER JOIN Job_Families jf ON jf.Job_Family = i.Job_Family
			INNER JOIN Job_Family_Working_Patterns wp ON wp.ID_215 = jf.ID AND (wp.End_Date >= GETDATE() OR End_Date IS NULL);




END