
IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[Custom_TrainingGapAnalysis]') AND xtype in (N'TR'))
	DROP TRIGGER [dbo].[Custom_TrainingGapAnalysis]
GO

CREATE TRIGGER [dbo].[Custom_TrainingGapAnalysis] ON [dbo].[tbuser_Personnel_Records]
    AFTER INSERT, UPDATE
AS
BEGIN
    SET NOCOUNT ON;

 	MERGE INTO [dbo].[Training_Needs] AS Target
		USING (SELECT (SELECT [ID] FROM inserted), jc.[Competency], GETDATE(), 'Gap Analysis', 'Mandatory', 'Essential', '', '', '', ''
			FROM [dbo].[Job_Competencies] jc
			WHERE jc.[ID_100] = (SELECT [ID] FROM [dbo].[Job_Records] jr WHERE jr.[Job_Title] = (SELECT [Job_Title] FROM inserted))
			AND jc.[Competency] NOT IN (SELECT [Course_Title] FROM [dbo].[Training_Booking] tb WHERE [ID_1] = (SELECT [ID] FROM inserted)
				AND (tb.[Certificate_Expiry_Date] IS NULL OR tb.[Certificate_Expiry_Date] > GETDATE())))
		AS Source ([ID_1], [Course_Title], [Date_Identified], [Entered_By], [Reason], [Desirable_Essential], [Approved_By], [Delivery_Method], [Development_Need], [Notes])
		ON Target.[ID_1] = Source.[ID_1] AND Target.[Course_Title] = Source.[Course_Title]
	WHEN NOT MATCHED BY TARGET THEN
		INSERT ([ID_1], [Course_Title], [Date_Identified], [Entered_By], [Reason], [Desirable_Essential], [Approved_By], [Delivery_Method], [Development_Need], [Notes])
		VALUES ([ID_1], [Course_Title], [Date_Identified], [Entered_By], [Reason], [Desirable_Essential], [Approved_By], [Delivery_Method], [Development_Need], [Notes])
	WHEN NOT MATCHED BY SOURCE AND Target.[Entered_By] = 'Gap Analysis' THEN DELETE;

 END
 
 GO
