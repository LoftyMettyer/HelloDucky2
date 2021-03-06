
/* Required Tables

242 - Appointment Allowances
243 - Appointment Benefits
244 - Appointment Deductions

*/

DELETE FROM ASRSysTableTriggers
WHERE [TriggerID] IN (17, 18, 19);

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[Custom_AppointmentAllowances]') AND xtype in (N'TR'))
	DROP TRIGGER [dbo].[Custom_AppointmentAllowances]
GO
 
IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[Custom_AppointmentDeductions]') AND xtype in (N'TR'))
	DROP TRIGGER [dbo].[Custom_AppointmentDeductions]
GO
  
IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[Custom_AppointmentBenefits]') AND xtype in (N'TR'))
	DROP TRIGGER [dbo].[Custom_AppointmentBenefits]
GO

INSERT ASRSysTableTriggers (TriggerID, TableID, Name, CodePosition, IsSystem, Content) VALUES (17, 242, 'Maintain Allowance on Personnel Records', 1, 1, '    MERGE INTO [dbo].[tbuser_Allowances] AS Target
		USING (SELECT a.[ID_1], i.[Type], i.[Effective_Date], i.[End_Date], i.[Frequency], i.[Amount], i.[Currency], a.[Post_ID], i.[ApptAllID], i.[_deleted], i.[_deleteddate]
			FROM [dbo].[tbuser_Appointment_Allowances] aa
			INNER JOIN inserted i ON i.[ID] = aa.[ID]
			INNER JOIN [dbo].[Appointments] a ON a.[ID] = i.[ID_3]
			WHERE i.[ApptAllID] > 0)
		AS Source ([ID_1], [Type], [Effective_Date], [End_Date], [Frequency], [Amount], [Currency], [Post_ID], [ApptAllID], [_deleted], [_deleteddate])
		ON Target.[ID_1] = Source.[ID_1] AND Target.[ApptAllID] = Source.[ApptAllID] AND Target.[Post_ID] = Source.[Post_ID]
	WHEN MATCHED THEN
		UPDATE SET [Type] = Source.[Type], [Effective_Date] = Source.[Effective_Date], [End_Date] = Source.[End_Date], [Frequency] = Source.[Frequency], [Amount] = Source.[Amount], [Currency] = Source.[Currency]
			, [_deleted] = Source.[_deleted], [_deleteddate] = Source.[_deleteddate]
	WHEN NOT MATCHED BY TARGET THEN
		INSERT ([ID_1], [Type], [Effective_Date], [End_Date], [Frequency], [Amount], [Currency], [Post_ID], [ApptAllID])
		VALUES ([ID_1], [Type], [Effective_Date], [End_Date], [Frequency], [Amount], [Currency], [Post_ID], [ApptAllID]);')
  GO

 INSERT ASRSysTableTriggers (TriggerID, TableID, Name, CodePosition, IsSystem, Content) VALUES (18, 243, 'Maintain Benefits on Personnel Records', 1, 1, '    MERGE INTO [dbo].[tbuser_Benefits] AS Target
		USING (SELECT a.[ID_1], i.[Type], i.[Effective_Date], i.[End_Date], i.[Frequency], i.[Cost], i.[Currency], i.[Provider], i.[Level_of_Cover], i.[Reference_Number]
			, i.[Percentage_to_Employer], i.[Annual_Cost_to_Employer], a.[Post_ID], i.[ApptBenID], i.[_deleted], i.[_deleteddate]
			FROM [dbo].[tbuser_Appointment_Benefits] ab
			INNER JOIN inserted i ON i.[ID] = ab.[ID]
			INNER JOIN [dbo].[Appointments] a ON a.[ID] = i.[ID_3]
			WHERE i.[ApptBenID] > 0)
		AS Source ([ID_1], [Type], [Effective_Date], [End_Date], [Frequency], [Cost], [Currency], [Provider], [Level_of_Cover], [Reference_Number]
			, [Percentage_to_Employer], [Annual_Cost_to_Employer], [Post_ID], [ApptBenID], [_deleted], [_deleteddate])
		ON Target.[ID_1] = Source.[ID_1] AND Target.[ApptBenID] = Source.[ApptBenID] AND Target.[Post_ID] = Source.[Post_ID]
	WHEN MATCHED THEN
		UPDATE SET [Type] = Source.[Type], [Effective_Date] = Source.[Effective_Date], [End_Date] = Source.[End_Date], [Frequency] = Source.[Frequency], [Cost] = Source.[Cost], [Currency] = Source.[Currency]
			, [Provider] = Source.[Provider], [Level_of_Cover] = Source.[Level_of_Cover], [Reference_Number] = Source.[Reference_Number], [Percentage_to_Employer] = Source.[Percentage_to_Employer]
			, [Annual_Cost_to_Employer] = Source.[Annual_Cost_to_Employer], [_deleted] = Source.[_deleted], [_deleteddate] = Source.[_deleteddate]
	WHEN NOT MATCHED BY TARGET THEN
		INSERT ([ID_1], [Type], [Effective_Date], [End_Date], [Frequency], [Cost], [Currency], [Provider], [Level_of_Cover], [Reference_Number], [Percentage_to_Employer], [Annual_Cost_to_Employer], [Post_ID], [ApptBenID])
		VALUES ([ID_1], [Type], [Effective_Date], [End_Date], [Frequency], [Cost], [Currency], [Provider], [Level_of_Cover], [Reference_Number], [Percentage_to_Employer], [Annual_Cost_to_Employer], [Post_ID], [ApptBenID]);')
 GO

 INSERT ASRSysTableTriggers (TriggerID, TableID, Name, CodePosition, IsSystem, Content) VALUES (19, 244, 'Maintain Deductions on Personnel Records', 1, 1, '    MERGE INTO [dbo].[tbuser_Deductions] AS Target
		USING (SELECT a.[ID_1], i.[Type], i.[Effective_Date], i.[End_Date], i.[Frequency], i.[Amount], i.[Currency], i.[Reference], a.[Post_ID], i.[ApptDedID], i.[_deleted], i.[_deleteddate]
			FROM [dbo].[tbuser_Appointment_Deductions] ad
			INNER JOIN inserted i ON i.[ID] = ad.[ID]
			INNER JOIN [dbo].[Appointments] a ON a.[ID] = i.[ID_3]
			WHERE i.[ApptDedID] > 0)
		AS Source ([ID_1], [Type], [Effective_Date], [End_Date], [Frequency], [Amount], [Currency], [Reference], [Post_ID], [ApptDedID], [_deleted], [_deleteddate])
		ON Target.[ID_1] = Source.[ID_1] AND Target.[ApptDedID] = Source.[ApptDedID] AND Target.[Post_ID] = Source.[Post_ID]
	WHEN MATCHED THEN
		UPDATE SET [Type] = Source.[Type], [Effective_Date] = Source.[Effective_Date], [End_Date] = Source.[End_Date], [Frequency] = Source.[Frequency], [Amount] = Source.[Amount], [Currency] = Source.[Currency]
			, [Reference] = Source.[Reference], [_deleted] = Source.[_deleted], [_deleteddate] = Source.[_deleteddate]
	WHEN NOT MATCHED BY TARGET THEN
		INSERT ([ID_1], [Type], [Effective_Date], [End_Date], [Frequency], [Amount], [Currency], [Reference], [Post_ID], [ApptDedID])
		VALUES ([ID_1], [Type], [Effective_Date], [End_Date], [Frequency], [Amount], [Currency], [Reference], [Post_ID], [ApptDedID]);')
 GO