
IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[Custom_InsertJobRole]') AND xtype in (N'TR'))
	DROP TRIGGER [dbo].[Custom_InsertJobRole]
GO

--SYSTEM MANAGER AUTOMATICALLY UPGRADED TO 4.3
CREATE TRIGGER [dbo].[Custom_InsertJobRole] ON [dbo].[tbuser_Job_Role]
    AFTER INSERT
AS
BEGIN
    SET NOCOUNT ON;


	DECLARE @ID_1 integer;
	SET @ID_1 = (SELECT [ID_1] FROM inserted);

	DECLARE @Courses TABLE
	(
		  [Course] varchar(50)
		, [Completed_Date] datetime
	);

	DECLARE @Modules TABLE
	(
		  [Module] varchar(50)
		, [Completed_Date] datetime
	);

	DECLARE @Qualifications TABLE
	(
		  [Qualification] varchar(30)
		, [Completed_Date] datetime
	);

	INSERT INTO @Courses ([Course], [Completed_Date])
	SELECT [Classroom_Course_1] AS Course, [Classroom_Course_Completed_Date_1] AS Completed_Date
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Classroom_Course_1] != '' AND [Classroom_Course_Completed_Date_1] IS NOT NULL
	UNION
	SELECT [Classroom_Course_2], [Classroom_Course_Completed_Date_2]
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Classroom_Course_2] != '' AND [Classroom_Course_Completed_Date_2] IS NOT NULL
	UNION
	SELECT [Classroom_Course_3], [Classroom_Course_Completed_Date_3]
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Classroom_Course_3] != '' AND [Classroom_Course_Completed_Date_3] IS NOT NULL
	UNION
	SELECT [Classroom_Course_4], [Classroom_Course_Completed_Date_4]
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Classroom_Course_4] != '' AND [Classroom_Course_Completed_Date_4] IS NOT NULL
	UNION
	SELECT [Classroom_Course_5], [Classroom_Course_Completed_Date_5]
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Classroom_Course_5] != '' AND [Classroom_Course_Completed_Date_5] IS NOT NULL
	UNION
	SELECT [Classroom_Course_6], [Classroom_Course_Completed_Date_6]
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Classroom_Course_6] != '' AND [Classroom_Course_Completed_Date_6] IS NOT NULL
	UNION
	SELECT [Classroom_Course_7], [Classroom_Course_Completed_Date_7]
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Classroom_Course_7] != '' AND [Classroom_Course_Completed_Date_7] IS NOT NULL
	UNION
	SELECT [Classroom_Course_8], [Classroom_Course_Completed_Date_8]
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Classroom_Course_8] != '' AND [Classroom_Course_Completed_Date_8] IS NOT NULL
	UNION
	SELECT [Classroom_Course_9], [Classroom_Course_Completed_Date_9]
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Classroom_Course_9] != '' AND [Classroom_Course_Completed_Date_9] IS NOT NULL
	UNION
	SELECT [Classroom_Course_10], [Classroom_Course_Completed_Date_10]
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Classroom_Course_10] != '' AND [Classroom_Course_Completed_Date_10] IS NOT NULL
	UNION
	SELECT [Classroom_Course_11], [Classroom_Course_Completed_Date_11]
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Classroom_Course_11] != '' AND [Classroom_Course_Completed_Date_11] IS NOT NULL
	UNION
	SELECT [Classroom_Course_12], [Classroom_Course_Completed_Date_12]
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Classroom_Course_12] != '' AND [Classroom_Course_Completed_Date_12] IS NOT NULL
	UNION
	SELECT [Classroom_Course_13], [Classroom_Course_Completed_Date_13]
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Classroom_Course_13] != '' AND [Classroom_Course_Completed_Date_13] IS NOT NULL
	UNION
	SELECT [Classroom_Course_14], [Classroom_Course_Completed_Date_14]
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Classroom_Course_14] != '' AND [Classroom_Course_Completed_Date_14] IS NOT NULL
	UNION
	SELECT [Classroom_Course_15], [Classroom_Course_Completed_Date_15]
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Classroom_Course_15] != '' AND [Classroom_Course_Completed_Date_15] IS NOT NULL
	UNION
	SELECT [Classroom_Course_16], [Classroom_Course_Completed_Date_16]
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Classroom_Course_16] != '' AND [Classroom_Course_Completed_Date_16] IS NOT NULL
	UNION
	SELECT [Classroom_Course_17], [Classroom_Course_Completed_Date_17]
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Classroom_Course_17] != '' AND [Classroom_Course_Completed_Date_17] IS NOT NULL
	UNION
	SELECT [Classroom_Course_18], [Classroom_Course_Completed_Date_18]
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Classroom_Course_18] != '' AND [Classroom_Course_Completed_Date_18] IS NOT NULL
	UNION
	SELECT [Classroom_Course_19], [Classroom_Course_Completed_Date_19]
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Classroom_Course_19] != '' AND [Classroom_Course_Completed_Date_19] IS NOT NULL
	UNION
	SELECT [Classroom_Course_20], [Classroom_Course_Completed_Date_20]
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Classroom_Course_20] != '' AND [Classroom_Course_Completed_Date_20] IS NOT NULL
	UNION
	SELECT [Classroom_Course_21], [Classroom_Course_Completed_Date_21]
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Classroom_Course_21] != '' AND [Classroom_Course_Completed_Date_21] IS NOT NULL
	UNION
	SELECT [Classroom_Course_22], [Classroom_Course_Completed_Date_22]
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Classroom_Course_22] != '' AND [Classroom_Course_Completed_Date_22] IS NOT NULL
	UNION
	SELECT [Classroom_Course_23], [Classroom_Course_Completed_Date_23]
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Classroom_Course_23] != '' AND [Classroom_Course_Completed_Date_23] IS NOT NULL
	UNION
	SELECT [Classroom_Course_24], [Classroom_Course_Completed_Date_24]
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Classroom_Course_24] != '' AND [Classroom_Course_Completed_Date_24] IS NOT NULL
	UNION
	SELECT [Classroom_Course_25], [Classroom_Course_Completed_Date_25]
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Classroom_Course_25] != '' AND [Classroom_Course_Completed_Date_25] IS NOT NULL
	ORDER BY Course, Completed_Date;


	INSERT INTO @Modules ([Module], [Completed_Date])
	SELECT [Evolve_Module_1] AS Module, [Evolve_Module_Completed_Date_1] AS Completed_Date
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Evolve_Module_1] != '' AND [Evolve_Module_Completed_Date_1] IS NOT NULL
	UNION
	SELECT [Evolve_Module_2] AS Course, [Evolve_Module_Completed_Date_2] AS Completed_Date
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Evolve_Module_2] != '' AND [Evolve_Module_Completed_Date_2] IS NOT NULL
	UNION
	SELECT [Evolve_Module_3] AS Course, [Evolve_Module_Completed_Date_3] AS Completed_Date
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Evolve_Module_3] != '' AND [Evolve_Module_Completed_Date_3] IS NOT NULL
	UNION
	SELECT [Evolve_Module_4] AS Course, [Evolve_Module_Completed_Date_4] AS Completed_Date
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Evolve_Module_4] != '' AND [Evolve_Module_Completed_Date_4] IS NOT NULL
	UNION
	SELECT [Evolve_Module_5] AS Course, [Evolve_Module_Completed_Date_5] AS Completed_Date
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Evolve_Module_5] != '' AND [Evolve_Module_Completed_Date_5] IS NOT NULL
	UNION
	SELECT [Evolve_Module_6] AS Course, [Evolve_Module_Completed_Date_6] AS Completed_Date
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Evolve_Module_6] != '' AND [Evolve_Module_Completed_Date_6] IS NOT NULL
	UNION
	SELECT [Evolve_Module_7] AS Course, [Evolve_Module_Completed_Date_7] AS Completed_Date
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Evolve_Module_7] != '' AND [Evolve_Module_Completed_Date_7] IS NOT NULL
	UNION
	SELECT [Evolve_Module_8] AS Course, [Evolve_Module_Completed_Date_8] AS Completed_Date
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Evolve_Module_8] != '' AND [Evolve_Module_Completed_Date_8] IS NOT NULL
	UNION
	SELECT [Evolve_Module_9] AS Course, [Evolve_Module_Completed_Date_9] AS Completed_Date
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Evolve_Module_9] != '' AND [Evolve_Module_Completed_Date_9] IS NOT NULL
	UNION
	SELECT [Evolve_Module_10] AS Course, [Evolve_Module_Completed_Date_10] AS Completed_Date
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Evolve_Module_10] != '' AND [Evolve_Module_Completed_Date_10] IS NOT NULL
	UNION
	SELECT [Evolve_Module_11] AS Course, [Evolve_Module_Completed_Date_11] AS Completed_Date
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Evolve_Module_11] != '' AND [Evolve_Module_Completed_Date_11] IS NOT NULL
	UNION
	SELECT [Evolve_Module_12] AS Course, [Evolve_Module_Completed_Date_12] AS Completed_Date
	FROM [dbo].[tbuser_Job_Role]
	WHERE [Evolve_Module_12] != '' AND [Evolve_Module_Completed_Date_12] IS NOT NULL
	UNION
	SELECT [Evolve_Module_13] AS Course, [Evolve_Module_Completed_Date_13] AS Completed_Date
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Evolve_Module_13] != '' AND [Evolve_Module_Completed_Date_13] IS NOT NULL
	UNION
	SELECT [Evolve_Module_14] AS Course, [Evolve_Module_Completed_Date_14] AS Completed_Date
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Evolve_Module_14] != '' AND [Evolve_Module_Completed_Date_14] IS NOT NULL
	UNION
	SELECT [Evolve_Module_15] AS Course, [Evolve_Module_Completed_Date_15] AS Completed_Date
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Evolve_Module_15] != '' AND [Evolve_Module_Completed_Date_15] IS NOT NULL
	UNION
	SELECT [Evolve_Module_16] AS Course, [Evolve_Module_Completed_Date_16] AS Completed_Date
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Evolve_Module_16] != '' AND [Evolve_Module_Completed_Date_16] IS NOT NULL
	UNION
	SELECT [Evolve_Module_17] AS Course, [Evolve_Module_Completed_Date_17] AS Completed_Date
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Evolve_Module_17] != '' AND [Evolve_Module_Completed_Date_17] IS NOT NULL
	UNION
	SELECT [Evolve_Module_18] AS Course, [Evolve_Module_Completed_Date_18] AS Completed_Date
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Evolve_Module_18] != '' AND [Evolve_Module_Completed_Date_18] IS NOT NULL
	UNION
	SELECT [Evolve_Module_19] AS Course, [Evolve_Module_Completed_Date_19] AS Completed_Date
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Evolve_Module_19] != '' AND [Evolve_Module_Completed_Date_19] IS NOT NULL
	UNION
	SELECT [Evolve_Module_20] AS Course, [Evolve_Module_Completed_Date_20] AS Completed_Date
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Evolve_Module_20] != '' AND [Evolve_Module_Completed_Date_20] IS NOT NULL
	UNION
	SELECT [Evolve_Module_21] AS Course, [Evolve_Module_Completed_Date_21] AS Completed_Date
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Evolve_Module_21] != '' AND [Evolve_Module_Completed_Date_21] IS NOT NULL
	UNION
	SELECT [Evolve_Module_22] AS Course, [Evolve_Module_Completed_Date_22] AS Completed_Date
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Evolve_Module_22] != '' AND [Evolve_Module_Completed_Date_22] IS NOT NULL
	UNION
	SELECT [Evolve_Module_23] AS Course, [Evolve_Module_Completed_Date_23] AS Completed_Date
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Evolve_Module_23] != '' AND [Evolve_Module_Completed_Date_23] IS NOT NULL
	UNION
	SELECT [Evolve_Module_24] AS Course, [Evolve_Module_Completed_Date_24] AS Completed_Date
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Evolve_Module_24] != '' AND [Evolve_Module_Completed_Date_24] IS NOT NULL
	UNION
	SELECT [Evolve_Module_25] AS Course, [Evolve_Module_Completed_Date_25] AS Completed_Date
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Evolve_Module_25] != '' AND [Evolve_Module_Completed_Date_25] IS NOT NULL
	ORDER BY Module, Completed_Date;


	INSERT INTO @Qualifications ([Qualification], [Completed_Date])
	SELECT [Qualification_1] AS Qualification, [Qualification_Completed_Date_1] AS Completed_Date
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Qualification_1] != '' AND [Qualification_Completed_Date_1] IS NOT NULL
	UNION
	SELECT [Qualification_2] AS Qualification, [Qualification_Completed_Date_2] AS Completed_Date
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Qualification_2] != '' AND [Qualification_Completed_Date_2] IS NOT NULL
	UNION
	SELECT [Qualification_3] AS Qualification, [Qualification_Completed_Date_3] AS Completed_Date
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Qualification_3] != '' AND [Qualification_Completed_Date_3] IS NOT NULL
	UNION
	SELECT [Qualification_4] AS Qualification, [Qualification_Completed_Date_4] AS Completed_Date
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Qualification_4] != '' AND [Qualification_Completed_Date_4] IS NOT NULL
	UNION
	SELECT [Qualification_5] AS Qualification, [Qualification_Completed_Date_5] AS Completed_Date
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Qualification_5] != '' AND [Qualification_Completed_Date_5] IS NOT NULL
	UNION
	SELECT [Qualification_6] AS Qualification, [Qualification_Completed_Date_6] AS Completed_Date
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Qualification_6] != '' AND [Qualification_Completed_Date_6] IS NOT NULL
	UNION
	SELECT [Qualification_7] AS Qualification, [Qualification_Completed_Date_7] AS Completed_Date
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Qualification_7] != '' AND [Qualification_Completed_Date_7] IS NOT NULL
	UNION
	SELECT [Qualification_8] AS Qualification, [Qualification_Completed_Date_8] AS Completed_Date
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Qualification_8] != '' AND [Qualification_Completed_Date_8] IS NOT NULL
	UNION
	SELECT [Qualification_9] AS Qualification, [Qualification_Completed_Date_9] AS Completed_Date
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Qualification_9] != '' AND [Qualification_Completed_Date_9] IS NOT NULL
	UNION
	SELECT [Qualification_10] AS Qualification, [Qualification_Completed_Date_10] AS Completed_Date
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Qualification_10] != '' AND [Qualification_Completed_Date_10] IS NOT NULL
	UNION
	SELECT [Qualification_11] AS Qualification, [Qualification_Completed_Date_11] AS Completed_Date
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Qualification_11] != '' AND [Qualification_Completed_Date_11] IS NOT NULL
	UNION
	SELECT [Qualification_12] AS Qualification, [Qualification_Completed_Date_12] AS Completed_Date
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Qualification_12] != '' AND [Qualification_Completed_Date_12] IS NOT NULL
	UNION
	SELECT [Qualification_13] AS Qualification, [Qualification_Completed_Date_13] AS Completed_Date
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Qualification_13] != '' AND [Qualification_Completed_Date_13] IS NOT NULL
	UNION
	SELECT [Qualification_14] AS Qualification, [Qualification_Completed_Date_14] AS Completed_Date
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Qualification_14] != '' AND [Qualification_Completed_Date_14] IS NOT NULL
	UNION
	SELECT [Qualification_15] AS Qualification, [Qualification_Completed_Date_15] AS Completed_Date
	FROM [dbo].[tbuser_Job_Role]
	WHERE [ID_1] = @ID_1 AND [Qualification_15] != '' AND [Qualification_Completed_Date_15] IS NOT NULL
	ORDER BY Qualification, Completed_Date;


	UPDATE [dbo].[tbuser_Job_Role]
	SET   [Classroom_Course_Completed_Date_1] = (SELECT TOP 1 [Completed_Date] FROM @Courses WHERE [Course] = base.[Classroom_Course_1])
		, [Classroom_Course_Completed_Date_2] = (SELECT TOP 1 [Completed_Date] FROM @Courses WHERE [Course] = base.[Classroom_Course_2])
		, [Classroom_Course_Completed_Date_3] = (SELECT TOP 1 [Completed_Date] FROM @Courses WHERE [Course] = base.[Classroom_Course_3])
		, [Classroom_Course_Completed_Date_4] = (SELECT TOP 1 [Completed_Date] FROM @Courses WHERE [Course] = base.[Classroom_Course_4])
		, [Classroom_Course_Completed_Date_5] = (SELECT TOP 1 [Completed_Date] FROM @Courses WHERE [Course] = base.[Classroom_Course_5])
		, [Classroom_Course_Completed_Date_6] = (SELECT TOP 1 [Completed_Date] FROM @Courses WHERE [Course] = base.[Classroom_Course_6])
		, [Classroom_Course_Completed_Date_7] = (SELECT TOP 1 [Completed_Date] FROM @Courses WHERE [Course] = base.[Classroom_Course_7])
		, [Classroom_Course_Completed_Date_8] = (SELECT TOP 1 [Completed_Date] FROM @Courses WHERE [Course] = base.[Classroom_Course_8])
		, [Classroom_Course_Completed_Date_9] = (SELECT TOP 1 [Completed_Date] FROM @Courses WHERE [Course] = base.[Classroom_Course_9])
		, [Classroom_Course_Completed_Date_10] = (SELECT TOP 1 [Completed_Date] FROM @Courses WHERE [Course] = base.[Classroom_Course_10])
		, [Classroom_Course_Completed_Date_11] = (SELECT TOP 1 [Completed_Date] FROM @Courses WHERE [Course] = base.[Classroom_Course_11])
		, [Classroom_Course_Completed_Date_12] = (SELECT TOP 1 [Completed_Date] FROM @Courses WHERE [Course] = base.[Classroom_Course_12])
		, [Classroom_Course_Completed_Date_13] = (SELECT TOP 1 [Completed_Date] FROM @Courses WHERE [Course] = base.[Classroom_Course_13])
		, [Classroom_Course_Completed_Date_14] = (SELECT TOP 1 [Completed_Date] FROM @Courses WHERE [Course] = base.[Classroom_Course_14])
		, [Classroom_Course_Completed_Date_15] = (SELECT TOP 1 [Completed_Date] FROM @Courses WHERE [Course] = base.[Classroom_Course_15])
		, [Classroom_Course_Completed_Date_16] = (SELECT TOP 1 [Completed_Date] FROM @Courses WHERE [Course] = base.[Classroom_Course_16])
		, [Classroom_Course_Completed_Date_17] = (SELECT TOP 1 [Completed_Date] FROM @Courses WHERE [Course] = base.[Classroom_Course_17])
		, [Classroom_Course_Completed_Date_18] = (SELECT TOP 1 [Completed_Date] FROM @Courses WHERE [Course] = base.[Classroom_Course_18])
		, [Classroom_Course_Completed_Date_19] = (SELECT TOP 1 [Completed_Date] FROM @Courses WHERE [Course] = base.[Classroom_Course_19])
		, [Classroom_Course_Completed_Date_20] = (SELECT TOP 1 [Completed_Date] FROM @Courses WHERE [Course] = base.[Classroom_Course_20])
		, [Classroom_Course_Completed_Date_21] = (SELECT TOP 1 [Completed_Date] FROM @Courses WHERE [Course] = base.[Classroom_Course_21])
		, [Classroom_Course_Completed_Date_22] = (SELECT TOP 1 [Completed_Date] FROM @Courses WHERE [Course] = base.[Classroom_Course_22])
		, [Classroom_Course_Completed_Date_23] = (SELECT TOP 1 [Completed_Date] FROM @Courses WHERE [Course] = base.[Classroom_Course_23])
		, [Classroom_Course_Completed_Date_24] = (SELECT TOP 1 [Completed_Date] FROM @Courses WHERE [Course] = base.[Classroom_Course_24])
		, [Classroom_Course_Completed_Date_25] = (SELECT TOP 1 [Completed_Date] FROM @Courses WHERE [Course] = base.[Classroom_Course_25])
		, [Evolve_Module_Completed_Date_1] = (SELECT TOP 1 [Completed_Date] FROM @Modules WHERE [Module] = base.[Evolve_Module_1])
		, [Evolve_Module_Completed_Date_2] = (SELECT TOP 1 [Completed_Date] FROM @Modules WHERE [Module] = base.[Evolve_Module_2])
		, [Evolve_Module_Completed_Date_3] = (SELECT TOP 1 [Completed_Date] FROM @Modules WHERE [Module] = base.[Evolve_Module_3])
		, [Evolve_Module_Completed_Date_4] = (SELECT TOP 1 [Completed_Date] FROM @Modules WHERE [Module] = base.[Evolve_Module_4])
		, [Evolve_Module_Completed_Date_5] = (SELECT TOP 1 [Completed_Date] FROM @Modules WHERE [Module] = base.[Evolve_Module_5])
		, [Evolve_Module_Completed_Date_6] = (SELECT TOP 1 [Completed_Date] FROM @Modules WHERE [Module] = base.[Evolve_Module_6])
		, [Evolve_Module_Completed_Date_7] = (SELECT TOP 1 [Completed_Date] FROM @Modules WHERE [Module] = base.[Evolve_Module_7])
		, [Evolve_Module_Completed_Date_8] = (SELECT TOP 1 [Completed_Date] FROM @Modules WHERE [Module] = base.[Evolve_Module_8])
		, [Evolve_Module_Completed_Date_9] = (SELECT TOP 1 [Completed_Date] FROM @Modules WHERE [Module] = base.[Evolve_Module_9])
		, [Evolve_Module_Completed_Date_10] = (SELECT TOP 1 [Completed_Date] FROM @Modules WHERE [Module] = base.[Evolve_Module_10])
		, [Evolve_Module_Completed_Date_11] = (SELECT TOP 1 [Completed_Date] FROM @Modules WHERE [Module] = base.[Evolve_Module_11])
		, [Evolve_Module_Completed_Date_12] = (SELECT TOP 1 [Completed_Date] FROM @Modules WHERE [Module] = base.[Evolve_Module_12])
		, [Evolve_Module_Completed_Date_13] = (SELECT TOP 1 [Completed_Date] FROM @Modules WHERE [Module] = base.[Evolve_Module_13])
		, [Evolve_Module_Completed_Date_14] = (SELECT TOP 1 [Completed_Date] FROM @Modules WHERE [Module] = base.[Evolve_Module_14])
		, [Evolve_Module_Completed_Date_15] = (SELECT TOP 1 [Completed_Date] FROM @Modules WHERE [Module] = base.[Evolve_Module_15])
		, [Evolve_Module_Completed_Date_16] = (SELECT TOP 1 [Completed_Date] FROM @Modules WHERE [Module] = base.[Evolve_Module_16])
		, [Evolve_Module_Completed_Date_17] = (SELECT TOP 1 [Completed_Date] FROM @Modules WHERE [Module] = base.[Evolve_Module_17])
		, [Evolve_Module_Completed_Date_18] = (SELECT TOP 1 [Completed_Date] FROM @Modules WHERE [Module] = base.[Evolve_Module_18])
		, [Evolve_Module_Completed_Date_19] = (SELECT TOP 1 [Completed_Date] FROM @Modules WHERE [Module] = base.[Evolve_Module_19])
		, [Evolve_Module_Completed_Date_20] = (SELECT TOP 1 [Completed_Date] FROM @Modules WHERE [Module] = base.[Evolve_Module_20])
		, [Evolve_Module_Completed_Date_21] = (SELECT TOP 1 [Completed_Date] FROM @Modules WHERE [Module] = base.[Evolve_Module_21])
		, [Evolve_Module_Completed_Date_22] = (SELECT TOP 1 [Completed_Date] FROM @Modules WHERE [Module] = base.[Evolve_Module_22])
		, [Evolve_Module_Completed_Date_23] = (SELECT TOP 1 [Completed_Date] FROM @Modules WHERE [Module] = base.[Evolve_Module_23])
		, [Evolve_Module_Completed_Date_24] = (SELECT TOP 1 [Completed_Date] FROM @Modules WHERE [Module] = base.[Evolve_Module_24])
		, [Evolve_Module_Completed_Date_25] = (SELECT TOP 1 [Completed_Date] FROM @Modules WHERE [Module] = base.[Evolve_Module_25])
		, [Qualification_Completed_Date_1] = (SELECT TOP 1 [Completed_Date] FROM @Qualifications WHERE [Qualification] = base.[Qualification_1])
		, [Qualification_Completed_Date_2] = (SELECT TOP 1 [Completed_Date] FROM @Qualifications WHERE [Qualification] = base.[Qualification_2])
		, [Qualification_Completed_Date_3] = (SELECT TOP 1 [Completed_Date] FROM @Qualifications WHERE [Qualification] = base.[Qualification_3])
		, [Qualification_Completed_Date_4] = (SELECT TOP 1 [Completed_Date] FROM @Qualifications WHERE [Qualification] = base.[Qualification_4])
		, [Qualification_Completed_Date_5] = (SELECT TOP 1 [Completed_Date] FROM @Qualifications WHERE [Qualification] = base.[Qualification_5])
		, [Qualification_Completed_Date_6] = (SELECT TOP 1 [Completed_Date] FROM @Qualifications WHERE [Qualification] = base.[Qualification_6])
		, [Qualification_Completed_Date_7] = (SELECT TOP 1 [Completed_Date] FROM @Qualifications WHERE [Qualification] = base.[Qualification_7])
		, [Qualification_Completed_Date_8] = (SELECT TOP 1 [Completed_Date] FROM @Qualifications WHERE [Qualification] = base.[Qualification_8])
		, [Qualification_Completed_Date_9] = (SELECT TOP 1 [Completed_Date] FROM @Qualifications WHERE [Qualification] = base.[Qualification_9])
		, [Qualification_Completed_Date_10] = (SELECT TOP 1 [Completed_Date] FROM @Qualifications WHERE [Qualification] = base.[Qualification_10])
		, [Qualification_Completed_Date_11] = (SELECT TOP 1 [Completed_Date] FROM @Qualifications WHERE [Qualification] = base.[Qualification_11])
		, [Qualification_Completed_Date_12] = (SELECT TOP 1 [Completed_Date] FROM @Qualifications WHERE [Qualification] = base.[Qualification_12])
		, [Qualification_Completed_Date_13] = (SELECT TOP 1 [Completed_Date] FROM @Qualifications WHERE [Qualification] = base.[Qualification_13])
		, [Qualification_Completed_Date_14] = (SELECT TOP 1 [Completed_Date] FROM @Qualifications WHERE [Qualification] = base.[Qualification_14])
		, [Qualification_Completed_Date_15] = (SELECT TOP 1 [Completed_Date] FROM @Qualifications WHERE [Qualification] = base.[Qualification_15])
	FROM [inserted] base WHERE base.[id] = [dbo].[tbuser_Job_Role].[id]


 END
 
 GO
