/* Required Tables

250	- Absence_Entry
251	- Appointment_Absence_Entry
252	- Absence_Breakdown


*/


IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[trcustom_Absence_Entry_P&E]') AND xtype in (N'TR'))
	DROP TRIGGER [dbo].[trcustom_Absence_Entry_P&E]
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[trcustom_Appointment_Absence_Entry_P&E]') AND xtype in (N'TR'))
	DROP TRIGGER [dbo].[trcustom_Appointment_Absence_Entry_P&E]
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[trcustom_Absence_Breakdown_P&E]') AND xtype in (N'TR'))
	DROP TRIGGER [dbo].[trcustom_Absence_Breakdown_P&E]
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[trcustom_Absence_Breakdown_P&E_D02]') AND xtype in (N'TR'))
	DROP TRIGGER [dbo].[trcustom_Absence_Breakdown_P&E_D02]
GO


-- Some system triggers that need disabling/removing
DISABLE TRIGGER trsys_Absence_Entry_d01 ON [dbo].[tbuser_Absence_Entry]
GO

DISABLE TRIGGER trsys_Appointment_Absence_Entry_d01 ON [dbo].[tbuser_Appointment_Absence_Entry]
GO

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[trsys_Absence_Breakdown]') AND xtype in (N'TR'))
	DROP TRIGGER [dbo].[trsys_Absence_Breakdown];
GO



CREATE TRIGGER [dbo].[trcustom_Absence_Entry_P&E] ON [dbo].[tbuser_Absence_Entry]
    AFTER INSERT, UPDATE, DELETE
AS
BEGIN
    SET NOCOUNT ON;

PRINT '[trcustom_Absence_Entry_P&E]'

	DECLARE	@employeeID		integer,
			@absenceID		integer,
			@start_date		datetime,
			@end_date		datetime,
			@absenceIn		varchar(5);

    DELETE [dbo].[tbuser_Absence_Breakdown] WHERE [id_250] IN (SELECT DISTINCT [id] FROM deleted);

	DECLARE AbsenceCursor CURSOR FOR SELECT [ID_1], [ID], [start_date], [end_date] FROM inserted
    OPEN AbsenceCursor
    FETCH NEXT FROM AbsenceCursor INTO @employeeID, @absenceID, @start_date, @end_date
    WHILE @@FETCH_STATUS = 0
    BEGIN

--		DELETE FROM Absence_Breakdown WHERE ID_250 = @AbsenceID;
		INSERT Absence_Breakdown([source], ID_250, Absence_Date)
			SELECT 'pers', @absenceID, dr.IndividualDate FROM [dbo].[udfsysDateRangeToTable]('d', @start_date, @end_date) dr

		FETCH NEXT FROM AbsenceCursor INTO @employeeID, @absenceID, @start_date, @end_date
    END
    CLOSE AbsenceCursor
    DEALLOCATE AbsenceCursor

END
GO

CREATE TRIGGER [dbo].[trcustom_Appointment_Absence_Entry_P&E] ON [dbo].[tbuser_Appointment_Absence_Entry]
    AFTER INSERT, UPDATE, DELETE
AS
BEGIN
    SET NOCOUNT ON;

	DECLARE	@employeeID		integer,
			@absenceID		integer,
			@start_date		datetime,
			@end_date		datetime,
			@absenceIn		varchar(5);

    DELETE [dbo].[tbuser_Absence_Breakdown] WHERE [id_251] IN (SELECT DISTINCT [id] FROM deleted);

	DECLARE AbsenceCursor CURSOR FOR SELECT [ID_3], [ID], [start_date], [end_date] FROM inserted
    OPEN AbsenceCursor
    FETCH NEXT FROM AbsenceCursor INTO @employeeID, @absenceID, @start_date, @end_date
    WHILE @@FETCH_STATUS = 0
    BEGIN

		INSERT Absence_Breakdown([source], ID_251, Absence_Date)
			SELECT 'post', @absenceID, dr.IndividualDate FROM [dbo].[udfsysDateRangeToTable]('d', @start_date, @end_date) dr

		FETCH NEXT FROM AbsenceCursor INTO @employeeID, @absenceID, @start_date, @end_date
    END
    CLOSE AbsenceCursor
    DEALLOCATE AbsenceCursor

END
GO


CREATE TRIGGER [dbo].[trcustom_Absence_Breakdown_P&E] ON [dbo].[tbuser_Absence_Breakdown]
    AFTER INSERT
AS
BEGIN
    SET NOCOUNT ON;

	PRINT '[trcustom_Absence_Breakdown_P&E_INSERT]'

	DECLARE @AbsenceID	integer,
			@startDate	datetime,
			@endDate	datetime;

	SELECT @startDate = MIN(absence_date), @endDate = MAX(absence_date) FROM inserted;

	INSERT Absence(Start_Date, End_Date)
		VALUES (@startDate, @endDate);

	SELECT @AbsenceID = MAX(ID) FROM Absence

	UPDATE [tbuser_Absence_Breakdown]
		SET id_2 = @AbsenceID
	FROM [inserted] base WHERE base.[id] = [dbo].[tbuser_Absence_Breakdown].[id]

END

GO

CREATE TRIGGER [dbo].[trcustom_Absence_Breakdown_P&E_D02] ON [dbo].[tbuser_Absence_Breakdown]
    INSTEAD OF DELETE
AS
BEGIN
    SET NOCOUNT ON;

	PRINT '[trcustom_Absence_Breakdown_P&E_D02]'

	DECLARE @AbsenceID	integer,
			@startDate	datetime,
			@endDate	datetime;

	--SELECT @startDate = MIN(absence_date), @endDate = MAX(absence_date) FROM inserted;

	--SELECT DISTINCT [id_2] FROM deleted;
	--select * from deleted;

    DELETE [dbo].[tbuser_Absence] WHERE [id] IN (SELECT DISTINCT [id_2] FROM deleted);

	WITH base AS (SELECT * FROM dbo.[tbuser_Absence_Breakdown]
        WHERE [id] IN (SELECT DISTINCT [id] FROM deleted))
        DELETE FROM base;



	--INSERT Absence(Start_Date, End_Date)
	--	VALUES (@startDate, @endDate);

	--SELECT @AbsenceID = MAX(ID) FROM Absence

	--UPDATE [tbuser_Absence_Breakdown] SET id_2 = @AbsenceID;

END








GO



EXEC sp_settriggerorder @triggername=N'[dbo].[trcustom_Absence_Entry_P&E]', @order=N'Last', @stmttype=N'INSERT'
--EXEC sp_settriggerorder @triggername=N'[dbo].[trcustom_Absence_Entry_P&E]', @order=N'Last', @stmttype=N'UPDATE'
--EXEC sp_settriggerorder @triggername=N'[dbo].[trcustom_Absence_Entry_P&E]', @order=N'Last', @stmttype=N'DELETE'

GO


