/* Required Tables

250	- Absence_Entry
251	- Appointment_Absence_Entry
252	- Absence_Breakdown


*/

IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[trsys_Absence_Breakdown_D01]') AND xtype in (N'TR'))
	DROP TRIGGER [dbo].[trsys_Absence_Breakdown_D01]
GO


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
DISABLE TRIGGER trsys_Absence_i01 ON [dbo].[tbuser_Absence]
GO

DISABLE TRIGGER trsys_Absence_i02 ON [dbo].[tbuser_Absence]
GO

DISABLE TRIGGER trsys_Absence_u01 ON [dbo].[tbuser_Absence]
GO

DISABLE TRIGGER trsys_Absence_u02 ON [dbo].[tbuser_Absence]
GO

DISABLE TRIGGER trsys_Absence_d01 ON [dbo].[tbuser_Absence]
GO


DISABLE TRIGGER trsys_Absence_Entry_i01 ON [dbo].[tbuser_Absence_Entry]
GO

DISABLE TRIGGER trsys_Absence_Entry_i02 ON [dbo].[tbuser_Absence_Entry]
GO

DISABLE TRIGGER trsys_Absence_Entry_u01 ON [dbo].[tbuser_Absence_Entry]
GO

DISABLE TRIGGER trsys_Absence_Entry_u02 ON [dbo].[tbuser_Absence_Entry]
GO

DISABLE TRIGGER trsys_Absence_Entry_d01 ON [dbo].[tbuser_Absence_Entry]
GO

DISABLE TRIGGER trsys_Appointment_Absence_Entry_i01 ON [dbo].[tbuser_Appointment_Absence_Entry]
GO

DISABLE TRIGGER trsys_Appointment_Absence_Entry_i02 ON [dbo].[tbuser_Appointment_Absence_Entry]
GO

DISABLE TRIGGER trsys_Appointment_Absence_Entry_u01 ON [dbo].[tbuser_Appointment_Absence_Entry]
GO

DISABLE TRIGGER trsys_Appointment_Absence_Entry_u02 ON [dbo].[tbuser_Appointment_Absence_Entry]
GO

DISABLE TRIGGER trsys_Appointment_Absence_Entry_d01 ON [dbo].[tbuser_Appointment_Absence_Entry]
GO


DISABLE TRIGGER trsys_Absence_Breakdown_i01 ON [dbo].[tbuser_Absence_Breakdown]
GO

DISABLE TRIGGER trsys_Absence_Breakdown_i02 ON [dbo].[tbuser_Absence_Breakdown]
GO

DISABLE TRIGGER trsys_Absence_Breakdown_u01 ON [dbo].[tbuser_Absence_Breakdown]
GO

DISABLE TRIGGER trsys_Absence_Breakdown_u02 ON [dbo].[tbuser_Absence_Breakdown]
GO


IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[trsys_Absence_Breakdown]') AND xtype in (N'TR'))
	DROP TRIGGER [dbo].[trsys_Absence_Breakdown];
GO


CREATE TRIGGER [dbo].[trcustom_Absence_Entry_P&E] ON [dbo].[tbuser_Absence_Entry]
    AFTER INSERT, UPDATE, DELETE
AS
BEGIN
	--SYSTEM MANAGER AUTOMATICALLY UPGRADED TO 4.3

    SET NOCOUNT ON;

    DELETE [dbo].[tbuser_Absence_Breakdown] WHERE [id_250] IN (SELECT DISTINCT [id] FROM deleted);

	INSERT Absence_Breakdown([source], ID_250, Post_ID, [Type], Payroll_Type_Code, Reason, Payroll_Reason_Code, Absence_In, Duration, Absence_Date, [Session]
		, Day_Pattern_AM, Day_Pattern_PM, Hour_Pattern_AM, Hour_Pattern_PM, Staff_Number, Payroll_Company_Code)	
		SELECT 'pers', i.ID, ap.ID, i.Absence_Type, ISNULL(at.Payroll_Code, ''), i.Reason, ISNULL(ar.Payroll_Code, ''), wp.Absence_In
			, dbo.udfsysDurationFromPattern(wp.Absence_In, dr.IndividualDate, dr.SessionType, wp.Sunday_Hours_AM, wp.Monday_Hours_AM, wp.Tuesday_Hours_AM, wp.Wednesday_Hours_AM, wp.Thursday_Hours_AM, wp.Friday_Hours_AM, wp.Saturday_Hours_AM, wp.Sunday_Hours_PM, wp.Monday_Hours_PM, wp.Tuesday_Hours_PM, wp.Wednesday_Hours_PM, wp.Thursday_Hours_PM, wp.Friday_Hours_PM, wp.Saturday_Hours_PM)
			, dr.IndividualDate, dr.SessionType
			, wp.Day_Pattern_AM, wp.Day_Pattern_PM, wp.Hour_Pattern_AM, wp.Hour_Pattern_PM
			, pr.Staff_Number, pr.Payroll_Company_Code
		FROM inserted i
			CROSS APPLY [dbo].[udfsysDateRangeToTable] ('d', i.Start_Date, i.Start_Session,  i.End_Date, i.End_Session) dr
			INNER JOIN Appointments ap ON ap.ID_1 = i.ID_1
			INNER JOIN Appointment_Working_Patterns wp ON wp.ID_3 = ap.ID
			INNER JOIN Personnel_Records pr ON pr.ID = i.ID_1
			LEFT JOIN Absence_Type_Table at ON at.Absence_Type = i.Absence_Type
			LEFT JOIN Absence_Reason_Table ar ON ar.Reason = i.Reason
		WHERE wp.Effective_Date <= dr.IndividualDate AND (wp.End_Date >= dr.IndividualDate OR wp.End_Date IS NULL);

END
GO

CREATE TRIGGER [dbo].[trcustom_Appointment_Absence_Entry_P&E] ON [dbo].[tbuser_Appointment_Absence_Entry]
    AFTER INSERT, UPDATE, DELETE
AS
BEGIN
	--SYSTEM MANAGER AUTOMATICALLY UPGRADED TO 4.3

    SET NOCOUNT ON;

    DELETE [dbo].[tbuser_Absence_Breakdown] WHERE [id_251] IN (SELECT DISTINCT [id] FROM deleted);

	INSERT Absence_Breakdown([source], ID_251, Post_ID, [Type], Payroll_Type_Code, Reason, Payroll_Reason_Code, Absence_In, Duration, Absence_Date, [Session]
		, Day_Pattern_AM, Day_Pattern_PM, Hour_Pattern_AM, Hour_Pattern_PM, Staff_Number, Payroll_Company_Code)	
		SELECT 'post', i.ID, wp.ID_3, i.Absence_Type, ISNULL(at.Payroll_Code, ''), i.Reason, ISNULL(ar.Payroll_Code, ''), wp.Absence_In
			, dbo.udfsysDurationFromPattern(wp.Absence_In, dr.IndividualDate, dr.SessionType, wp.Sunday_Hours_AM, wp.Monday_Hours_AM, wp.Tuesday_Hours_AM, wp.Wednesday_Hours_AM, wp.Thursday_Hours_AM, wp.Friday_Hours_AM, wp.Saturday_Hours_AM, wp.Sunday_Hours_PM, wp.Monday_Hours_PM, wp.Tuesday_Hours_PM, wp.Wednesday_Hours_PM, wp.Thursday_Hours_PM, wp.Friday_Hours_PM, wp.Saturday_Hours_PM)
			, dr.IndividualDate, dr.SessionType
			, wp.Day_Pattern_AM, wp.Day_Pattern_PM, wp.Hour_Pattern_AM, wp.Hour_Pattern_PM
			, pr.Staff_Number, pr.Payroll_Company_Code
		FROM inserted i
			CROSS APPLY [dbo].[udfsysDateRangeToTable] ('d', i.Start_Date, i.Start_Session,  i.End_Date, i.End_Session) dr
			INNER JOIN Appointments ap ON ap.ID = i.ID_3
			INNER JOIN Appointment_Working_Patterns wp ON wp.ID_3 = i.ID_3
			INNER JOIN Personnel_Records pr ON pr.ID = ap.ID_1
			LEFT JOIN Absence_Type_Table at ON at.Absence_Type = i.Absence_Type
			LEFT JOIN Absence_Reason_Table ar ON ar.Reason = i.Reason
		WHERE wp.Effective_Date <= dr.IndividualDate AND (wp.End_Date >= dr.IndividualDate OR wp.End_Date IS NULL);

END
GO


CREATE TRIGGER [dbo].[trcustom_Absence_Breakdown_P&E] ON [dbo].[tbuser_Absence_Breakdown]
    AFTER INSERT
AS
BEGIN
	--SYSTEM MANAGER AUTOMATICALLY UPGRADED TO 4.3

    SET NOCOUNT ON;

	DECLARE @AbsenceID	integer,
			@startDate	datetime,
			@endDate	datetime;

	INSERT Absence(Absence_Type, Payroll_Code, Reason, Payroll_Reason, Start_Date, Start_Session, End_Date, End_Session, Duration_Days, Duration_Hours, ID_1, Absence_In)
		SELECT DISTINCT ab.Type, ab.Payroll_Type_Code, ab.Reason, ab.Payroll_Reason_Code
			, m.startdate
			, (SELECT DISTINCT CASE WHEN [Session] = 'Day' THEN 'AM' ELSE [Session] END FROM inserted WHERE (ID_250 = ab.ID_250 OR ID_251 = ab.ID_251) AND Absence_Date = m.startdate)
			, m.enddate
			, (SELECT DISTINCT CASE WHEN [Session] = 'Day' THEN 'PM' ELSE [Session] END FROM inserted WHERE (ID_250 = ab.ID_250 OR ID_251 = ab.ID_251) AND Absence_Date = m.enddate)
			, CASE WHEN  ab.Absence_In = 'Days' THEN m.Duration ELSE 0 END
			, CASE WHEN  ab.Absence_In = 'Hours' THEN m.Duration ELSE 0 END
			, CASE WHEN  ab.[source] = 'pers' THEN ae.ID_1 ELSE a.ID_1 END
			, ab.Absence_In
		FROM inserted ab
			LEFT JOIN Absence_Entry ae ON ae.ID = ab.ID_250
			LEFT JOIN Appointment_Absence_Entry aae ON aae.ID = ab.ID_251
			LEFT JOIN Appointments a ON a.ID = aae.ID_3
		CROSS APPLY (
			SELECT MIN(range.Absence_Date) AS startdate, MAX(range.Absence_Date) AS enddate, SUM(Duration) AS Duration
			FROM inserted range
			WHERE ab.ID_250 = range.ID_250 OR ab.ID_251 = range.ID_251) m;


	SELECT @AbsenceID = MAX(ID) FROM Absence

	UPDATE [tbuser_Absence_Breakdown]
		SET id_2 = @AbsenceID
	FROM [inserted] base WHERE base.[id] = [dbo].[tbuser_Absence_Breakdown].[id]


    DECLARE @recordID int,
	    @TStamp int,
	    @hResult int,
        @iAccordBatchID integer,
        @iAccordDefaultStatus integer = 1,
        @iAccordManualSendType smallint = -1,
        @bAccordResend bit,
        @bAccordBypassFilter bit,
        @recordDesc varchar(255),
		@cursInsertedRecords cursor,
        @fValidRecord bit = 1;

    DECLARE @sTempInsCol varchar(MAX), @insCol_3940 varchar(MAX), @insCol_3942 varchar(MAX), @insCol_3941 varchar(MAX), @insCol_3946 varchar(MAX), @insCol_3943 datetime, @insCol_3944 varchar(MAX), @insCol_3948 varchar(MAX), @insCol_3949 varchar(MAX), @insCol_3954 numeric(4,2), @insCol_3950 varchar(MAX), @insCol_3951 varchar(MAX), @insCol_3952 varchar(MAX), @insCol_3953 varchar(MAX), @insParentID_2 integer
    DECLARE @sTempDelCol varchar(MAX), @delCol_3940 varchar(MAX), @delCol_3942 varchar(MAX), @delCol_3941 varchar(MAX), @delCol_3946 varchar(MAX), @delCol_3943 datetime, @delCol_3944 varchar(MAX), @delCol_3948 varchar(MAX), @delCol_3949 varchar(MAX), @delCol_3954 numeric(4,2), @delCol_3950 varchar(MAX), @delCol_3951 varchar(MAX), @delCol_3952 varchar(MAX), @delCol_3953 varchar(MAX), @delParentID_2 integer

    /* Loop through the virtual 'inserted' table, getting the record ID of each updated record. */
    SET @cursInsertedRecords = CURSOR LOCAL FAST_FORWARD READ_ONLY FOR SELECT inserted.id, convert(int,inserted.timestamp), inserted.[_description], inserted.Post_ID, inserted.Payroll_Company_Code, inserted.Staff_Number, inserted.Payroll_Type_Code, inserted.Absence_Date, inserted.Session, inserted.Payroll_Reason_Code, inserted.Absence_In, inserted.Duration, inserted.Day_Pattern_AM, inserted.Day_Pattern_PM, inserted.Hour_Pattern_AM, inserted.Hour_Pattern_PM, isnull(inserted.ID_2,0), deleted.Post_ID, deleted.Payroll_Company_Code, deleted.Staff_Number, deleted.Payroll_Type_Code, deleted.Absence_Date, deleted.Session, deleted.Payroll_Reason_Code, deleted.Absence_In, deleted.Duration, deleted.Day_Pattern_AM, deleted.Day_Pattern_PM, deleted.Hour_Pattern_AM, deleted.Hour_Pattern_PM, isnull(deleted.ID_2,0) FROM inserted
    LEFT OUTER JOIN deleted ON inserted.id = deleted.id

    OPEN @cursInsertedRecords
    FETCH NEXT FROM @cursInsertedRecords INTO @recordID, @TStamp, @recorddesc, @insCol_3940, @insCol_3942, @insCol_3941, @insCol_3946, @insCol_3943, @insCol_3944, @insCol_3948, @insCol_3949, @insCol_3954, @insCol_3950, @insCol_3951, @insCol_3952, @insCol_3953, @insParentID_2, @delCol_3940, @delCol_3942, @delCol_3941, @delCol_3946, @delCol_3943, @delCol_3944, @delCol_3948, @delCol_3949, @delCol_3954, @delCol_3950, @delCol_3951, @delCol_3952, @delCol_3953, @delParentID_2
    WHILE (@@fetch_status = 0) AND (@fValidRecord = 1)
    BEGIN

        /* ----------------------- */
        /* Payroll Triggers. Cribbed from the system generated one. I do apologise */
        /* ----------------------- */
        IF @fValidRecord = 1
        BEGIN
          DECLARE @iAccordTransactionID as int
          DECLARE @bFilter as bit
          DECLARE @bAccordSendAllFields as bit
          DECLARE @intDefaultAccordStatus as int = 1
          DECLARE @intDefaultAccordType as int

		  PRINT @iAccordManualSendType
		  PRINT @bAccordBypassFilter


          EXEC @hResult = dbo.sp_ASRExpr_56056 @bFilter OUTPUT, @recordID

		-- P&E needs to ignore filter
		PRINT @bFilter
		SET @bFilter = 1


          IF (@iAccordManualSendType = 81 AND @bAccordBypassFilter = 1)
            OR (@iAccordManualSendType = 81 AND @bAccordBypassFilter = 0 AND @bFilter = 1)
            OR (@bFilter = 1 AND @iAccordManualSendType = -1)

          BEGIN



          EXEC dbo.spASRAccordNeedToSendAll  81, @recordID, @bAccordResend OUTPUT

            IF (ISNULL(@insCol_3940,'') <> ISNULL(@delCol_3940,'') OR ISNULL(@insCol_3942,'') <> ISNULL(@delCol_3942,'') OR ISNULL(@insCol_3941,'') <> ISNULL(@delCol_3941,'') OR ISNULL(@insCol_3946,'') <> ISNULL(@delCol_3946,'') OR ISNULL(@insCol_3943,'') <> ISNULL(@delCol_3943,'') OR ISNULL(@insCol_3944,'') <> ISNULL(@delCol_3944,'') OR ISNULL(@insCol_3948,'') <> ISNULL(@delCol_3948,'') OR ISNULL(@insCol_3949,'') <> ISNULL(@delCol_3949,'') OR ISNULL(@insCol_3954,0) <> ISNULL(@delCol_3954,0) OR ISNULL(@insCol_3950,'') <> ISNULL(@delCol_3950,'') OR ISNULL(@insCol_3951,'') <> ISNULL(@delCol_3951,'') OR ISNULL(@insCol_3952,'') <> ISNULL(@delCol_3952,'') OR ISNULL(@insCol_3953,'') <> ISNULL(@delCol_3953,'') OR  @bAccordResend = 1)
            BEGIN

              EXEC dbo.spASRAccordPopulateTransaction @iAccordTransactionID OUTPUT, 81, 1 , @iAccordDefaultStatus, @recordID, 1, @bAccordSendAllFields OUTPUT
              SET @sTempInsCol = ISNULL(CONVERT(varchar(255), @insCol_3940), '')
              SET @sTempDelCol = ISNULL(CONVERT(varchar(255), @delCol_3940), '')
              EXEC dbo.spASRAccordPopulateTransactionData @iAccordTransactionID,0, @sTempDelCol,@sTempInsCol
              SET @sTempInsCol = ISNULL(CONVERT(varchar(255), @insCol_3942), '')
              SET @sTempDelCol = ISNULL(CONVERT(varchar(255), @delCol_3942), '')
              EXEC dbo.spASRAccordPopulateTransactionData @iAccordTransactionID,1, @sTempDelCol,@sTempInsCol
              UPDATE ASRSysAccordTransactions SET [CompanyCode] = @sTempInsCol WHERE [TransactionID] = @iAccordTransactionID
              SET @sTempInsCol = ISNULL(CONVERT(varchar(255), @insCol_3941), '')
              SET @sTempDelCol = ISNULL(CONVERT(varchar(255), @delCol_3941), '')
              EXEC dbo.spASRAccordPopulateTransactionData @iAccordTransactionID,2, @sTempDelCol,@sTempInsCol
              UPDATE ASRSysAccordTransactions SET [EmployeeCode] = @sTempInsCol WHERE [TransactionID] = @iAccordTransactionID
              SET @sTempInsCol = ISNULL(CONVERT(varchar(255), @insCol_3946), '')
              SET @sTempDelCol = ISNULL(CONVERT(varchar(255), @delCol_3946), '')
              EXEC dbo.spASRAccordPopulateTransactionData @iAccordTransactionID,3, @sTempDelCol,@sTempInsCol
              SET @sTempInsCol = ISNULL(CONVERT(varchar(255),DATEPART(year, @insCol_3943)) + RIGHT('0' + CONVERT(varchar(2),DATEPART(month, @insCol_3943)),2) + RIGHT('0' + CONVERT(varchar(2),DATEPART(day, @insCol_3943)),2),'00000000')
              SET @sTempDelCol = ISNULL(CONVERT(varchar(255),DATEPART(year, @delCol_3943)) + RIGHT('0' + CONVERT(varchar(2),DATEPART(month, @delCol_3943)),2) + RIGHT('0' + CONVERT(varchar(2),DATEPART(day, @delCol_3943)),2),'00000000')
              EXEC dbo.spASRAccordPopulateTransactionData @iAccordTransactionID,4, @sTempDelCol,@sTempInsCol
              SET @sTempInsCol = ISNULL(CONVERT(varchar(255), @insCol_3944), '')
              SET @sTempDelCol = ISNULL(CONVERT(varchar(255), @delCol_3944), '')
              EXEC dbo.spASRAccordPopulateTransactionData @iAccordTransactionID,5, @sTempDelCol,@sTempInsCol
              SET @sTempInsCol = ISNULL(CONVERT(varchar(255), @insCol_3948), '')
              SET @sTempDelCol = ISNULL(CONVERT(varchar(255), @delCol_3948), '')
              EXEC dbo.spASRAccordPopulateTransactionData @iAccordTransactionID,6, @sTempDelCol,@sTempInsCol
              SET @sTempInsCol = ISNULL(CONVERT(varchar(255), @insCol_3949), '')
              SET @sTempDelCol = ISNULL(CONVERT(varchar(255), @delCol_3949), '')
              EXEC dbo.spASRAccordPopulateTransactionData @iAccordTransactionID,7, @sTempDelCol,@sTempInsCol
              SET @sTempInsCol = ISNULL(CONVERT(varchar(255), @insCol_3954), '')
              SET @sTempDelCol = ISNULL(CONVERT(varchar(255), @delCol_3954), '')
              EXEC dbo.spASRAccordPopulateTransactionData @iAccordTransactionID,8, @sTempDelCol,@sTempInsCol
              SET @sTempInsCol = ISNULL(CONVERT(varchar(255), @insCol_3950), '')
              SET @sTempDelCol = ISNULL(CONVERT(varchar(255), @delCol_3950), '')
              EXEC dbo.spASRAccordPopulateTransactionData @iAccordTransactionID,9, @sTempDelCol,@sTempInsCol
              SET @sTempInsCol = ISNULL(CONVERT(varchar(255), @insCol_3951), '')
              SET @sTempDelCol = ISNULL(CONVERT(varchar(255), @delCol_3951), '')
              EXEC dbo.spASRAccordPopulateTransactionData @iAccordTransactionID,10, @sTempDelCol,@sTempInsCol
              SET @sTempInsCol = ISNULL(CONVERT(varchar(255), @insCol_3952), '')
              SET @sTempDelCol = ISNULL(CONVERT(varchar(255), @delCol_3952), '')
              EXEC dbo.spASRAccordPopulateTransactionData @iAccordTransactionID,11, @sTempDelCol,@sTempInsCol
              SET @sTempInsCol = ISNULL(CONVERT(varchar(255), @insCol_3953), '')
              SET @sTempDelCol = ISNULL(CONVERT(varchar(255), @delCol_3953), '')
              EXEC dbo.spASRAccordPopulateTransactionData @iAccordTransactionID,12, @sTempDelCol,@sTempInsCol
            END

          END
          EXEC dbo.spASRAccordPurgeTemp 1, @recordID

        END


        IF @fValidRecord = 1 FETCH NEXT FROM @cursInsertedRecords INTO @recordID, @TStamp, @recorddesc, @insCol_3940, @insCol_3942, @insCol_3941, @insCol_3946, @insCol_3943, @insCol_3944, @insCol_3948, @insCol_3949, @insCol_3954, @insCol_3950, @insCol_3951, @insCol_3952, @insCol_3953, @insParentID_2, @delCol_3940, @delCol_3942, @delCol_3941, @delCol_3946, @delCol_3943, @delCol_3944, @delCol_3948, @delCol_3949, @delCol_3954, @delCol_3950, @delCol_3951, @delCol_3952, @delCol_3953, @delParentID_2
    END
    IF @fValidRecord = 1 CLOSE @cursInsertedRecords;
    DEALLOCATE @cursInsertedRecords;



END

GO

CREATE TRIGGER [dbo].[trcustom_Absence_Breakdown_P&E_D02] ON [dbo].[tbuser_Absence_Breakdown]
    INSTEAD OF DELETE
AS
BEGIN
	--SYSTEM MANAGER AUTOMATICALLY UPGRADED TO 4.3

    SET NOCOUNT ON;

	DECLARE @AbsenceID	integer,
			@startDate	datetime,
			@endDate	datetime;

    DELETE [dbo].[tbuser_Absence] WHERE [id] IN (SELECT DISTINCT [id_2] FROM deleted);

	WITH base AS (SELECT * FROM dbo.[tbuser_Absence_Breakdown]
        WHERE [id] IN (SELECT DISTINCT [id] FROM deleted))
        DELETE FROM base;

END








GO


--EXEC sp_settriggerorder @triggername=N'[dbo].[trcustom_Absence_Entry_P&E]', @order=N'Last', @stmttype=N'INSERT'
--EXEC sp_settriggerorder @triggername=N'[dbo].[trcustom_Absence_Entry_P&E]', @order=N'Last', @stmttype=N'UPDATE'
--EXEC sp_settriggerorder @triggername=N'[dbo].[trcustom_Absence_Entry_P&E]', @order=N'Last', @stmttype=N'DELETE'

GO


