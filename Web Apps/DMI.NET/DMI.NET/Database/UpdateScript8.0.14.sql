IF EXISTS (SELECT *	FROM dbo.sysobjects	WHERE id = object_id(N'[dbo].[sp_ASR_Bradford_DeleteAbsences]') AND xtype in (N'P'))
	DROP PROCEDURE [dbo].[sp_ASR_Bradford_DeleteAbsences];
GO



GO

CREATE PROCEDURE [dbo].[sp_ASR_Bradford_DeleteAbsences]
(
	@pdReportStart	  	datetime,
	@pdReportEnd				datetime,
	@pbOmitBeforeStart	bit,
	@pbOmitAfterEnd			bit,
	@pcReportTableName	char(30)
)
AS
BEGIN

	SET NOCOUNT ON;

	declare @piID as integer;
	declare @pdStartDate as datetime;
	declare @pdEndDate as datetime;
	declare @iDuration as float;
	declare @pbDeleteThisAbsence as bit;
	declare @sSQL as varchar(MAX);

	set @sSQL = 'DECLARE BradfordIndexCursor2 CURSOR FOR SELECT Absence_ID, Start_Date, End_Date, Duration FROM ' + @pcReportTableName;
	execute(@sSQL);
	open BradfordIndexCursor2;

	Fetch Next From BradfordIndexCursor2 Into @piID, @pdStartDate, @pdEndDate, @iDuration;
	while @@FETCH_STATUS = 0
		begin
			set @pbDeleteThisAbsence = 0;
			if @pdEndDate < @pdReportStart set @pbDeleteThisAbsence = 1;
			if @pdStartDate > @pdReportEnd set @pbDeleteThisAbsence = 1;
			if @iDuration = 0 set @pbDeleteThisAbsence = 1;

			if @pbOmitBeforeStart = 1 and (@pdStartDate < @pdReportStart)  set @pbDeleteThisAbsence = 1;
			if @pbOmitAfterEnd = 1 and (@pdEndDate > @pdReportEnd)  set @pbDeleteThisAbsence = 1;

			if @pbDeleteThisAbsence = 1
				begin
					set @sSQL = 'DELETE FROM ' + @pcReportTableName + ' Where Absence_ID = Convert(Int,' + Convert(char(10),@piId) + ')';
					execute(@sSQL);
				end

			Fetch Next From BradfordIndexCursor2 Into @piID, @pdStartDate, @pdEndDate, @iDuration;
		end

	close BradfordIndexCursor2;
	deallocate BradfordIndexCursor2;

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
