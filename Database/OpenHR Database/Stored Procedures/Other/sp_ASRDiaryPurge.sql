CREATE PROCEDURE [dbo].[sp_ASRDiaryPurge]
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @PurgeDate	datetime,
			@sSQL		nvarchar(MAX),
			@unit		char(1),
            @period		int,
            @today		datetime;

    /* Note can't use sp_ASRPurgeDate as the diary dates include the time !!! */

    select @today = getdate();

    /* Get purge period details */
    select @unit = unit, @period = (period * -1)
		from asrsyspurgeperiods where purgekey =  'DIARYSYS';

    /* calculate purge date */
    SELECT @purgedate = CASE @unit
        WHEN 'D' THEN dateadd(dd,@period,@today)
        WHEN 'W' THEN dateadd(ww,@period,@today)
        WHEN 'M' THEN dateadd(mm,@period,@today)
        WHEN 'Y' THEN dateadd(yy,@period,@today)
    END;

    SELECT @sSQL = 'DELETE FROM ASRSysDiaryEvents WHERE EventDate < ''' + 
		convert(varchar,@PurgeDate,101) + ''' AND ColumnID > 0';

    EXEC sp_executesql @sSQL;


    /* Get purge period details */
    select @unit = unit, @period = (period * -1)
		from asrsyspurgeperiods where purgekey =  'DIARYMAN';

    /* calculate purge date */
    SELECT @purgedate = CASE @unit
        WHEN 'D' THEN dateadd(dd,@period,@today)
        WHEN 'W' THEN dateadd(ww,@period,@today)
        WHEN 'M' THEN dateadd(mm,@period,@today)
        WHEN 'Y' THEN dateadd(yy,@period,@today)
    END;

    SELECT @sSQL = 'DELETE FROM ASRSysDiaryEvents WHERE EventDate < ''' 
		+ convert(varchar,@PurgeDate,101) + ' ' + convert(varchar,@PurgeDate,108)
		+ ''' AND ColumnID = 0';
    EXEC sp_executesql @sSQL;

END