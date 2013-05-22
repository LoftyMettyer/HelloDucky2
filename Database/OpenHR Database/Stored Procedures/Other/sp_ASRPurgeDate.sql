CREATE PROCEDURE [dbo].[sp_ASRPurgeDate]
(
    @purgedate varchar(MAX) OUTPUT,
    @purgekey varchar(MAX)
)
AS
BEGIN
    DECLARE @unit char(1),
            @period int,
            @lastPurge datetime,
            @today datetime;

    /* Only get date and not current time */
    select @today = convert(datetime,convert(varchar,getdate(),101));

    /* Get purge period details */
    SELECT @unit = unit
         , @period = (period * -1)
         , @lastPurge = lastpurgedate
    FROM   asrsyspurgeperiods
    WHERE  purgekey = @purgekey;

    /* calculate purge date */
    SELECT @purgedate = CASE @unit
        WHEN 'D' THEN dateadd(dd,@period,@today)
        WHEN 'W' THEN dateadd(ww,@period,@today)
        WHEN 'M' THEN dateadd(mm,@period,@today)
        WHEN 'Y' THEN dateadd(yy,@period,@today)
    END;

    IF @purgedate IS NULL OR datediff(d,@purgedate,@lastPurge) > 0
    BEGIN
      SET @purgedate = @lastPurge;
    END

END