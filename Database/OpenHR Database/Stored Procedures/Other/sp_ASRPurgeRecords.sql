CREATE PROCEDURE [dbo].[sp_ASRPurgeRecords]
(
    @PurgeKey varchar(255),
    @TableName varchar(255),
    @DateColumn varchar(255)
)
AS
BEGIN

    /* EXEC sp_ASRPurgeRecords 'EMAIL', 'ASRSysEmailQueue', 'DateDue' */

    DECLARE @PurgeDate datetime;
    DECLARE @sSQL nvarchar(MAX);

    EXEC [dbo].[sp_ASRPurgeDate] @PurgeDate OUTPUT, @PurgeKey;

    SELECT @sSQL = 'DELETE FROM ' + @TableName + ' WHERE ' + @DateColumn + ' < ''' + convert(varchar,@PurgeDate,101) + '''';
    EXEC sp_executesql @sSQL;

END