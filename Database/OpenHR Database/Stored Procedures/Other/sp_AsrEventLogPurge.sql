CREATE PROCEDURE [sp_AsrEventLogPurge] AS

/* First retrieve the frequency/period info from the ASRSysEventLogPurge table */
DECLARE @intFrequency int,
        @strPeriod char(2)

/* Get the start date of the given course. */
SELECT @intFrequency = Frequency
FROM ASRSysEventLogPurge

SELECT @strPeriod = Period
FROM ASRSysEventLogPurge

IF (@intFrequency IS NOT NULL) AND (@strPeriod IS NOT NULL)

BEGIN

  /* Delete rows from the EventLog Header table that are older than the criteria specified */

  IF @strPeriod = 'dd'
  BEGIN
    DELETE FROM ASRSysEventLog WHERE [DateTime] < DATEADD(dd,-@intfrequency,getdate())
  END

  IF @strPeriod = 'wk'
  BEGIN
    DELETE FROM ASRSysEventLog WHERE [DateTime] < DATEADD(wk,-@intfrequency,getdate())
  END

  IF @strPeriod = 'mm'
  BEGIN
    DELETE FROM ASRSysEventLog WHERE [DateTime] < DATEADD(mm,-@intfrequency,getdate())
  END

  IF @strPeriod = 'yy'
  BEGIN
    DELETE FROM ASRSysEventLog WHERE [DateTime] < DATEADD(yy,-@intfrequency,getdate())
  END

  /* Delete the child rows for the header records we have just deleted */
  DELETE FROM ASRSysEventLogDetails WHERE [EventLogID] NOT IN (SELECT ID FROM ASRSysEventLog)

END
GO

