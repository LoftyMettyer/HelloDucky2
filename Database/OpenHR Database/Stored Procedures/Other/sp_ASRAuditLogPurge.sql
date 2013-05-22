CREATE PROCEDURE [dbo].[sp_ASRAuditLogPurge]
AS
BEGIN

	DECLARE @intFrequency	integer,
			@strPeriod		char(2);

	SET @strPeriod = null;
	SET @intFrequency = null;

	SELECT @intFrequency = Frequency
	FROM AsrSysAuditCleardown
	WHERE [Type] = 'Users';

	SELECT @strPeriod = Period
	FROM AsrSysAuditCleardown
	WHERE [Type] = 'Users';

	IF (@intFrequency IS NOT NULL) AND (@strPeriod IS NOT NULL)
	BEGIN

		IF @strPeriod = 'dd'
			DELETE FROM AsrSysAuditGroup WHERE [DateTimeStamp] < DATEADD(dd,-@intfrequency,getdate());

		IF @strPeriod = 'wk'
			DELETE FROM AsrSysAuditGroup WHERE [DateTimeStamp] < DATEADD(wk,-@intfrequency,getdate());

		IF @strPeriod = 'mm'
			DELETE FROM AsrSysAuditGroup WHERE [DateTimeStamp] < DATEADD(mm,-@intfrequency,getdate());

		IF @strPeriod = 'yy'
			DELETE FROM AsrSysAuditGroup WHERE [DateTimeStamp] < DATEADD(yy,-@intfrequency,getdate());

	END

	SET @strPeriod = null;
	SET @intFrequency = null;

	SELECT @intFrequency = Frequency
	FROM AsrSysAuditCleardown
	WHERE [Type] = 'Permissions';

	SELECT @strPeriod = Period
	FROM AsrSysAuditCleardown
	WHERE [Type] = 'Permissions';

	IF (@intFrequency IS NOT NULL) AND (@strPeriod IS NOT NULL)
	BEGIN
		IF @strPeriod = 'dd'
			DELETE FROM AsrSysAuditPermissions WHERE [DateTimeStamp] < DATEADD(dd,-@intfrequency,getdate());

		IF @strPeriod = 'wk'
			DELETE FROM AsrSysAuditPermissions WHERE [DateTimeStamp] < DATEADD(wk,-@intfrequency,getdate());

		IF @strPeriod = 'mm'
			DELETE FROM AsrSysAuditPermissions WHERE [DateTimeStamp] < DATEADD(mm,-@intfrequency,getdate());

		IF @strPeriod = 'yy'
			DELETE FROM AsrSysAuditPermissions WHERE [DateTimeStamp] < DATEADD(yy,-@intfrequency,getdate());
	END

	SET @strPeriod = null;
	SET @intFrequency = null;

	SELECT @intFrequency = Frequency
	FROM AsrSysAuditCleardown
	WHERE [Type] = 'Data';

	SELECT @strPeriod = Period
	FROM AsrSysAuditCleardown
	WHERE [Type] = 'Data';

	IF (@intFrequency IS NOT NULL) AND (@strPeriod IS NOT NULL)
	BEGIN
		IF @strPeriod = 'dd'
			DELETE FROM AsrSysAuditTrail  WHERE [DateTimeStamp] < DATEADD(dd,-@intfrequency,getdate());

		IF @strPeriod = 'wk'
			DELETE FROM AsrSysAuditTrail WHERE [DateTimeStamp] < DATEADD(wk,-@intfrequency,getdate());

		IF @strPeriod = 'mm'
			DELETE FROM AsrSysAuditTrail WHERE [DateTimeStamp] < DATEADD(mm,-@intfrequency,getdate());

		IF @strPeriod = 'yy'
			DELETE FROM AsrSysAuditTrail WHERE [DateTimeStamp] < DATEADD(yy,-@intfrequency,getdate());

	END

	SET @strPeriod = null;
	SET @intFrequency = null;

	SELECT @intFrequency = Frequency
	FROM AsrSysAuditCleardown
	WHERE [Type] = 'Access';

	SELECT @strPeriod = Period
	FROM AsrSysAuditCleardown
	WHERE [Type] = 'Access';

	IF (@intFrequency IS NOT NULL) AND (@strPeriod IS NOT NULL)
	BEGIN

		IF @strPeriod = 'dd'
			DELETE FROM AsrSysAuditAccess WHERE [DateTimeStamp] < DATEADD(dd,-@intfrequency,getdate());

		IF @strPeriod = 'wk'
			DELETE FROM AsrSysAuditAccess WHERE [DateTimeStamp] < DATEADD(wk,-@intfrequency,getdate());

		IF @strPeriod = 'mm'
			DELETE FROM AsrSysAuditAccess WHERE [DateTimeStamp] < DATEADD(mm,-@intfrequency,getdate());

		IF @strPeriod = 'yy'
			DELETE FROM AsrSysAuditAccess WHERE [DateTimeStamp] < DATEADD(yy,-@intfrequency,getdate());

	END
END