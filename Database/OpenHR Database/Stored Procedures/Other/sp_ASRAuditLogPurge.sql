CREATE PROCEDURE [dbo].[sp_ASRAuditLogPurge]
AS
BEGIN

	DECLARE @intFrequency	integer,
			@strPeriod		char(2);

	SET @strPeriod = null;
	SET @intFrequency = null;

	SELECT @intFrequency = Frequency
	FROM ASRSysAuditCleardown
	WHERE [Type] = 'Users';

	SELECT @strPeriod = Period
	FROM ASRSysAuditCleardown
	WHERE [Type] = 'Users';

	IF (@intFrequency IS NOT NULL) AND (@strPeriod IS NOT NULL)
	BEGIN

		IF @strPeriod = 'dd'
			DELETE FROM ASRSysAuditGroup WHERE [DateTimeStamp] < DATEADD(dd,-@intfrequency,getdate());

		IF @strPeriod = 'wk'
			DELETE FROM ASRSysAuditGroup WHERE [DateTimeStamp] < DATEADD(wk,-@intfrequency,getdate());

		IF @strPeriod = 'mm'
			DELETE FROM ASRSysAuditGroup WHERE [DateTimeStamp] < DATEADD(mm,-@intfrequency,getdate());

		IF @strPeriod = 'yy'
			DELETE FROM ASRSysAuditGroup WHERE [DateTimeStamp] < DATEADD(yy,-@intfrequency,getdate());

	END

	SET @strPeriod = null;
	SET @intFrequency = null;

	SELECT @intFrequency = Frequency
	FROM ASRSysAuditCleardown
	WHERE [Type] = 'Permissions';

	SELECT @strPeriod = Period
	FROM ASRSysAuditCleardown
	WHERE [Type] = 'Permissions';

	IF (@intFrequency IS NOT NULL) AND (@strPeriod IS NOT NULL)
	BEGIN
		IF @strPeriod = 'dd'
			DELETE FROM ASRSysAuditPermissions WHERE [DateTimeStamp] < DATEADD(dd,-@intfrequency,getdate());

		IF @strPeriod = 'wk'
			DELETE FROM ASRSysAuditPermissions WHERE [DateTimeStamp] < DATEADD(wk,-@intfrequency,getdate());

		IF @strPeriod = 'mm'
			DELETE FROM ASRSysAuditPermissions WHERE [DateTimeStamp] < DATEADD(mm,-@intfrequency,getdate());

		IF @strPeriod = 'yy'
			DELETE FROM ASRSysAuditPermissions WHERE [DateTimeStamp] < DATEADD(yy,-@intfrequency,getdate());
	END

	SET @strPeriod = null;
	SET @intFrequency = null;

	SELECT @intFrequency = Frequency
	FROM ASRSysAuditCleardown
	WHERE [Type] = 'Data';

	SELECT @strPeriod = Period
	FROM ASRSysAuditCleardown
	WHERE [Type] = 'Data';

	IF (@intFrequency IS NOT NULL) AND (@strPeriod IS NOT NULL)
	BEGIN
		IF @strPeriod = 'dd'
			DELETE FROM ASRSysAuditTrail  WHERE [DateTimeStamp] < DATEADD(dd,-@intfrequency,getdate());

		IF @strPeriod = 'wk'
			DELETE FROM ASRSysAuditTrail WHERE [DateTimeStamp] < DATEADD(wk,-@intfrequency,getdate());

		IF @strPeriod = 'mm'
			DELETE FROM ASRSysAuditTrail WHERE [DateTimeStamp] < DATEADD(mm,-@intfrequency,getdate());

		IF @strPeriod = 'yy'
			DELETE FROM ASRSysAuditTrail WHERE [DateTimeStamp] < DATEADD(yy,-@intfrequency,getdate());

	END

	SET @strPeriod = null;
	SET @intFrequency = null;

	SELECT @intFrequency = Frequency
	FROM ASRSysAuditCleardown
	WHERE [Type] = 'Access';

	SELECT @strPeriod = Period
	FROM ASRSysAuditCleardown
	WHERE [Type] = 'Access';

	IF (@intFrequency IS NOT NULL) AND (@strPeriod IS NOT NULL)
	BEGIN

		IF @strPeriod = 'dd'
			DELETE FROM ASRSysAuditAccess WHERE [DateTimeStamp] < DATEADD(dd,-@intfrequency,getdate());

		IF @strPeriod = 'wk'
			DELETE FROM ASRSysAuditAccess WHERE [DateTimeStamp] < DATEADD(wk,-@intfrequency,getdate());

		IF @strPeriod = 'mm'
			DELETE FROM ASRSysAuditAccess WHERE [DateTimeStamp] < DATEADD(mm,-@intfrequency,getdate());

		IF @strPeriod = 'yy'
			DELETE FROM ASRSysAuditAccess WHERE [DateTimeStamp] < DATEADD(yy,-@intfrequency,getdate());

	END
END