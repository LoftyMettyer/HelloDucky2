CREATE PROCEDURE sp_ASRLockWrite (@LockType int, @Module int, @NotifyGroups nvarchar(MAX))
AS
BEGIN

	DECLARE @LockDesc varchar(50)
	DECLARE @OrigTranCount int

	SELECT @LockDesc = case @LockType
		WHEN 1 THEN 'Saving'
		WHEN 2 THEN 'Manual'
		WHEN 3 THEN 'Read Write'
		ELSE ''
	END

	IF @LockDesc <> ''
	BEGIN

		SET @OrigTranCount = @@trancount
		IF @OrigTranCount = 0 BEGIN TRANSACTION

		DELETE FROM ASRSysLock WHERE Priority = @LockType

		INSERT ASRSysLock (Priority, Description, Username, Hostname, Lock_Time, Login_Time, SPID, Module, NotifyGroups)
		SELECT @LockType, @LockDesc, system_user, host_name(), getdate(), Login_Time, @@spid, @Module, ISNULL(@NotifyGroups, '')
			FROM master..sysprocesses WHERE spid = @@spid

		IF @OrigTranCount = 0 COMMIT TRANSACTION

	END

END

