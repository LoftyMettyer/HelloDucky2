CREATE PROCEDURE sp_ASRLockDelete (@LockType int, @Module int)
AS
BEGIN
	DELETE FROM ASRSysLock WHERE Priority = @LockType AND Module = @Module
END

