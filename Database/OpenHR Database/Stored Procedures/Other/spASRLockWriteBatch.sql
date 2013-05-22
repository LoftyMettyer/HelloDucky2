CREATE Procedure spASRLockWriteBatch (@BatchJobID int, @Clearlock bit, @LockedByOther int OUTPUT)
		AS
		BEGIN

			DECLARE @OrigTranCount int
			DECLARE @Realspid int

			SET @OrigTranCount = @@trancount
			IF @OrigTranCount = 0 BEGIN TRANSACTION

			SELECT @LockedByOther = COUNT(ID) FROM ASRSysBatchJobName
			JOIN master..sysprocesses syspro ON spid = LockSpid
			WHERE LockLoginTime = syspro.login_time AND LockSpid <> @@spid
			AND ID = @BatchJobID

			IF @LockedByOther = 0
			BEGIN

				--Need to get spid of parent process
				SELECT @Realspid = a.spid
				FROM master..sysprocesses a
				FULL OUTER JOIN master..sysprocesses b
					ON a.hostname = b.hostname
					AND a.hostprocess = b.hostprocess
					AND a.spid <> b.spid
				WHERE b.spid = @@Spid

				--If there is no parent spid then use current spid
				--IF @Realspid is null SET @Realspid = @@spid

				IF @Clearlock = 0
					UPDATE ASRSysBatchJobName SET
					LockSpid = @Realspid,
					LockLoginTime = (
						SELECT login_time
						FROM master..sysprocesses
						WHERE spid = @Realspid)
					WHERE ID = @BatchJobID
				ELSE
					UPDATE ASRSysBatchJobName SET
					LockSpid = 0,
					LockLoginTime = null
					WHERE ID = @BatchJobID

			END

			IF @OrigTranCount = 0 COMMIT TRANSACTION

		END
GO

