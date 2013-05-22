CREATE PROCEDURE spASRCheckNTLogin 
		(
		    @sLoginName varchar(800)
		)
		AS 
		BEGIN
		
			DECLARE @hResult integer
			DECLARE @strFoundName varchar(800)
			DECLARE @bFound bit
			
			SELECT @strFoundName = name from master..syslogins where name = @sLoginName and isntname = 1
		
			SET @bFound = 0	
			IF (@strFoundName IS NULL)
				BEGIN
					EXEC @hResult = sp_grantlogin @sLoginName
		
					IF @hResult = 0
						BEGIN
							EXEC sp_revokelogin @sLoginName
							SET @bFound = 1
						END
					ELSE
						BEGIN
							SET @bFound = 0
						END
		
				END
			ELSE 
				SET @bFound = 1
		
			SELECT @bFound
		
		END
GO

