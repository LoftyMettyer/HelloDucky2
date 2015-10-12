CREATE PROCEDURE [dbo].[spASRIntGetSSIWelcomeDetails]	
(
		@piWelcomeColumnID integer,
		@piPhotographColumnID integer,
		@piSingleRecordViewID integer,
		@psUserName varchar(255),	
		@psWelcomeMessage varchar(255) OUTPUT,
		@psSelfServiceWelcomeColumn varchar(255) OUTPUT,
		@psSelfServicePhotograph varbinary(max) OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE @sql nvarchar(max)
	DECLARE @dtLastLogon datetime
	DECLARE @myval varchar(max)
	DECLARE @myvalVarBinary as varbinary(max)
	DECLARE @psLogonTime varchar(20)
	DECLARE @psLogonDay varchar(20)
	DECLARE @psWelcomeName varchar(255)
	DECLARE @psLastLogon varchar(50)		

	--- try to get the users welcome name

	BEGIN TRY
		SELECT @sql = 'SELECT @outparm = ['+c.columnname+'] FROM ['+v.viewname+']'
			FROM ASRSysColumns c, ASRSysViews v
			WHERE c.columnID = @piWelcomeColumnID AND v.ViewID = @piSingleRecordViewID

		EXEC sp_executesql @sql, N'@outparm nvarchar(max) output', @myval OUTPUT
	
		IF LEN(LTRIM(RTRIM(@myval))) = 0 OR @@ROWCOUNT = 0 or ISNULL(@myval, '') = ''
		BEGIN
			SET @psWelcomeName = ''
		END
		ELSE
		BEGIN
			SET @psWelcomeName = ' ' + isnull(@myval, '')
		END

	END TRY
	
	BEGIN CATCH
		SET @psWelcomeName = ''
	END CATCH
	
	-- Get the user's photograph
	BEGIN TRY
		SELECT @sql = 'SELECT @outparm = ['+c.columnname+'] FROM ['+v.viewname+']'
			FROM ASRSysColumns c, ASRSysViews v
			WHERE c.columnID = @piPhotographColumnID AND v.ViewID = @piSingleRecordViewID

		EXEC sp_executesql @sql, N'@outparm varbinary(max) output', @myvalVarBinary OUTPUT
	
		SET @psSelfServicePhotograph = @myvalVarBinary
	END TRY
	
	BEGIN CATCH
		SET @psSelfServicePhotograph = null
	END CATCH

	--- Now get the last logon details

	SELECT TOP 1 @dtLastLogon = DateTimeStamp
				FROM ASRSysAuditAccess WHERE [UserName] = @psUserName
				AND [HRProModule] in ('OpenHR Web', 'Intranet', 'Self-service')
				AND [Action] = 'log in'
							--AND ID NOT IN (                  
							--								SELECT top 1 ID
							--								FROM ASRSysAuditAccess WHERE [UserName] = @psUserName															
							--								AND [HRProModule] in ('OpenHR Web', 'Intranet', 'Self-service')
							--								AND [Action] = 'log in'
							--								ORDER BY DateTimeStamp DESC)                  
	ORDER BY DateTimeStamp DESC
			

	IF @@ROWCOUNT > 0 
	BEGIN
		SET @psLogonTime = CONVERT(varchar(5),@dtLastLogon, 108)
		SELECT @psLogonDay = 
			CASE datediff(day, @dtLastLogon, GETDATE())
			WHEN 0 THEN 'today'
			WHEN 1 THEN 'yesterday'
			ELSE 'on ' + CAST(DAY(@dtLastLogon) AS VARCHAR(2)) + ' ' + DATENAME(MM, @dtLastLogon) + ' ' + CAST(YEAR(@dtLastLogon) AS VARCHAR(4))
		END
		SET @psWelcomeMessage = 'Welcome back' + @psWelcomeName + ', you last logged in at ' + @psLogonTime + ' ' + @psLogonDay
	END
	ELSE
	BEGIN
		SET @psWelcomeMessage = 'Welcome ' + @psWelcomeName
	END

	SET @psSelfServiceWelcomeColumn = @psWelcomeName;

END
