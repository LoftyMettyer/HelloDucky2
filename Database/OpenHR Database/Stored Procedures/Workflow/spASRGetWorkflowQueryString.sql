CREATE PROCEDURE [dbo].[spASRGetWorkflowQueryString]
		(
			@piInstanceID	integer,
			@piElementID	integer,
			@psQueryString	varchar(MAX)	output
		)
		AS
		BEGIN
			DECLARE
				@hResult		integer,
				@objectToken	integer,
				@sURL			varchar(MAX),
				@sParam1		varchar(MAX),
				@sServerName	varchar(128),
				@sDBName		varchar(128),
				@sSQLVersion	integer;
		
			SET @psQueryString = '';
			SET @sSQLVersion = dbo.udfASRSQLVersion();
		
			SELECT @sURL = parameterValue
			FROM ASRSysModuleSetup
			WHERE moduleKey = 'MODULE_WORKFLOW'
				AND parameterKey = 'Param_URL';
				
			IF upper(right(@sURL, 5)) <> '.ASPX'
				AND right(@sURL, 1) <> '/'
				AND len(@sURL) > 0
			BEGIN
				SET @sURL = @sURL + '/';
			END
		
			SELECT @sParam1 = parameterValue
			FROM ASRSysModuleSetup
			WHERE moduleKey = 'MODULE_WORKFLOW'
				AND parameterKey = 'Param_Web1';
		
			IF (len(@sURL) > 0)
			BEGIN
				--Fetch server and database name using system functions to ensure consistency
				EXEC dbo.[spASRGetSQLMetadata] @sServerName OUTPUT, @sDBName OUTPUT
				
				SELECT @psQueryString = dbo.[udfASRNetGetWorkflowQueryString]( @piInstanceID, @piElementID, @sParam1, @sServerName, @sDBName);
			
				IF len(@psQueryString) > 0
				BEGIN
					SET @psQueryString = @sURL + '?' + @psQueryString;
				END
			END
		END