CREATE PROCEDURE [dbo].[spASRIntGetSelfServiceRecordID] (
	@piRecordID		integer 		OUTPUT,
	@piRecordCount	integer 		OUTPUT,
	@piViewID		integer
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE	@sViewName		sysname,
		@sCommand			nvarchar(MAX),
		@sParamDefinition	nvarchar(500),
		@iRecordID			integer,
		@iRecordCount		integer, 
		@fSysSecMgr			bit,
		@fAccessGranted		bit;
		
	SET @iRecordID = 0;
	SET @iRecordCount = 0;

	SELECT @sViewName = viewName
		FROM ASRSysViews
		WHERE viewID = @piViewID;

	IF len(@sViewName) > 0
	BEGIN
		/* Check if the user has permission to read the Self-service view. */
		exec spASRIntSysSecMgr @fSysSecMgr OUTPUT;

		IF @fSysSecMgr = 1
		BEGIN
			SET @fAccessGranted = 1;
		END
		ELSE
		BEGIN
		
			SELECT @fAccessGranted =
				CASE p.protectType
					WHEN 205 THEN 1
					WHEN 204 THEN 1
					ELSE 0
				END 
			FROM #sysprotects p
			INNER JOIN sysobjects ON p.id = sysobjects.id
			INNER JOIN syscolumns ON p.id = syscolumns.id
			WHERE p.action = 193 
				AND syscolumns.name = 'ID'
				AND sysobjects.name = @sViewName
				AND (((convert(tinyint,substring(p.columns,1,1))&1) = 0
				AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) != 0)
				OR ((convert(tinyint,substring(p.columns,1,1))&1) != 0
				AND (convert(int,substring(p.columns,sysColumns.colid/8+1,1))&power(2,sysColumns.colid&7)) = 0));
		END
	
		IF @fAccessGranted = 1
		BEGIN
			SET @sCommand = 'SELECT @iValue = COUNT(ID)' + 
				' FROM ' + @sViewName;
			SET @sParamDefinition = N'@iValue integer OUTPUT';
			EXEC sp_executesql @sCommand,  @sParamDefinition, @iRecordCount OUTPUT;

			IF @iRecordCount = 1 
			BEGIN
				SET @sCommand = 'SELECT @iValue = ' + @sViewName + '.ID ' + 
					' FROM ' + @sViewName;
				SET @sParamDefinition = N'@iValue integer OUTPUT';
				EXEC sp_executesql @sCommand,  @sParamDefinition, @iRecordID OUTPUT;
			END
		END
	END

	SET @piRecordID = @iRecordID;
	SET @piRecordCount = @iRecordCount;
END