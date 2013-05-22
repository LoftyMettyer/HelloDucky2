CREATE PROCEDURE [dbo].[sp_ASRIntGetEmpIDFromTBID] (
	@piEmpRecordID	integer		OUTPUT,
	@piTBRecordID	integer
)
AS
BEGIN
	DECLARE @iUserGroupID		integer,
		@sUserGroupName			sysname,
		@iChildViewID 			integer,
		@sTempExecString		nvarchar(MAX),
		@sTempParamDefinition	nvarchar(500),
		@iEmpTableID			integer,
		@iTBTableID				integer,
		@sTBRealSource			varchar(255),
		@sTBTableName			sysname,
		@sActualUserName		sysname;

	/* Get the current user's group ID. */
	EXEC [dbo].[spASRIntGetActualUserDetails]
		@sActualUserName OUTPUT,
		@sUserGroupName OUTPUT,
		@iUserGroupID OUTPUT;

	/* NB. To reach this point we have already checked that the user has 'read' permission on the Waiting List table. */
	SELECT @iEmpTableID = convert(integer, parameterValue)
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_TRAININGBOOKING'
		AND parameterKey = 'Param_EmployeeTable'
	IF @iEmpTableID IS NULL SET @iEmpTableID = 0;

	SELECT @iTBTableID = convert(integer, parameterValue)
	FROM ASRSysModuleSetup
	WHERE moduleKey = 'MODULE_TRAININGBOOKING'
		AND parameterKey = 'Param_TrainBookTable'
	IF @iTBTableID IS NULL SET @iTBTableID = 0;

	SELECT @sTBTableName = tableName
	FROM ASRSysTables
	WHERE tableID = @iTBTableID;

	SELECT @iChildViewID = childViewID
	FROM ASRSysChildViews2
	WHERE tableID = @iTBTableID
		AND role = @sUserGroupName;
		
	IF @iChildViewID IS null SET @iChildViewID = 0;
		
	IF @iChildViewID > 0 
	BEGIN
		SET @sTBRealSource = 'ASRSysCV' + 
			convert(varchar(1000), @iChildViewID) +
			'#' + replace(@sTBTableName, ' ', '_') +
			'#' + replace(@sUserGroupName, ' ', '_');
		SET @sTBRealSource = left(@sTBRealSource, 255);
	END

	SET @sTempExecString = 'SELECT @iID = ID_' + convert(nvarchar(100), @iEmpTableID) +
		' FROM ' + @sTBRealSource +
		' WHERE id = ' + convert(nvarchar(100), @piTBRecordID);
	SET @sTempParamDefinition = N'@iID integer OUTPUT';
	EXEC sp_executesql @sTempExecString, @sTempParamDefinition, @piEmpRecordID OUTPUT;

	IF @piEmpRecordID IS null SET @piEmpRecordID = 0;
END
































GO

