CREATE PROCEDURE [dbo].[spASRIntGetExprFunctions] (
	@piTableID 		integer,
	@pbAbsenceEnabled	bit
)
AS
BEGIN

	SET NOCOUNT ON;

	/* Return a recordset of tab-delimited runtime function definitions ;
	<function id><tab><function name><tab><function category> */
	DECLARE @iTemp 					integer,
		@iPersonnelTableID		integer,
		@iHierarchyTableID		integer,
		@iPostAllocationTableID	integer,
		@iIdentifyingColumnID	integer,
		@iReportsToColumnID		integer,
		@iLoginColumnID			integer,
		@iSecondLoginColumnID	integer,
		@fIsPostSubOfOK			bit = 0,
		@fIsPostSubOfUserOK		bit = 0,
		@fIsPersSubOfOK			bit = 0,
		@fIsPersSubOfUserOK		bit = 0,
		@fHasPostSubOK			bit = 0,
		@fHasPostSubUserOK		bit = 0,
		@fHasPersSubOK			bit = 0,
		@fHasPersSubUserOK		bit = 0,
		@fPostBased				bit = 0, 
		@sSQLVersion			integer,
		@fBaseTablePersonnelOK	bit = 0;
	
	SELECT @iPersonnelTableID = convert(integer, isnull(parameterValue, '0'))
	FROM ASRSysModuleSetup 
	WHERE moduleKey = 'MODULE_PERSONNEL' 
		AND parameterKey = 'Param_TablePersonnel';

	SELECT @iHierarchyTableID = convert(integer, isnull(parameterValue, '0'))
	FROM ASRSysModuleSetup 
	WHERE moduleKey = 'MODULE_HIERARCHY' 
		AND parameterKey = 'Param_TableHierarchy';

	SELECT @iIdentifyingColumnID = convert(integer, isnull(parameterValue, '0'))
	FROM ASRSysModuleSetup 
	WHERE moduleKey = 'MODULE_HIERARCHY' 
		AND parameterKey = 'Param_FieldIdentifier';

	SELECT @iReportsToColumnID = convert(integer, isnull(parameterValue, '0'))
	FROM ASRSysModuleSetup 
	WHERE moduleKey = 'MODULE_HIERARCHY' 
		AND parameterKey = 'Param_FieldReportsTo';

	SELECT @iPostAllocationTableID = convert(integer, isnull(parameterValue, '0'))
	FROM ASRSysModuleSetup 
	WHERE moduleKey = 'MODULE_HIERARCHY' 
		AND parameterKey = 'Param_TablePostAllocation';

	SELECT @iLoginColumnID = convert(integer, isnull(parameterValue, '0'))
	FROM ASRSysModuleSetup 
	WHERE moduleKey = 'MODULE_PERSONNEL' 
		AND parameterKey = 'Param_FieldsLoginName';

	SELECT @iSecondLoginColumnID = convert(integer, isnull(parameterValue, '0'))
	FROM ASRSysModuleSetup 
	WHERE moduleKey = 'MODULE_PERSONNEL' 
		AND parameterKey = 'Param_FieldsSecondLoginName';

	IF (@iLoginColumnID = 0) AND (@iSecondLoginColumnID > 0)
	BEGIN
		SET @iLoginColumnID = @iSecondLoginColumnID;
		SET @iSecondLoginColumnID = 0;
	END

	IF @iPersonnelTableID <> @iHierarchyTableID SET @fPostBased = 1;
	IF @iPersonnelTableID = @piTableID 
	BEGIN
		IF (@iIdentifyingColumnID > 0) AND
			(@iReportsToColumnID > 0) AND
			((@fPostBased = 0) OR (@iPersonnelTableID > 0)) AND
			((@fPostBased = 0) OR (@iPostAllocationTableID > 0)) 
		BEGIN
			SET @fIsPersSubOfOK = 1;
			SET @fHasPersSubOK = 1;
		END
		IF (@iIdentifyingColumnID > 0) AND
			(@iReportsToColumnID > 0) AND
			(@iPersonnelTableID > 0) AND
			(@iLoginColumnID > 0) AND
			((@fPostBased = 0) OR (@iPostAllocationTableID > 0)) 
		BEGIN
			SET @fIsPersSubOfUserOK = 1;
			SET @fHasPersSubUserOK = 1;
		END
	END
				
	IF @iHierarchyTableID = @piTableID 
	BEGIN
		IF (@iIdentifyingColumnID > 0) AND
			(@iReportsToColumnID > 0) AND
			(@fPostBased = 1)
		BEGIN
			SET @fIsPostSubOfOK = 1;
			SET @fHasPostSubOK = 1;
		END
		IF (@iIdentifyingColumnID > 0) AND
			(@iReportsToColumnID > 0) AND
			(@iPersonnelTableID > 0) AND
			(@iLoginColumnID > 0) AND
			(@fPostBased = 1) AND
			(@iPostAllocationTableID > 0)
		BEGIN
			SET @fIsPostSubOfUserOK = 1;
			SET @fHasPostSubUserOK = 1;
		END
	END
	IF @iPersonnelTableID = @piTableID 
	BEGIN
		SET @fBaseTablePersonnelOK = 1;
	END
	ELSE
	BEGIN
		SELECT @iTemp = COUNT(*)
		FROM ASRSysRelations
		WHERE parentID = @iPersonnelTableID
			AND childID = @piTableID;
		IF @iTemp > 0
		BEGIN
			SET @fBaseTablePersonnelOK = 1;
		END
	END

	SELECT 
		convert(varchar(255), functionID) + char(9) +
		functionName + 
		CASE 
			WHEN len(shortcutKeys) > 0 THEN ' ' + shortcutKeys
			ELSE ''
		END + char(9) +
		category AS [definitionString]
	FROM ASRSysFunctions
	WHERE (runtime = 1 OR UDF = 1)
		AND ((functionID <> 65) OR (@fIsPostSubOfOK = 1))
		AND ((functionID <> 66) OR (@fIsPostSubOfUserOK = 1))
		AND ((functionID <> 67) OR (@fIsPersSubOfOK = 1))
		AND ((functionID <> 68) OR (@fIsPersSubOfUserOK = 1))
		AND ((functionID <> 69) OR (@fHasPostSubOK = 1))
		AND ((functionID <> 70) OR (@fHasPostSubUserOK = 1))
		AND ((functionID <> 71) OR (@fHasPersSubOK = 1))
		AND ((functionID <> 72) OR (@fHasPersSubUserOK = 1))
		AND ((functionID <> 30) OR (@fBaseTablePersonnelOK = 1))
		AND ((functionID <> 46) OR (@fBaseTablePersonnelOK = 1))
		AND ((functionID <> 47) OR (@fBaseTablePersonnelOK = 1))
		AND ((functionID <> 73) OR ((@fBaseTablePersonnelOK = 1) AND (@pbAbsenceEnabled = 1)));
END
