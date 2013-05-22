CREATE PROCEDURE [dbo].[sp_ASRIntGetExprFunctions] (
	@piTableID 		integer
)
AS
BEGIN
	/* Return a recordset of tab-delimited runtime function definitions ;
	<function id><tab><function name><tab><function category> */
	DECLARE @fEnableUDFFunctions	bit,
		@fNewModuleCode			bit,
		@fAbsenceEnabled		bit,
		@iTemp 					integer,
		@iPersonnelTableID		integer,
		@iHierarchyTableID		integer,
		@iPostAllocationTableID	integer,
		@iIdentifyingColumnID	integer,
		@iReportsToColumnID		integer,
		@iLoginColumnID			integer,
		@iSecondLoginColumnID	integer,
		@fIsPostSubOfOK			bit,
		@fIsPostSubOfUserOK		bit,
		@fIsPersSubOfOK			bit,
		@fIsPersSubOfUserOK		bit,
		@fHasPostSubOK			bit,
		@fHasPostSubUserOK		bit,
		@fHasPersSubOK			bit,
		@fHasPersSubUserOK		bit,
		@fPostBased				bit, 
		@sSQLVersion			integer,
		@fBaseTablePersonnelOK	bit;

	SET @fEnableUDFFunctions = 0;
	SET @fIsPostSubOfOK = 0;
	SET @fIsPostSubOfUserOK = 0;
	SET @fIsPersSubOfOK = 0;
	SET @fIsPersSubOfUserOK = 0;
	SET @fHasPostSubOK = 0;
	SET @fHasPostSubUserOK = 0;
	SET @fHasPersSubOK = 0;
	SET @fHasPersSubUserOK = 0;
	SET @fPostBased = 0;
	SET @fBaseTablePersonnelOK = 0;
	
	SELECT @sSQLVersion = dbo.udfASRSQLVersion();

	IF @sSQLVersion >= 8
	BEGIN  
		SET @fEnableUDFFunctions = 1;
		
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
	END

	-- Activate module
	EXEC [dbo].[spASRIntActivateModule] 'ABSENCE', @fAbsenceEnabled OUTPUT
	SELECT 
		convert(varchar(255), functionID) + char(9) +
		functionName + 
		CASE 
			WHEN len(shortcutKeys) > 0 THEN ' ' + shortcutKeys
			ELSE ''
		END + char(9) +
		category AS [definitionString]
	FROM ASRSysFunctions
	WHERE ((runtime = 1)	
			OR ((UDF = 1) AND (@fEnableUDFFunctions = 1)))
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
		AND ((functionID <> 73) OR ((@fBaseTablePersonnelOK = 1) AND (@fAbsenceEnabled = 1)));
END
