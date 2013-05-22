CREATE PROCEDURE [dbo].[sp_ASRIntValidateExpression] (
	@psUtilName 			varchar(255), 
	@piUtilID 				integer, 
	@piUtilType 			integer, 
	@psUtilOwner 			sysname, 
	@piBaseTableID 			integer, 
	@psComponentDefn		varchar(MAX),
	@piTimestamp 			integer, 
	@psDeletedKeys			varchar(MAX)	OUTPUT,	/*	Tab-delimited string of node keys of any deleted calcs/filters. */
	@psHiddenOwnerKeys		varchar(MAX)	OUTPUT,	/*	Tab-delimited string of node keys of any hidden calcs/filters that the current user owns. */
	@psHiddenNotOwnerKeys	varchar(MAX)	OUTPUT,	/*	Tab-delimited string of node keys of any hidden calcs/filters that the current user does not own. */
	@psDeletedDescs			varchar(MAX)	OUTPUT,	/*	Tab-delimited string of node descriptions of any deleted calcs/filters. */
	@psHiddenOwnerDescs		varchar(MAX)	OUTPUT,	/*	Tab-delimited string of node descriptions of any hidden calcs/filters that the current user owns. */
	@psHiddenNotOwnerDescs	varchar(MAX)	OUTPUT,	/*	Tab-delimited string of node descriptions of any hidden calcs/filters that the current user does not own. */
	@piErrorCode			integer 		OUTPUT	/* 	0 = No error (but must check the strings of keys above)
										1 = Expression deleted by another user. Save as new ? 
										2 = Made hidden/read-only by another user. Save as new ? 
										3 = Modified by another user (still writable). Overwrite ? 
										4 = Non-unique name. Save fails */
	/* 	If there are any keys in the @psDeletedKeys string then these components need to be removed from the expression. The save fails.
		If there are any keys in the @psHiddenOwnerKeys or @psHiddenNotOwnerKeys strings then
			If current user does NOT own the expression then
				the expression needs to be made hidden, and the current user cannot edit it. Save and edit fails.
			Else
				If there are any keys in the @psHiddenNotOwnerKeys string then
					the hidden components need to be removed from the expression. Save fails.
				Else
					expression must also be made hidden
	*/
)
AS
BEGIN
	DECLARE	@iTimestamp			integer,
			@sOwner				varchar(255),
			@sTemp				varchar(MAX),
			@sCompType			char(1),		/* 'U' = unknown, 'E' = expression, 'C' = component */
			@sParameter			varchar(MAX),
			@iComponentIndex	integer,
			@sTempAccess		varchar(MAX),
			@fHidden			bit,
			@sTempOwner			varchar(255),
			@sCurrentUser		sysname,
			@iCount				integer,
			@sExprID			varchar(100),
			@sName				varchar(255),
			@sTableID			varchar(100),
			@sReturnType		varchar(100),
			@sReturnSize		varchar(100),
			@sReturnDecimals	varchar(100),
			@sType				varchar(100),
			@sParentComponentID	varchar(100),
			@sUserName			varchar(255),
			@sAccess			varchar(MAX),
			@sDescription		varchar(MAX),
			@sTimestamp			varchar(100),
			@sViewInColour		varchar(100),
			@sExpandedNode		varchar(100),
			@fTemp				bit,
			@iCalculationID		integer,
			@sNodeKey			varchar(100),
			@sCompID			varchar(100),
			@sFieldColumnID		varchar(100),
			@sFieldPassBy			varchar(100),
			@sFieldSelectionTableID	varchar(100),
			@sFieldSelectionRecord	varchar(100),
			@sFieldSelectionLine	varchar(100),
			@sFieldSelectionOrderID	varchar(100),
			@sFieldSelectionFilter	varchar(MAX),
			@sFunctionID			varchar(100),
			@sCalculationID			varchar(100),
			@sOperatorID			varchar(100),
			@sValueType				varchar(100),
			@sValueCharacter		varchar(MAX),
			@sValueNumeric			varchar(100),
			@sValueLogic			varchar(100),
			@sValueDate				varchar(100),
			@sPromptDescription		varchar(MAX),
			@sPromptMask			varchar(MAX),
			@sPromptSize			varchar(100),
			@sPromptDecimals		varchar(100),
			@sFunctionReturnType	varchar(100),
			@sLookupTableID			varchar(100),
			@sLookupColumnID		varchar(100),
			@sFilterID				varchar(100),
			@sPromptDateType		varchar(100),
			@sFieldTableID			varchar(100),
			@sFieldSelectionOrderName	varchar(255),
			@sFieldSelectionFilterName	varchar(255);

	SET @sCurrentUser = SYSTEM_USER
	SET @piErrorCode = 0
	SET @psDeletedKeys = ''
	SET @psHiddenOwnerKeys = ''
	SET @psHiddenNotOwnerKeys = ''
	SET @psDeletedDescs = ''
	SET @psHiddenOwnerDescs = ''
	SET @psHiddenNotOwnerDescs = ''

	/* Loop through each component in the definition. */
	SET @sTemp = @psComponentDefn
	SET @sCompType = 'U'

	WHILE LEN(@sTemp) > 0
	BEGIN
		IF CHARINDEX(char(9), @sTemp) > 0
		BEGIN
			SET @sParameter = LEFT(@sTemp, CHARINDEX(char(9), @sTemp) - 1)
			SET @sTemp = RIGHT(@sTemp, LEN(@sTemp) - CHARINDEX(char(9), @sTemp))
		END
		ELSE
		BEGIN
			SET @sParameter = @sTemp
			SET @sTemp = ''
		END

		IF @sCompType = 'U' 
		BEGIN
			/* Reading a new component. */
			IF @sParameter = 'ROOT'
			BEGIN
				SET @sCompType = 'C'
			END
			ELSE
			BEGIN
				IF left(@sParameter, 1) = 'C'
				BEGIN
					SET @sCompType = 'E'
				END
				ELSE
				BEGIN
					SET @sCompType = 'C'
				END
			END	

			SET @iComponentIndex = 1
		END
		ELSE
		BEGIN
			IF @sCompType = 'E' 
			BEGIN
				/* Currently reading an expression. */
				IF @iComponentIndex = 1 SET @sNodeKey = @sParameter
				IF @iComponentIndex = 2 SET @sExprID = @sParameter
				IF @iComponentIndex = 3 SET @sName = @sParameter
				IF @iComponentIndex = 4 SET @sTableID = @sParameter
				IF @iComponentIndex = 5 SET @sReturnType = @sParameter
				IF @iComponentIndex = 6 SET @sReturnSize = @sParameter
				IF @iComponentIndex = 7 SET @sReturnDecimals = @sParameter
				IF @iComponentIndex = 8 SET @sType = @sParameter
				IF @iComponentIndex = 9 SET @sParentComponentID = @sParameter
				IF @iComponentIndex = 10 SET @sUserName = @sParameter
				IF @iComponentIndex = 11 SET @sAccess = @sParameter
				IF @iComponentIndex = 12 SET @sDescription = @sParameter
				IF @iComponentIndex = 13 SET @sTimestamp = @sParameter
				IF @iComponentIndex = 14 SET @sViewInColour = @sParameter
				IF @iComponentIndex = 15 
				BEGIN
					SET @sExpandedNode = @sParameter
					SET @sCompType = 'U'
				END
			END
			ELSE
			BEGIN
				/* Currently reading a component. */
				IF @iComponentIndex = 1 SET @sNodeKey = @sParameter
				IF @iComponentIndex = 2 SET @sCompID = @sParameter
				IF @iComponentIndex = 3 SET @sExprID = @sParameter
				IF @iComponentIndex = 4 SET @sType = @sParameter
				IF @iComponentIndex = 5 SET @sFieldColumnID = @sParameter
				IF @iComponentIndex = 6 SET @sFieldPassBy = @sParameter
				IF @iComponentIndex = 7 SET @sFieldSelectionTableID = @sParameter
				IF @iComponentIndex = 8 SET @sFieldSelectionRecord = @sParameter
				IF @iComponentIndex = 9 SET @sFieldSelectionLine = @sParameter
				IF @iComponentIndex = 10 SET @sFieldSelectionOrderID = @sParameter
				IF @iComponentIndex = 11 SET @sFieldSelectionFilter = @sParameter
				IF @iComponentIndex = 12 SET @sFunctionID = @sParameter
				IF @iComponentIndex = 13 SET @sCalculationID = @sParameter
				IF @iComponentIndex = 14 SET @sOperatorID = @sParameter
				IF @iComponentIndex = 15 SET @sValueType = @sParameter
				IF @iComponentIndex = 16 SET @sValueCharacter = @sParameter
				IF @iComponentIndex = 17 SET @sValueNumeric = @sParameter
				IF @iComponentIndex = 18 SET @sValueLogic = @sParameter
				IF @iComponentIndex = 19 SET @sValueDate = @sParameter
				IF @iComponentIndex = 20 SET @sPromptDescription = @sParameter
				IF @iComponentIndex = 21 SET @sPromptMask = @sParameter
				IF @iComponentIndex = 22 SET @sPromptSize = @sParameter
				IF @iComponentIndex = 23 SET @sPromptDecimals = @sParameter
				IF @iComponentIndex = 24 SET @sFunctionReturnType = @sParameter
				IF @iComponentIndex = 25 SET @sLookupTableID = @sParameter
				IF @iComponentIndex = 26 SET @sLookupColumnID = @sParameter
				IF @iComponentIndex = 27 SET @sFilterID = @sParameter
				IF @iComponentIndex = 28 SET @sExpandedNode = @sParameter
				IF @iComponentIndex = 29 SET @sPromptDateType = @sParameter
				IF @iComponentIndex = 30 SET @sDescription = @sParameter
				IF @iComponentIndex = 31 SET @sFieldTableID = @sParameter
				IF @iComponentIndex = 32 SET @sFieldSelectionOrderName = @sParameter
				IF @iComponentIndex = 33 
				BEGIN
					SET @sFieldSelectionFilterName = @sParameter
					SET @sCompType = 'U'
					IF (@sType = '3') 
						OR (@sType = '10') 
						OR ((@sType = '1') AND (convert(integer, @sFieldSelectionFilter) > 0))
					BEGIN
						/* Check if the calculation/filter still exists and hasn't been made hidden. */
						IF (@sType = '3') 
						BEGIN
							SET @iCalculationID = convert(integer, @sCalculationID)
						END
						ELSE
						BEGIN
							IF (@sType = '10') 
							BEGIN
								SET @iCalculationID = convert(integer, @sFilterID)
							END
							ELSE
							BEGIN
								SET @iCalculationID = convert(integer, @sFieldSelectionFilter)
							END
						END

						SELECT @sTempAccess = access ,
							 @sTempOwner = userName 
						FROM ASRSysExpressions
						WHERE exprID = @iCalculationID
						IF @sTempAccess IS null
						BEGIN
							/* Calculation has been deleted. */
							SET @psDeletedKeys = @psDeletedKeys +
								CASE 
									WHEN len(@psDeletedKeys) > 0 THEN char(9)
									ELSE ''
								END +
								@sNodeKey
							SET @psDeletedDescs = @psDeletedDescs +
								CASE 
									WHEN len(@psDeletedDescs) > 0 THEN char(9)
									ELSE ''
								END +
								@sDescription
						END
						ELSE 
						BEGIN
							/* Calculation still exists. Is it hidden ? */
							SET @fHidden = 0
							
							IF @sTempAccess = 'HD'
							BEGIN
								SET @fHidden = 1
							END
							ELSE
							BEGIN
								/* The calc isn't hidden. Are any sub-components hidden ? */
								execute sp_ASRIntExpressionHasHiddenComponents @iCalculationID, @fTemp OUTPUT

								IF @fTemp = 1 
								BEGIN
									SET @fHidden = 1
								END
							END

							IF @fHidden = 1
							BEGIN
								IF @sTempOwner = @sCurrentUser
								BEGIN
									SET @psHiddenOwnerKeys = @psHiddenOwnerKeys +
										CASE 
											WHEN len(@psHiddenOwnerKeys) > 0 THEN char(9)
											ELSE ''
										END +
										@sNodeKey
									SET  @psHiddenOwnerDescs = @psHiddenOwnerDescs +
										CASE 
											WHEN len(@psHiddenOwnerDescs) > 0 THEN char(9)
											ELSE ''
										END +
										@sDescription
								END
								ELSE
								BEGIN
									SET @psHiddenNotOwnerKeys = @psHiddenNotOwnerKeys +
										CASE 
											WHEN len(@psHiddenNotOwnerKeys) > 0 THEN char(9)
											ELSE ''
										END +
										@sNodeKey
									SET  @psHiddenNotOwnerDescs = @psHiddenNotOwnerDescs +
										CASE 
											WHEN len(@psHiddenNotOwnerDescs) > 0 THEN char(9)
											ELSE ''
										END +
										@sDescription
								END
							END
						END
					END
				END
			END		

			SET @iComponentIndex = @iComponentIndex + 1
		END
	END

	IF @piUtilID > 0
	BEGIN
		/* Check if this definition has been changed by another user. */
		SELECT @iCount = COUNT(*)
		FROM ASRSysExpressions
		WHERE exprID = @piUtilID

		IF @iCount = 0
		BEGIN
			/* Expression has been deleted by another user. Save as new ? */
			SET @piErrorCode = 1
		END
		ELSE
		BEGIN
			SELECT @iTimestamp = convert(integer, timestamp), 
				@sAccess = access, 
				@sOwner = userName
			FROM ASRSysExpressions
			WHERE exprID = @piUtilID

			IF (@iTimestamp <>@piTimestamp)
			BEGIN
				IF (@sOwner <> @sCurrentUser) AND (@sAccess <> 'RW')
				BEGIN
					/* Modified by another user, and made hidden/read-only. Save as new ? */
					SET @piErrorCode = 2
				END
				ELSE
				BEGIN
					/* Modified by another user, still writable. Overwrite ? */
					SET @piErrorCode = 3
				END
			END
		END
	END

	/* Check that the expression name is unique. */
	SELECT @iCount = COUNT(*)
	FROM ASRSysExpressions
	WHERE exprID <> @piUtilID
		AND parentComponentID = 0
		AND name = @psUtilName
		AND TableID = @piBaseTableID
		AND type = @piUtilType

	IF @iCount > 0 
	BEGIN
		SET @piErrorCode = 4
	END
END































GO

