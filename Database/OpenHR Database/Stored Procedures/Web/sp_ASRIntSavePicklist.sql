CREATE PROCEDURE [dbo].[sp_ASRIntSavePicklist] (
	@psName			varchar(255),
	@psDescription	varchar(MAX),
	@psAccess		varchar(MAX),
	@psUserName		varchar(255),
	@psColumns		varchar(MAX),
	@psColumns2		varchar(MAX),
	@piID			integer	OUTPUT,
	@piTableID		integer	
)
AS
BEGIN

	SET NOCOUNT ON;

	DECLARE
		@iIndex		integer,
		@iCount	integer,
		@sSubstring	varchar(MAX);

	DECLARE	@outputTable table (id int NOT NULL);

	/* Clean the input string parameters. */
	IF len(@psColumns) > 0 SET @psColumns = replace(@psColumns, '''', '''''')
	IF len(@psColumns2) > 0 SET @psColumns2 = replace(@psColumns2, '''', '''''')

	IF @piID = 0 
	BEGIN
		/* Saving a new picklist. */
		INSERT INTO ASRSysPickListName
			(name, description, tableID, access, userName)
		OUTPUT inserted.picklistID INTO @outputTable
		VALUES 
			(@psName, @psDescription, @piTableID, @psAccess, @psUserName)

		-- Get the ID of the inserted record.
		SELECT @piID = id FROM @outputTable;

		WHILE len(@psColumns) > 0
		BEGIN
			SET @iIndex = charindex(',', @psColumns)
	
			IF @iIndex > 0
			BEGIN
				SET @sSubstring = left(@psColumns, @iIndex -1)
				SET @psColumns = substring(@psColumns, @iIndex + 1, len(@psColumns) - @iIndex)

				INSERT INTO ASRSysPickListItems (pickListID, recordID)
				VALUES(@piID, convert(integer, @sSubstring))

				IF (len(@psColumns2) > 0) AND (len(@psColumns) < 7000)
				BEGIN
					SET @psColumns = @psColumns + left(@psColumns2, 1000)
					IF len(@psColumns2) > 1000
					BEGIN
						SET @psColumns2 = substring(@psColumns, 1001, len(@psColumns2) - 1000)
					END
					ELSE
					BEGIN
						SET @psColumns2 = ''
					END
				END
			END
			ELSE
			BEGIN
				BREAK
			END
		END

		/* Update the util access log. */
		INSERT INTO ASRSysUtilAccessLog 
			(type, utilID, createdBy, createdDate, createdHost, savedBy, savedDate, savedHost)
		VALUES (10, @piID, system_user, getdate(), host_name(), system_user, getdate(), host_name())
	END
	ELSE
	BEGIN
		/* Saving an existing picklist. */

		IF @psAccess = 'HD'
		BEGIN
			/* Hide any utilities that use this picklist. NB. The check to see if we can do this has already been done in sp_ASRIntCheckCanMakeHidden. */
			exec sp_ASRIntMakeUtilitiesHidden 10, @piID
		END
		
		DELETE FROM ASRSysPickListItems
		WHERE pickListID = @piID

		UPDATE ASRSysPickListName SET
			name = @psName, 
			description = @psDescription, 
			tableID = @piTableID,
			access = @psAccess
		WHERE pickListID = @piID

		WHILE len(@psColumns) > 0
		BEGIN
			SET @iIndex = charindex(',', @psColumns)
	
			IF @iIndex > 0
			BEGIN
				SET @sSubstring = left(@psColumns, @iIndex -1)
				SET @psColumns = substring(@psColumns, @iIndex + 1, len(@psColumns) - @iIndex)

				INSERT INTO ASRSysPickListItems (pickListID, recordID)
				VALUES(@piID, convert(integer, @sSubstring))

				IF (len(@psColumns2) > 0) AND (len(@psColumns) < 7000)
				BEGIN
					SET @psColumns = @psColumns + left(@psColumns2, 1000)
					IF len(@psColumns2) > 1000
					BEGIN
						SET @psColumns2 = substring(@psColumns, 1001, len(@psColumns2) - 1000)
					END
					ELSE
					BEGIN
						SET @psColumns2 = ''
					END
				END
			END
			ELSE
			BEGIN
				BREAK
			END
		END

		/* Update the last saved log. */
		/* Is there an entry in the log already? */
		SELECT @iCount = COUNT(*) 
		FROM ASRSysUtilAccessLog
		WHERE utilID = @piID
			AND type = 10

		IF @iCount = 0 
		BEGIN
			INSERT INTO ASRSysUtilAccessLog
 				(type, utilID, savedBy, savedDate, savedHost)
			VALUES (10, @piID, system_user, getdate(), host_name())
		END
		ELSE
		BEGIN
			UPDATE ASRSysUtilAccessLog 
			SET savedBy = system_user,
				savedDate = getdate(), 
				savedHost = host_name() 
			WHERE utilID = @piID
				AND type = 10
		END
	END
END