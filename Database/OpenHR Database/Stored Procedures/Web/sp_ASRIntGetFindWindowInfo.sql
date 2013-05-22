CREATE PROCEDURE [dbo].[sp_ASRIntGetFindWindowInfo] (
	@psTitle 		varchar(500) 	OUTPUT, 
	@pfQuickEntry 	bit 			OUTPUT, 
	@plngScreenID	integer,
	@plngViewID		integer
)
AS
BEGIN
	/* Return the OUTPUT variable @psTitle with the find window title for the given screen. 	*/
	DECLARE @sScreenName	sysname,
			@sViewName		sysname;
	
	/* Get the screen name. */
	IF @plngScreenID > 0
	BEGIN
		/* Find title is just the table name. */
		SELECT @psTitle = name,
			@pfQuickEntry = quickEntry
		FROM [dbo].[ASRSysScreens]
		WHERE screenID = @plngScreenID;

		IF @psTitle IS NULL 
		BEGIN
			SET @psTitle = '<unknown screen>';
		END
		IF @pfQuickEntry IS NULL 
		BEGIN
			SET @pfQuickEntry = 0;
		END
	END
	ELSE
	BEGIN
		SET @psTitle = '<unknown screen>';
	END	

	/* Get the view name. */
	IF @plngViewID > 0
	BEGIN
		/* Find title is just the table name. */
		SELECT @sViewName = viewName
		FROM [dbo].[ASRSysViews]
		WHERE viewID = @plngViewID;

		IF @sViewName IS NULL 
		BEGIN
			SET @psTitle = @psTitle + ' (<unknown view>)';
		END
		ELSE
		BEGIN
			SET @psTitle = @psTitle + ' (' + @sViewName + ' view)';
		END
	END
END