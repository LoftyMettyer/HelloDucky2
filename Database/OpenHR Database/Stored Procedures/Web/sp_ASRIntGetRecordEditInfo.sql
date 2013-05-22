CREATE PROCEDURE [dbo].[sp_ASRIntGetRecordEditInfo] (
	@psTitle 		varchar(500) 	OUTPUT, 
	@pfQuickEntry	bit				OUTPUT, 
	@piScreenID 	integer, 
	@piViewID 		integer
)
AS
BEGIN
	/* Return the OUTPUT variable @psTitle with the Record Edit window title for the given screen/view . 
	    The title is in the format <screen name>[ - <view name> view)]
	*/
	DECLARE @sScreenName	sysname,
			@sViewName		sysname;
	
	/* The title always starts with the screen name. */
	SELECT @psTitle = ASRSysScreens.name,
		@pfQuickEntry = ASRSysScreens.quickEntry
	FROM ASRSysScreens
	WHERE ASRSysScreens.screenID = @piScreenID;

	IF @psTitle IS NULL 
	BEGIN
		SET @psTitle = '<unknown screen>';
	END
	IF @pfQuickEntry IS NULL 
	BEGIN
		SET @pfQuickEntry = 0;
	END

	IF @piViewID > 0
	BEGIN
		/* Find title is the table name with the view name in brackets. */
		SELECT @sViewName = ASRSysViews.viewName
		FROM ASRSysViews
		WHERE ASRSysViews.viewID = @piViewID;

		IF (@sViewName IS NULL) 
		BEGIN
			SET @psTitle = @psTitle + ' (<unknown view>)';
		END
		ELSE
		BEGIN
			SET @psTitle = @psTitle + ' (' + @sViewName + ' view)';
		END
	END
END