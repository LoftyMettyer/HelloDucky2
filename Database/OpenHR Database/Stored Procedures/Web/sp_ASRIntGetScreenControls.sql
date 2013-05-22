CREATE PROCEDURE [dbo].[sp_ASRIntGetScreenControls] (
	@plngScreenID 	int,
	@plngViewId	int)
AS
BEGIN
	/* Return a recordset of the controls in the given screen. */
	SELECT tableID, columnID, controlType,
		topCoord, leftCoord, height, width,	caption
	FROM [dbo].[ASRSysControls]
	WHERE screenID = @plngScreenID;
END