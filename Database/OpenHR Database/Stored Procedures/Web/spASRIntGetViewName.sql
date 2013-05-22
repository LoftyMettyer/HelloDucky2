CREATE PROCEDURE [dbo].[spASRIntGetViewName] (
	@piViewID	integer,
	@psViewName	varchar(255) 	OUTPUT
)
AS
BEGIN
	SELECT @psViewName = viewName
	FROM [dbo].[ASRSysViews]
	WHERE viewID = @piViewID;
END