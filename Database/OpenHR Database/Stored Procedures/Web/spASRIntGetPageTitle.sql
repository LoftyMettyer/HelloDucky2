CREATE PROCEDURE spASRIntGetPageTitle (
	@piTableID		integer,
	@piViewID		integer,
	@psPageTitle	varchar(200) 	OUTPUT
)
AS
BEGIN

	SET NOCOUNT ON;

	SELECT @psPageTitle = PageTitle
	FROM ASRSysSSIViews
	WHERE (TableID = @piTableID) AND  (ViewID = @piViewID);

END
GO

