CREATE PROCEDURE [dbo].[spASRIntGetColumnTableID] (
	@piColumnID	integer,
	@piTableID	integer OUTPUT
)
AS
BEGIN
	
	SET NOCOUNT ON;

	SELECT @piTableID = tableID
	FROM ASRSysColumns
	WHERE columnID = @piColumnID;

	IF @piTableID IS null SET @piTableID = 0;
END