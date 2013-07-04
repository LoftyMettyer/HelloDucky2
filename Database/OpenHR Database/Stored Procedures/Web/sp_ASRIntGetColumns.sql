CREATE PROCEDURE [dbo].[sp_ASRIntGetColumns] (
	@piTableID	integer
)
AS
BEGIN

	SET NOCOUNT ON;

	SELECT ColumnID, ColumnName, OLEType
		FROM [dbo].[ASRSysColumns]
		WHERE tableID = @piTableID AND NOT(ColumnName = 'ID')
		ORDER BY ColumnName;
END