CREATE PROCEDURE [dbo].[sp_ASRIntGetExprColumns] (
	@piTableID			integer,
	@piComponentType	integer,
	@piNumericsOnly		integer
)
AS
BEGIN
	/* Return a recordset of tab-delimted column definitions ;
	<column id><tab><column name><tab><data type> */
	DECLARE @iDataType	integer;

	IF @piComponentType = 1
	BEGIN
		SET @iDataType = -3;
	END
	ELSE
	BEGIN
		SET @iDataType = -7;
		SET @piNumericsOnly = 0;
	END

	SELECT 
		convert(varchar(255), columnID) + char(9) +
		columnName + char(9) +
		convert(varchar(255), dataType) AS [definitionString]
	FROM [dbo].[ASRSysColumns]
	WHERE tableID = @piTableID
		AND dataType <> -4
		AND dataType <> -3
		AND dataType <> @iDataType
		AND columnType <> 4
		AND columnType <> 3
		AND ((@piNumericsOnly = 0) OR (dataType = 2) OR (dataType = 4))
	ORDER BY columnName;
END