CREATE PROCEDURE [dbo].[sp_ASROp_IsContainedWithin]
(
	@pfResult   		bit OUTPUT,
	@psSearchString 	varchar(MAX),
	@psWholeString   	varchar(MAX)
)
AS
BEGIN
	DECLARE @iTemp integer;

	SET @iTemp = charindex(@psSearchString, @psWholeString);

	IF @iTemp > 0
	BEGIN
		SET @pfResult = 1;
	END
	ELSE
	BEGIN
		SET @pfResult = 0;
	END
END