CREATE PROCEDURE [dbo].[sp_ASRFn_IfThenElse_3_1_1]
(
	@psResult   	varchar(MAX) OUTPUT,
	@pfTestValue	bit,
	@psString1		varchar(MAX),
	@psString2		varchar(MAX)
)
AS
BEGIN
	IF @pfTestValue = 1
	BEGIN
		SET @psResult = @psString1;
	END
	ELSE
	BEGIN
		SET @psResult = @psString2;
	END	
END