CREATE PROCEDURE [dbo].[sp_ASROp_IsLessThanOrEqualTo_1_1]
(
	@pfResult	bit OUTPUT,
	@psString1	varchar(MAX),
	@psString2	varchar(MAX)
)
AS
BEGIN
	IF @psString1 <= @psString2
	BEGIN
		SET @pfResult = 1;
	END
	ELSE
	BEGIN
		SET @pfResult = 0;
	END	
END