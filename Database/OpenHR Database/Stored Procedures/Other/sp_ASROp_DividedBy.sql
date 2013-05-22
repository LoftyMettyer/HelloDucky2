
CREATE PROCEDURE sp_ASROp_DividedBy 
(
	@pdblResult	float OUTPUT,
	@pdblFirst 	float,
	@pdblSecond	float
)
AS
BEGIN
	IF @pdblSecond <> 0
	BEGIN
		SET @pdblResult = @pdblFirst / @pdblSecond	
	END
	ELSE
	BEGIN
		SET @pdblResult = 0
	END
END




GO

