CREATE PROCEDURE [dbo].[sp_ASRFn_FirstNameFromForenames]
(
	@psResult		varchar(MAX) OUTPUT,
	@psForenames	varchar(MAX)
)
AS
BEGIN
	IF (LEN(@psForenames) = 0) OR (@psForenames IS NULL)
	BEGIN
		SET @psResult = '';
	END
	ELSE
	BEGIN
		IF CHARINDEX(' ', @psForenames) > 0
		BEGIN
			SET @psResult = RTRIM(LTRIM(LEFT(@psForenames, CHARINDEX(' ', @psForenames))));
		END
		ELSE
		BEGIN
			SET @psResult = RTRIM(LTRIM(@psForenames));
		END
	END
END
