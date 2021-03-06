CREATE PROCEDURE [dbo].[sp_ASRFn_ReplaceCharsInString]
		(
			@psResult		varchar(MAX) OUTPUT,
			@input varchar(MAX),
			@searchstring varchar(MAX),
			@replacestring varchar(MAX)
		)
		AS
		BEGIN

			IF ISNULL(@input, '') = '' RETURN;			
			
			SET @psResult = REPLACE(@input, @searchstring, @replacestring);

		END