CREATE PROCEDURE [dbo].[sp_ASRFn_IsValidForPayrollCharset]
		(
			@result integer OUTPUT,
			@input varchar(MAX),
			@Charset varchar(1)
		)
		AS
		BEGIN

			--Charset A - typically Address
			--Charset C - typically Forename
			--Charset D - typically Surname

			DECLARE @ValidCharacters varchar(MAX);
			DECLARE @Index int;


			IF      @Charset = 'A' SET @ValidCharacters = 'abcdefghijhklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ-''0123456789,&/(). =!"%&*;<>+:?'
			ELSE IF @Charset = 'B' SET @ValidCharacters = 'abcdefghijhklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 '
			ELSE IF @Charset = 'C' SET @ValidCharacters = 'abcdefghijhklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ-'''
			ELSE IF @Charset = 'D' SET @ValidCharacters = 'abcdefghijhklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ-''0123456789,&/(). '
			ELSE IF @Charset = 'G' SET @ValidCharacters = 'abcdefghijhklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ-''0123456789,&/(). =!"%&*;<>+:?'
			ELSE IF @Charset = 'H' SET @ValidCharacters = 'abcdefghijhklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ-''. '
			
			SET @result = 1;
			SET @Index = 1;
			WHILE (@Index <= datalength(@input) AND @result = 1)
			BEGIN
				IF charindex(substring(@input,@Index,1),@ValidCharacters) = 0
					SET @result = 0;
				SET @Index = @Index + 1;
			END	

		END