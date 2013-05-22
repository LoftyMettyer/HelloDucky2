CREATE PROCEDURE spadmin_writepicture(@guid uniqueidentifier, @name varchar(255), @pictureID integer OUTPUT, @pictureHex varbinary(MAX))
	AS
	BEGIN

		IF NOT EXISTS(SELECT [guid] FROM dbo.[ASRSysPictures] WHERE [guid] = @guid)	
		BEGIN

			SELECT @pictureID = ISNULL(MAX(PictureID), 0) + 1 FROM dbo.[ASRSysPictures];

			INSERT [ASRSysPictures] (PictureID, Name, PictureType, [guid], [Picture]) 
				SELECT @pictureID, @name, 1, @guid, @pictureHex;

		END
		ELSE
		BEGIN
			SELECT @pictureID = [PictureID] FROM dbo.[ASRSysPictures] WHERE [guid] = @guid;
			UPDATE [ASRSysPictures] SET [Name] = @name, Picture = @pictureHex WHERE [guid] = @guid;
		END

	END