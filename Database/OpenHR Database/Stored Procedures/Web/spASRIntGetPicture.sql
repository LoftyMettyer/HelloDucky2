CREATE PROCEDURE [dbo].[spASRIntGetPicture]
	(
		@piPictureID		integer
	)
	AS
	BEGIN
		SET NOCOUNT ON
		SELECT TOP 1 name, picture
		FROM ASRSysPictures
		WHERE pictureID = @piPictureID
	END