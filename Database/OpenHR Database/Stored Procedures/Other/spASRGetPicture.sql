CREATE PROCEDURE dbo.spASRGetPicture
(
	@piPictureID		integer
)
AS
BEGIN
	SELECT TOP 1 name, picture
	FROM ASRSysPictures
	WHERE pictureID = @piPictureID
END
GO

