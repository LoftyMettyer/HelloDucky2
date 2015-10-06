CREATE PROCEDURE [dbo].[spASRIntMailMergeUploadTemplate]
(
	@MailMergeId	int = 0,
	@Template		image = null,
	@TemplateName	nvarchar(255) = ''
)
AS
BEGIN

	DELETE FROM ASRSysMailMergeTemplate WHERE MailMergeID = @MailMergeId;

	INSERT dbo.ASRSysMailMergeTemplate (MailMergeID, Template, TemplateName, UploadDate, UploadedUser)
		VALUES (@MailMergeId, @Template, @TemplateName, GETDATE(), SYSTEM_USER);
END
