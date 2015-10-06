CREATE PROCEDURE [dbo].[spASRIntMailMergeDownloadTemplate]
(
	@MailMergeId int = 0
)
AS
BEGIN
	SELECT TOP 1 Template, TemplateName FROM dbo.ASRSysMailMergeTemplate
		WHERE MailMergeID = @MailMergeId
		ORDER BY id DESC;
END
