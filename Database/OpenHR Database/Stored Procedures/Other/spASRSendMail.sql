CREATE PROCEDURE [dbo].[spASRSendMail](
	@hResult int OUTPUT,
	@To varchar(MAX),
	@CC varchar(MAX),
	@BCC varchar(MAX),
	@Subject varchar(MAX),
	@Message varchar(MAX),
	@Attachment varchar(MAX))
AS
BEGIN
	EXEC @hResult = master..xp_sendmail
		@recipients=@To,
		@copy_recipients=@CC,
		@blind_copy_recipients=@BCC,
		@subject=@Subject,
		@message=@Message,
		@attachments=@Attachment;
END