CREATE PROCEDURE [dbo].[spASREmailImmediate]
	(@Username varchar(255))
AS
	BEGIN
		DECLARE @QueueID int,
				@LinkID int,
				@RecordID int,
				@sSQL nvarchar(max),
				@DateDue datetime,
				@hResult int,
				@TableID int,
				@tmpUser varchar(255),
				@RecalculateRecordDesc int,
				@AttachmentFolder varchar(max)

		DECLARE @TempText nvarchar(max),
				@To varchar(max),
				@CC varchar(max),
				@BCC varchar(max),
				@Subject varchar(max),
				@MsgText varchar(max),
				@Attachment varchar(max)

		SET @AttachmentFolder = ''
		SELECT @AttachmentFolder = settingvalue
		FROM asrsyssystemsettings
		WHERE [section] = 'email' and [settingkey] = 'attachment path'

		DECLARE emailqueue_cursor
		CURSOR LOCAL FAST_FORWARD FOR 
		  SELECT ASRSysEmailQueue.QueueID
			   , ASRSysEmailQueue.LinkID
			   , ASRSysEmailQueue.RecordID
			   , ASRSysEmailQueue.TableID
			   , ASRSysEmailQueue.DateDue
			   , ASRSysEmailQueue.UserName
			   , ASRSysEmailQueue.RecalculateRecordDesc
			   , ASRSysEmailQueue.RepTo
               , ASRSysEmailQueue.RepCC
               , ASRSysEmailQueue.RepBCC
               , ASRSysEmailQueue.[Subject]
               , ASRSysEmailQueue.MsgText
               , ASRSysEmailQueue.Attachment
          FROM ASRSysEmailQueue
		  LEFT OUTER JOIN
				 ASRSysEmailLinks
			  ON ASRSysEmailLinks.LinkID = ASRSysEmailQueue.LinkID
		  WHERE  DateSent IS Null
			AND  datediff(dd,DateDue,getdate()) >= 0
			AND  (LOWER(substring(@Username,charindex('\',@Username)+1,999)) = LOWER(substring([Username],charindex('\',[Username])+1,999))
				  OR @Username = ''
				  )
		  ORDER BY dateDue

		OPEN emailqueue_cursor
		FETCH NEXT FROM emailqueue_cursor INTO @QueueID, @LinkID, @RecordID, @TableID, @DateDue, @tmpUser, @RecalculateRecordDesc, @To, @CC, @BCC, @Subject, @MsgText, @Attachment

		WHILE (@@fetch_status = 0)
		BEGIN

			SET @hResult = 0
			IF @RecalculateRecordDesc = 1 OR rtrim(isnull(@To,'')) = ''
			BEGIN
				SELECT @sSQL = 'spASREmail_' + convert(varchar,@LinkID)
				IF EXISTS (SELECT * FROM sysobjects WHERE type = 'P' AND name = @sSQL)
				BEGIN
					EXEC @hResult = @sSQL @queueid, @recordid, @tmpUser, @To OUTPUT, @CC OUTPUT, @BCC OUTPUT, @Subject OUTPUT, @MsgText OUTPUT, @Attachment OUTPUT
				END
			END

			IF @hResult = 0 AND RTrim(@To) <> ''
			BEGIN
				IF @Attachment <> '' SET @Attachment = @AttachmentFolder+@Attachment
				EXEC spASRSendMail @hResult OUTPUT, @To, @CC, @BCC, @Subject, @MsgText, @Attachment
				IF @hResult = 0
					UPDATE ASRSysEmailQueue
					SET DateSent = getdate(), RecalculateRecordDesc = 0
					WHERE QueueID = @QueueID
			END

			FETCH NEXT FROM emailqueue_cursor INTO @QueueID, @LinkID, @RecordID, @TableID, @DateDue, @tmpUser, @RecalculateRecordDesc, @To, @CC, @BCC, @Subject, @MsgText, @Attachment
		END

		CLOSE emailqueue_cursor
		DEALLOCATE emailqueue_cursor

	END
