CREATE  PROCEDURE [dbo].[sp_ASR_TBCheckPreRequisites] (
		@plngCourseRecordID int,
		@plngEmployeeRecordID int,
		@piPreReqsMet int OUTPUT
)
AS
BEGIN

	SET @piPreReqsMet = 0;

END