CREATE PROCEDURE [dbo].[spASRWorkflowValidRecord]
	@piInstanceID				integer,
	@piRecordType				integer,
	@piRecordID					integer,
	@sElementIdentifier			varchar(MAX),
	@sElementItemIdentifier		varchar(MAX),
	@pfValid					bit		OUTPUT
AS
BEGIN
	DECLARE
		@iTableID				integer,
		@iWorkflowID			integer,
		@iElementType			integer;

	SET @pfValid = 0;

	SELECT @iWorkflowID = WF.ID,
		@iTableID = 
			CASE
				WHEN @piRecordType = 4 THEN isnull(WF.baseTable, 0)
				ELSE 0
			END
	FROM ASRSysWorkflows WF
	INNER JOIN ASRSysWorkflowInstances WFI ON WF.ID = WFI.workflowID
		AND WFI.ID = @piInstanceID;

	IF @piRecordType = 0
	BEGIN
		-- Initiator's record
		SELECT @iTableID = convert(integer, ISNULL(parameterValue, '0'))
		FROM ASRSysModuleSetup
		WHERE moduleKey = 'MODULE_PERSONNEL'
			AND parameterKey = 'Param_TablePersonnel';

		IF @iTableID = 0
		BEGIN
			SELECT @iTableID = convert(integer, isnull(parameterValue, 0))
			FROM ASRSysModuleSetup
			WHERE moduleKey = 'MODULE_WORKFLOW'
			AND parameterKey = 'Param_TablePersonnel';
		END
	END

	IF @piRecordType = 1
	BEGIN
		-- Identified record
		SELECT @iElementType = ASRSysWorkflowElements.type,
			@iTableID = 
				CASE
					WHEN ASRSysWorkflowElements.type = 5 THEN isnull(ASRSysWorkflowElements.dataTableID, 0)
					ELSE 0
				END
		FROM ASRSysWorkflowElements
		WHERE ASRSysWorkflowElements.workflowID = @iWorkflowID
			AND upper(rtrim(ltrim(ASRSysWorkflowElements.identifier))) = upper(rtrim(ltrim(@sElementIdentifier)));

		IF @iElementType = 2
		BEGIN
			 -- WebForm
			SELECT @iTableID = WFEI.tableID
			FROM ASRSysWorkflowElementItems WFEI
			INNER JOIN ASRSysWorkflowElements WFE ON WFEI.elementID = WFE.ID
				AND WFE.identifier = @sElementIdentifier
				AND WFE.workflowID = @iWorkflowID
			WHERE WFEI.identifier = @sElementItemIdentifier;
		END
	END

	EXEC [dbo].[spASRWorkflowValidTableRecord]
		@iTableID,
		@piRecordID,
		@pfValid	OUTPUT;
END