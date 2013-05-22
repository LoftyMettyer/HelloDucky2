CREATE PROCEDURE [dbo].[spASRWorkflowUsesInitiator]
(
	@piWorkflowID		integer,			
	@pfUsesInitiator		bit	OUTPUT
)
AS
BEGIN
	/* Return 1 if the given workflow uses the initiator's personnel record; else return 0 */
	DECLARE	@iCount	integer;

	SET @pfUsesInitiator = 0

	/* Initiator's record used by a Stored Data element action? */
	SELECT @iCount = COUNT(*)
	FROM ASRSysWorkflowElements
	WHERE ASRSysWorkflowElements.type = 5 -- 5 = Stored Data element
		AND (ASRSysWorkflowElements.dataRecord = 0 OR ASRSysWorkflowElements.secondaryDataRecord = 0) -- 0 = Initiator's record
		AND ASRSysWorkflowElements.workflowID = @piWorkflowID;

	IF @iCount > 0 SET @pfUsesInitiator = 1;

	IF @pfUsesInitiator = 0
	BEGIN
		/* Initiator's record used by a Stored Data element Database Value item? */
		SELECT @iCount = COUNT(*)
		FROM ASRSysWorkflowElementColumns
		INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowElementColumns.elementID = ASRSysWorkflowElements.ID
		WHERE ASRSysWorkflowElements.type = 5 -- 5 = Stored Data element
			AND ASRSysWorkflowElementColumns.valueType = 2 -- 2 = Database value
			AND ASRSysWorkflowElementColumns.dbRecord = 0 -- 0 = Initiator's record
			AND ASRSysWorkflowElements.workflowID = @piWorkflowID;
	
		IF @iCount > 0 SET @pfUsesInitiator = 1;
	END

	IF @pfUsesInitiator = 0
	BEGIN
		/* Initiator's record used by an Email element address? */
		SELECT @iCount = COUNT(*)
		FROM ASRSysWorkflowElements
		WHERE ASRSysWorkflowElements.type = 3 -- 3 = Email element
			AND ASRSysWorkflowElements.emailRecord = 0 -- 0 = Initiator's record
			AND ASRSysWorkflowElements.workflowID = @piWorkflowID;

		IF @iCount > 0 SET @pfUsesInitiator = 1;
	END

	IF @pfUsesInitiator = 0
	BEGIN
		/* Initiator's record used by an Email element Database Value item? */
		SELECT @iCount = COUNT(*)
		FROM ASRSysWorkflowElementItems
		INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowElementItems.elementID = ASRSysWorkflowElements.ID
		WHERE ASRSysWorkflowElements.type = 3 -- 3 = Email element
			AND ASRSysWorkflowElementItems.itemType = 1 -- 1 = Database value
			AND ASRSysWorkflowElementItems.dbRecord = 0 -- 0 = Initiator's record
			AND ASRSysWorkflowElements.workflowID = @piWorkflowID;

		IF @iCount > 0 SET @pfUsesInitiator = 1;
	END

	IF @pfUsesInitiator = 0
	BEGIN
		/* Initiator's record used by a Web Form element Database Value? */
		SELECT @iCount = COUNT(*)
		FROM ASRSysWorkflowElementItems
		INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowElementItems.elementID = ASRSysWorkflowElements.ID
		WHERE ASRSysWorkflowElements.type = 2 -- 2 = Web Form element
			AND ASRSysWorkflowElementItems.itemType = 1 -- 1 = Database value
			AND ASRSysWorkflowElementItems.dbRecord = 0 -- 0 = Initiator's record
			AND ASRSysWorkflowElements.workflowID = @piWorkflowID;

		IF @iCount > 0 SET @pfUsesInitiator = 1;
	END

	IF @pfUsesInitiator = 0
	BEGIN
		/* Initiator's record used by a Web Form element Record Selector? */
		SELECT @iCount = COUNT(*)
		FROM ASRSysWorkflowElementItems
		INNER JOIN ASRSysWorkflowElements ON ASRSysWorkflowElementItems.elementID = ASRSysWorkflowElements.ID
		WHERE ASRSysWorkflowElements.type = 2 -- 2 = Web Form element
			AND ASRSysWorkflowElementItems.itemType = 11 -- 11 = Record Selector
			AND ASRSysWorkflowElementItems.dbRecord = 0 -- 0 = Initiator's record
			AND ASRSysWorkflowElements.workflowID = @piWorkflowID;

		IF @iCount > 0 SET @pfUsesInitiator = 1;
	END

	/* Now Expressions by element */

	IF @pfUsesInitiator = 0
	BEGIN					
		/* Expressions - Decision Element */
		SELECT @iCount = COUNT(*) FROM ASRSysWorkflowElements WHERE TrueFlowExprID IN 
			(SELECT ExprID
			FROM ASRSysExprComponents EC
			WHERE EC.exprID in (SELECT E.exprID
					FROM ASRSysExpressions E
					WHERE E.utilityid = 114)
				AND EC.workflowRecord = 0); -- 0 = Initiator's record
				
		IF @iCount > 0 SET @pfUsesInitiator = 1;			
	END
	
	IF @pfUsesInitiator = 0
	BEGIN					
		/* Expressions - Web form (Descriptions) Element */
		SELECT @iCount = COUNT(*) FROM ASRSysWorkflowElements WHERE DescriptionExprID IN 
			(SELECT ExprID
			FROM ASRSysExprComponents EC
			WHERE EC.exprID in (SELECT E.exprID
					FROM ASRSysExpressions E
					WHERE E.utilityid = 114)
				AND EC.workflowRecord = 0); -- 0 = Initiator's record
				
		IF @iCount > 0 SET @pfUsesInitiator = 1;			
	END	
	
	IF @pfUsesInitiator = 0
	BEGIN					
		/* Expressions - Web form (Record Filters) Element */
		SELECT @iCount = COUNT(*) FROM ASRSysWorkflowElementItems WHERE RecordFilterID IN 
			(SELECT ExprID
			FROM ASRSysExprComponents EC
			WHERE EC.exprID in (SELECT E.exprID
					FROM ASRSysExpressions E
					WHERE E.utilityid = 114)
				AND EC.workflowRecord = 0); -- 0 = Initiator's record
				
		IF @iCount > 0 SET @pfUsesInitiator = 1;			
	END	
		
	IF @pfUsesInitiator = 0
	BEGIN					
		/* Expressions - Web form (Label Calculations & Default Calculations & Email) Element */
		SELECT @iCount = COUNT(*) FROM ASRSysWorkflowElementItems WHERE CalcID IN 
			(SELECT ExprID
			FROM ASRSysExprComponents EC
			WHERE EC.exprID in (SELECT E.exprID
					FROM ASRSysExpressions E
					WHERE E.utilityid = 114)
				AND EC.workflowRecord = 0); -- 0 = Initiator's record
				
		IF @iCount > 0 SET @pfUsesInitiator = 1;			
	END	
	
	IF @pfUsesInitiator = 0
	BEGIN					
		/* Expressions - Web form (Validation Calculations) Element */
		SELECT @iCount = COUNT(*) FROM ASRSysWorkflowElementValidations WHERE ExprID IN 
			(SELECT ExprID
			FROM ASRSysExprComponents EC
			WHERE EC.exprID in (SELECT E.exprID
					FROM ASRSysExpressions E
					WHERE E.utilityid = 114)
				AND EC.workflowRecord = 0); -- 0 = Initiator's record
				
		IF @iCount > 0 SET @pfUsesInitiator = 1;			
	END	
	
	IF @pfUsesInitiator = 0
	BEGIN					
		/* Expressions - Web form (Stored Data) Element */
		SELECT @iCount = COUNT(*) FROM ASRSysWorkflowElementColumns WHERE CalcID IN 
			(SELECT ExprID
			FROM ASRSysExprComponents EC
			WHERE EC.exprID in (SELECT E.exprID
					FROM ASRSysExpressions E
					WHERE E.utilityid = 114)
				AND EC.workflowRecord = 0); -- 0 = Initiator's record
				
		IF @iCount > 0 SET @pfUsesInitiator = 1;			
	END	
	
	
	
END